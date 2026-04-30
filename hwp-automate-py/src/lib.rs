//! `hwp_automate` — Python 바인딩 for HWP 자동화 (rhwp 기반).
//!
//! 핵심 패턴: 사용자가 제공한 양식(.hwp)을 로드해 빈 셀에 값을 채우는 것.
//! from-scratch 로 새 문서를 만드는 함수는 의도적으로 노출하지 않는다.
//!
//! 노출 함수:
//!   - analyze_template(path: str) -> dict
//!         양식의 표/스타일/번호 인벤토리 (read-only).
//!   - fill_template(template_path: str, out_path: str, operations: list[dict]) -> dict
//!         여러 표·여러 컬럼·여러 셀을 한 번에 채움.  Pre-flight 검증 후 batch 적용.
//!   - fill_template_table(template_path: str, out_path: str, mapping: dict) -> dict
//!         단일 표·단일 컬럼 편의 함수 — 내부적으로 fill_template 한 번 호출.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use rhwp::document_core::DocumentCore;
use rhwp::error::HwpError;
use rhwp::model::control::Control;
use rhwp::parser::cfb_reader::LenientCfbReader;
use rhwp::serializer::mini_cfb;
use std::fs;

fn map_err<T>(r: Result<T, HwpError>) -> PyResult<T> {
    r.map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("HwpError: {}", e)))
}

fn io<T>(r: std::io::Result<T>) -> PyResult<T> {
    r.map_err(|e| pyo3::exceptions::PyIOError::new_err(e.to_string()))
}

/// HWP 표준 layout 에 따라 leaf 이름 → 전체 storage path 추론.
///
/// LenientCfbReader 가 leaf 이름만 저장하므로 (예: "BIN0001.bmp", "Section0") parent storage 를
/// 모른다. HWP 5.0 의 CFB 구조는 표준화되어 있어 leaf 이름 패턴으로 storage 결정 가능.
fn leaf_to_hwp_path(leaf: &str) -> String {
    let l = leaf.trim_start_matches('/');
    // BinData: BIN****.* 형식 (BIN0001.bmp, BIN001F.png 등)
    if l.starts_with("BIN") && l.contains('.') && l.len() >= 8 {
        return format!("/BinData/{}", l);
    }
    // Preview: PrvText, PrvImage
    if l == "PrvText" || l == "PrvImage" {
        return format!("/Preview/{}", l);
    }
    // BodyText: Section{N}
    if l.starts_with("Section") && l[7..].chars().all(|c| c.is_ascii_digit()) {
        return format!("/BodyText/{}", l);
    }
    // ViewText: ViewSection{N}
    if l.starts_with("ViewSection") && l[11..].chars().all(|c| c.is_ascii_digit()) {
        return format!("/ViewText/{}", l);
    }
    // Scripts: DefaultJScript, JScriptVersion
    if l == "DefaultJScript" || l == "JScriptVersion" {
        return format!("/Scripts/{}", l);
    }
    // 그 외 root level: DocInfo, FileHeader, HwpSummaryInformation, _LinkDoc 등
    format!("/{}", l)
}

/// 두 HWP CFB 를 머지하여 새 CFB 를 만든다 — **rhwp output 베이스 + BinData/Preview 만 input 에서 보존**.
///
/// 동작:
///   - `rhwp_output_bytes` 의 모든 stream 을 베이스로 사용 (rhwp 의 valid CFB)
///   - 단, BinData/* 와 Preview/* 는 `input_bytes` 의 raw bytes 로 교체 (rhwp 라운드트립 손실 회피)
///   - rhwp output 에 없지만 input 에 있는 BinData/Preview stream 도 추가 (rhwp 가 빠뜨렸을 수 있는 이미지)
///
/// 이 로직이 정답인 이유:
///   - rhwp output 자체는 cell text 변경이 반영된 valid HWP CFB (rhwp 의 891+ tests 가 보장)
///   - BinData/Preview 만 raw bytes 그대로 교체하므로 layout/압축 일관성 유지
///   - rhwp 가 BinData 를 약간 변형 또는 빠뜨려도 input 원본으로 복원
///
/// 반환: (final_bytes, count_from_rhwp, count_from_input)
fn merge_cfb_preserving_input(
    input_bytes: &[u8],
    rhwp_output_bytes: &[u8],
) -> Result<(Vec<u8>, usize, usize), String> {
    let input_reader = LenientCfbReader::open(input_bytes)
        .map_err(|e| format!("input CFB 파싱 실패: {:?}", e))?;
    let rhwp_reader = LenientCfbReader::open(rhwp_output_bytes)
        .map_err(|e| format!("rhwp output CFB 파싱 실패: {:?}", e))?;

    // input 으로부터 raw bytes 그대로 가져올 stream 패턴
    let take_from_input = |path: &str| -> bool {
        let p = path.strip_prefix('/').unwrap_or(path);
        p.starts_with("BinData/") || p == "PrvText" || p == "PrvImage" || p.starts_with("Preview")
    };

    let mut named_owned: Vec<(String, Vec<u8>)> = Vec::new();
    let mut from_rhwp = 0usize;
    let mut from_input = 0usize;
    let mut added_paths: std::collections::HashSet<String> = std::collections::HashSet::new();

    // take_from_input 도 leaf name 기준으로 검사 — BIN****.* 패턴 매칭
    let take_from_input_leaf = |leaf: &str| -> bool {
        let l = leaf.trim_start_matches('/');
        (l.starts_with("BIN") && l.contains('.'))
            || l == "PrvText"
            || l == "PrvImage"
    };

    // 1단계: rhwp output 을 베이스로 모든 stream 처리
    for (path, _start, _size, obj_type) in rhwp_reader.list_entries() {
        if *obj_type != 2 {
            continue;
        }
        if path == "Root Entry" || path.is_empty() {
            continue;
        }
        // HWP 표준 layout 으로 path 재구성 (leaf → /storage/leaf)
        let canonical = leaf_to_hwp_path(path);

        let bytes_opt = if take_from_input_leaf(path) && input_reader.has_stream(path) {
            from_input += 1;
            input_reader.read_stream(path).ok()
        } else {
            from_rhwp += 1;
            rhwp_reader.read_stream(path).ok()
        };
        if let Some(bytes) = bytes_opt {
            added_paths.insert(canonical.clone());
            named_owned.push((canonical, bytes));
        }
    }

    // 2단계: input 에만 있는 BinData/Preview stream (rhwp 가 빠뜨린 것) 추가
    for (path, _start, _size, obj_type) in input_reader.list_entries() {
        if *obj_type != 2 {
            continue;
        }
        if path == "Root Entry" || path.is_empty() {
            continue;
        }
        if !take_from_input_leaf(path) {
            continue;
        }
        let canonical = leaf_to_hwp_path(path);
        if added_paths.contains(&canonical) {
            continue;
        }
        if let Ok(bytes) = input_reader.read_stream(path) {
            from_input += 1;
            named_owned.push((canonical, bytes));
        }
    }

    // mini_cfb::build_cfb 입력 형식 변환
    let refs: Vec<(&str, &[u8])> = named_owned
        .iter()
        .map(|(p, b)| (p.as_str(), b.as_slice()))
        .collect();
    let final_bytes = mini_cfb::build_cfb(&refs)?;
    Ok((final_bytes, from_rhwp, from_input))
}

/// 단독 호출용 — `source_hwp` 의 모든 BinData/Preview stream 등을 보존하면서
/// `target_hwp` 의 BodyText/DocInfo/FileHeader 만 살린 새 CFB 로 `out_hwp` 를 만든다.
///
/// fill_template 의 preserve_images=True 와 동일 효과를 별도 단계로 적용 가능.
#[pyfunction]
#[pyo3(signature = (source_hwp, target_hwp, out_hwp))]
fn preserve_images_from_source(
    source_hwp: &str,
    target_hwp: &str,
    out_hwp: &str,
) -> PyResult<usize> {
    let source = io(fs::read(source_hwp))?;
    let target = io(fs::read(target_hwp))?;
    let (merged, from_rhwp, _from_input) = merge_cfb_preserving_input(&source, &target)
        .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("CFB 머지 실패: {}", e)))?;
    if let Some(parent) = std::path::Path::new(out_hwp).parent() {
        if !parent.as_os_str().is_empty() {
            io(fs::create_dir_all(parent))?;
        }
    }
    io(fs::write(out_hwp, &merged))?;
    Ok(from_rhwp)
}

/// 표 위치 + 헤더 행 캐시.
#[derive(Debug, Clone)]
struct TableLocation {
    section: usize,
    parent_para: usize,
    control: usize,
    rows: u16,
    cols: u16,
    /// (row, col, joined_text) — 셀 텍스트 캐시 (다문단은 "|" 결합)
    cells: Vec<(u16, u16, String)>,
}

/// 적용할 셀 채우기 단위.
///
/// `cell_idx` 는 rhwp 의 `Table.cells` Vec 안에서의 0-based 위치.
/// 병합된 셀이 있는 표에서는 row × cols + col 공식이 틀리므로 위치 기반 검색으로 미리 산출.
#[derive(Debug, Clone)]
struct CellFill {
    row: u16,
    col: u16,
    cell_idx: usize,
    value: String,
}

/// 표의 cells 벡터에서 (row, col) 위치를 가진 셀의 인덱스를 찾는다.
/// 병합으로 그 위치가 표에 존재하지 않으면 None.
fn find_cell_idx(table: &TableLocation, row: u16, col: u16) -> Option<usize> {
    table.cells.iter().position(|(r, c, _)| *r == row && *c == col)
}

/// 빈 셀 (row, col) 의 의미를 추론하기 위해 인접 텍스트 셀을 라벨로 추정.
///
/// 한국 표 양식의 통상 패턴:
///   1순위: 같은 행의 왼쪽 셀 (라벨-값이 가로 페어, 예: "기업명 | (값)")
///   2순위: 같은 열의 위쪽 셀 (헤더 행 또는 위 라벨, 예: 컬럼 헤더 + 값)
///
/// AI 가 양식 의미를 추론하는 데 가장 큰 단서. 추론 실패 시 None.
fn find_neighbor_label(cells: &[(u16, u16, String)], row: u16, col: u16) -> Option<String> {
    // 1순위: 같은 행 왼쪽 — 가장 가까운 비지 않은 셀
    if col > 0 {
        for c in (0..col).rev() {
            if let Some((_, _, txt)) = cells.iter().find(|(r, cc, _)| *r == row && *cc == c) {
                let trimmed = txt.trim();
                if !trimmed.is_empty() {
                    return Some(trimmed.to_string());
                }
            }
        }
    }
    // 2순위: 같은 열 위쪽 — 가장 가까운 비지 않은 셀
    if row > 0 {
        for r in (0..row).rev() {
            if let Some((_, _, txt)) = cells.iter().find(|(rr, c, _)| *rr == r && *c == col) {
                let trimmed = txt.trim();
                if !trimmed.is_empty() {
                    return Some(trimmed.to_string());
                }
            }
        }
    }
    None
}

/// 양식의 모든 표 인벤토리.
fn discover_all_tables(core: &DocumentCore) -> Vec<TableLocation> {
    let mut out = Vec::new();
    let doc = core.document();
    for (sec_idx, section) in doc.sections.iter().enumerate() {
        for (para_idx, para) in section.paragraphs.iter().enumerate() {
            for (ctrl_idx, ctrl) in para.controls.iter().enumerate() {
                if let Control::Table(t) = ctrl {
                    let cells: Vec<(u16, u16, String)> = t
                        .cells
                        .iter()
                        .map(|c| {
                            let txt = c
                                .paragraphs
                                .iter()
                                .map(|p| p.text.as_str())
                                .collect::<Vec<_>>()
                                .join("|");
                            (c.row, c.col, txt)
                        })
                        .collect();
                    out.push(TableLocation {
                        section: sec_idx,
                        parent_para: para_idx,
                        control: ctrl_idx,
                        rows: t.row_count,
                        cols: t.col_count,
                        cells,
                    });
                }
            }
        }
    }
    out
}

/// operation dict 에서 표 위치 식별 — `header_match` 우선, 없으면 `table_at: [sec, para, ctrl]`.
fn resolve_table<'a>(
    op: &Bound<'_, PyDict>,
    tables: &'a [TableLocation],
    op_idx: usize,
) -> PyResult<&'a TableLocation> {
    if let Some(hm) = op.get_item("header_match")? {
        let header_match: String = hm.extract()?;
        return tables
            .iter()
            .find(|t| {
                t.cells
                    .iter()
                    .any(|(r, _, txt)| *r == 0 && txt.contains(&header_match))
            })
            .ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].header_match='{}' 으로 표를 찾을 수 없음",
                    op_idx, header_match
                ))
            });
    }
    if let Some(ta) = op.get_item("table_at")? {
        let coords: Vec<usize> = ta.extract()?;
        if coords.len() != 3 {
            return Err(pyo3::exceptions::PyValueError::new_err(format!(
                "operations[{}].table_at 은 [sec, para, ctrl] 3원소 list",
                op_idx
            )));
        }
        let (sec, para, ctrl) = (coords[0], coords[1], coords[2]);
        return tables
            .iter()
            .find(|t| t.section == sec && t.parent_para == para && t.control == ctrl)
            .ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].table_at=[{}, {}, {}] 위치에 표 없음",
                    op_idx, sec, para, ctrl
                ))
            });
    }
    Err(pyo3::exceptions::PyKeyError::new_err(format!(
        "operations[{}] 에 'header_match' 또는 'table_at' 중 하나는 필수",
        op_idx
    )))
}

/// 단일 operation 을 CellFill list 로 변환 (pre-flight 검증 포함).
fn op_to_fills(
    op: &Bound<'_, PyDict>,
    table: &TableLocation,
    op_idx: usize,
) -> PyResult<Vec<CellFill>> {
    let mut fills: Vec<CellFill> = Vec::new();

    // Mode A: column_by_header — `column` + `values: {row: value}`
    if let Some(col_obj) = op.get_item("column")? {
        let column: String = col_obj.extract()?;
        let target_col = table
            .cells
            .iter()
            .find(|(r, _, txt)| *r == 0 && *txt == column)
            .map(|(_, c, _)| *c)
            .ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].column='{}' 헤더를 찾을 수 없음 (대상 표 cols={})",
                    op_idx, column, table.cols
                ))
            })?;

        let values: Bound<'_, PyDict> = op
            .get_item("values")?
            .ok_or_else(|| {
                pyo3::exceptions::PyKeyError::new_err(format!(
                    "operations[{}] 에 'column' 이 있으면 'values' 도 필요",
                    op_idx
                ))
            })?
            .downcast_into()
            .map_err(|_| {
                pyo3::exceptions::PyTypeError::new_err(format!(
                    "operations[{}].values 는 dict {{row: value}}",
                    op_idx
                ))
            })?;

        for (row_obj, val_obj) in values.iter() {
            let row: u32 = row_obj.extract()?;
            let value: String = val_obj.extract()?;
            if row >= table.rows as u32 {
                return Err(pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].values 의 row={} 가 표 범위 초과 (rows={})",
                    op_idx, row, table.rows
                )));
            }
            let r = row as u16;
            let cell_idx = find_cell_idx(table, r, target_col).ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].values: (row={}, col={}) 위치에 셀이 없음 (병합되었거나 범위 외)",
                    op_idx, r, target_col
                ))
            })?;
            fills.push(CellFill {
                row: r,
                col: target_col,
                cell_idx,
                value,
            });
        }
        return Ok(fills);
    }

    // Mode B: cells direct — `cells: [{row, col, value}, ...]`
    if let Some(cells_obj) = op.get_item("cells")? {
        let cells: Bound<'_, PyList> = cells_obj.downcast_into().map_err(|_| {
            pyo3::exceptions::PyTypeError::new_err(format!(
                "operations[{}].cells 는 list[{{row, col, value}}]",
                op_idx
            ))
        })?;
        for (i, cell_obj) in cells.iter().enumerate() {
            let cell: Bound<'_, PyDict> = cell_obj.downcast_into().map_err(|_| {
                pyo3::exceptions::PyTypeError::new_err(format!(
                    "operations[{}].cells[{}] 는 {{row, col, value}} dict",
                    op_idx, i
                ))
            })?;
            let row: u32 = cell
                .get_item("row")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyKeyError::new_err(format!(
                        "operations[{}].cells[{}].row 누락",
                        op_idx, i
                    ))
                })?
                .extract()?;
            let col: u32 = cell
                .get_item("col")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyKeyError::new_err(format!(
                        "operations[{}].cells[{}].col 누락",
                        op_idx, i
                    ))
                })?
                .extract()?;
            let value: String = cell
                .get_item("value")?
                .ok_or_else(|| {
                    pyo3::exceptions::PyKeyError::new_err(format!(
                        "operations[{}].cells[{}].value 누락",
                        op_idx, i
                    ))
                })?
                .extract()?;
            if row >= table.rows as u32 {
                return Err(pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].cells[{}].row={} 가 표 범위 초과 (rows={})",
                    op_idx, i, row, table.rows
                )));
            }
            if col >= table.cols as u32 {
                return Err(pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].cells[{}].col={} 가 표 범위 초과 (cols={})",
                    op_idx, i, col, table.cols
                )));
            }
            let r = row as u16;
            let c = col as u16;
            let cell_idx = find_cell_idx(table, r, c).ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!(
                    "operations[{}].cells[{}]: (row={}, col={}) 위치에 셀이 없음 (병합되었거나 표 외)",
                    op_idx, i, r, c
                ))
            })?;
            fills.push(CellFill {
                row: r,
                col: c,
                cell_idx,
                value,
            });
        }
        return Ok(fills);
    }

    Err(pyo3::exceptions::PyKeyError::new_err(format!(
        "operations[{}] 에 'column' 또는 'cells' 중 하나는 필수",
        op_idx
    )))
}

/// 양식의 표/스타일/번호 인벤토리 (read-only).
#[pyfunction]
fn analyze_template<'py>(py: Python<'py>, path: &str) -> PyResult<Bound<'py, PyDict>> {
    let bytes = io(fs::read(path))?;
    let core = map_err(DocumentCore::from_bytes(&bytes))?;
    let doc = core.document();
    let di = &doc.doc_info;

    let result = PyDict::new_bound(py);
    result.set_item("path", path)?;
    result.set_item("file_size", bytes.len())?;
    result.set_item("section_count", doc.sections.len())?;
    result.set_item(
        "paragraph_count",
        doc.sections.iter().map(|s| s.paragraphs.len()).sum::<usize>(),
    )?;
    result.set_item("char_shape_count", di.char_shapes.len())?;
    result.set_item("para_shape_count", di.para_shapes.len())?;
    result.set_item("style_count", di.styles.len())?;
    result.set_item("numbering_count", di.numberings.len())?;
    result.set_item("bullet_count", di.bullets.len())?;
    result.set_item("border_fill_count", di.border_fills.len())?;
    let style_names: Vec<&str> = di.styles.iter().map(|s| s.local_name.as_str()).collect();
    result.set_item("style_names", style_names)?;

    let tables_out = PyList::empty_bound(py);
    let tables = discover_all_tables(&core);
    for t in tables {
        let entry = PyDict::new_bound(py);
        entry.set_item("section", t.section)?;
        entry.set_item("parent_para", t.parent_para)?;
        entry.set_item("control", t.control)?;
        entry.set_item("rows", t.rows)?;
        entry.set_item("cols", t.cols)?;

        // 헤더 행 (row 0) 텍스트 — 표의 의미 식별자
        let header: Vec<&str> = t
            .cells
            .iter()
            .filter(|(r, _, _)| *r == 0)
            .map(|(_, _, txt)| txt.as_str())
            .collect();
        entry.set_item("header", header)?;

        // 모든 셀 + 빈 셀 + suggested_fields 동시 빌드 (AI 가 의미 추론하는 핵심 단서)
        let cells_list = PyList::empty_bound(py);
        let empty_cells_list = PyList::empty_bound(py);
        let suggested_fields = PyList::empty_bound(py);

        for (r, c, txt) in &t.cells {
            let is_empty = txt.trim().is_empty();
            let neighbor_label = if is_empty {
                find_neighbor_label(&t.cells, *r, *c)
            } else {
                None
            };

            // cells (전체)
            let cell_dict = PyDict::new_bound(py);
            cell_dict.set_item("row", *r)?;
            cell_dict.set_item("col", *c)?;
            cell_dict.set_item("text", txt)?;
            cell_dict.set_item("is_empty", is_empty)?;
            if let Some(lbl) = &neighbor_label {
                cell_dict.set_item("neighbor_label", lbl)?;
            }
            cells_list.append(cell_dict)?;

            // empty_cells (빠른 lookup)
            if is_empty {
                let ec = PyDict::new_bound(py);
                ec.set_item("row", *r)?;
                ec.set_item("col", *c)?;
                if let Some(lbl) = &neighbor_label {
                    ec.set_item("neighbor_label", lbl)?;
                }
                empty_cells_list.append(ec)?;

                // suggested_fields: 라벨 추론 성공한 빈 셀만 포함
                if let Some(lbl) = &neighbor_label {
                    let sf = PyDict::new_bound(py);
                    sf.set_item("label", lbl)?;
                    sf.set_item("row", *r)?;
                    sf.set_item("col", *c)?;
                    suggested_fields.append(sf)?;
                }
            }
        }

        entry.set_item("cells", cells_list)?;
        entry.set_item("empty_cells", empty_cells_list)?;
        entry.set_item("suggested_fields", suggested_fields)?;

        tables_out.append(entry)?;
    }
    result.set_item("tables", tables_out)?;
    Ok(result)
}

/// 여러 표·여러 컬럼·여러 셀을 한 번에 채움. Pre-flight 검증 후 batch 적용 + post-fill 라운드트립 검증.
///
/// operations 의 각 dict 형식:
///
///   표 식별 (둘 중 하나 필수):
///     - "header_match": str    — 헤더 행 (row 0) 어딘가에 포함될 텍스트
///     - "table_at": [sec, para, ctrl]   — 직접 좌표
///
///   채우기 모드 (둘 중 하나 필수):
///     - "column": str + "values": {row: str}
///         지정한 컬럼 헤더의 셀들에 행별 값
///     - "cells": [{"row": int, "col": int, "value": str}, ...]
///         (row, col) 직접 지정
///
/// 옵션:
///   - dry_run: bool = False         — true 면 적용·저장 없이 plan 만 반환 (검증 전용)
///   - verify: bool = True           — true 면 저장 후 재파싱하여 모든 셀 값이 보존됐는지 확인
///   - preserve_images: bool = True  — true 면 저장 후 원본 양식의 BinData/Preview stream 을
///                                     raw bytes 그대로 덮어써서 rhwp 의 이미지 라운드트립 손실 우회.
///                                     이미지를 추가/삭제하는 변경이 아닐 때 안전하고 권장됨.
#[pyfunction]
#[pyo3(signature = (template_path, out_path, operations, dry_run=false, verify=true, preserve_images=true))]
fn fill_template<'py>(
    py: Python<'py>,
    template_path: &str,
    out_path: &str,
    operations: Bound<'py, PyList>,
    dry_run: bool,
    verify: bool,
    preserve_images: bool,
) -> PyResult<Bound<'py, PyDict>> {
    if operations.is_empty() {
        return Err(pyo3::exceptions::PyValueError::new_err(
            "operations 가 비어있음 (최소 1 개 필요)",
        ));
    }

    let bytes = io(fs::read(template_path))?;
    let mut core = map_err(DocumentCore::from_bytes(&bytes))?;
    let tables = discover_all_tables(&core);

    // === Pre-flight: 모든 op 의 표·컬럼·범위 검증을 먼저 끝낸다 ===
    let mut planned: Vec<(TableLocation, Vec<CellFill>)> = Vec::with_capacity(operations.len());
    for (i, op_obj) in operations.iter().enumerate() {
        let op: Bound<'_, PyDict> = op_obj.downcast_into().map_err(|_| {
            pyo3::exceptions::PyTypeError::new_err(format!("operations[{}] 는 dict", i))
        })?;
        let table = resolve_table(&op, &tables, i)?;
        let fills = op_to_fills(&op, table, i)?;
        if fills.is_empty() {
            return Err(pyo3::exceptions::PyValueError::new_err(format!(
                "operations[{}] 에서 채울 셀이 없음 ('values' 또는 'cells' 비어있음)",
                i
            )));
        }
        planned.push((table.clone(), fills));
    }

    // dry_run 이면 여기서 plan 만 반환 (양식 무수정)
    if dry_run {
        let r = build_result_dict(py, out_path, 0, &planned, "dry_run", None)?;
        r.set_item("preserved_streams", 0)?;
        return Ok(r);
    }

    // === Apply: 모든 op 가 유효함이 확인됐으니 batch 모드로 일괄 적용 ===
    // 주의: cell_idx 는 op_to_fills 에서 이미 (row, col) → 위치 검색으로 산출됨.
    // 병합 셀이 있는 표는 row*cols+col 공식이 어긋나므로 위치 기반이 정확.
    map_err(core.begin_batch_native())?;
    for (location, fills) in &planned {
        for fill in fills {
            map_err(core.insert_text_in_cell_native(
                location.section,
                location.parent_para,
                location.control,
                fill.cell_idx,
                0,
                0,
                &fill.value,
            ))?;
        }
    }
    map_err(core.end_batch_native())?;

    // === Save ===
    let rhwp_out_bytes = map_err(core.export_hwp_native())?;

    // === BinData/Preview 등은 입력 양식에서 보존, BodyText/DocInfo/FileHeader 만 rhwp 출력 사용 ===
    // rhwp 의 BinData 라운드트립 손실(= 한컴이 "손상" 으로 판정하거나 그림 일부 누락) 회피.
    let (final_bytes, preserved_streams) = if preserve_images {
        let (merged, _from_rhwp, from_input) =
            merge_cfb_preserving_input(&bytes, &rhwp_out_bytes).map_err(|e| {
                pyo3::exceptions::PyRuntimeError::new_err(format!("CFB 머지 실패: {}", e))
            })?;
        (merged, from_input)
    } else {
        (rhwp_out_bytes.clone(), 0)
    };
    let out_bytes = final_bytes;

    if let Some(parent) = std::path::Path::new(out_path).parent() {
        if !parent.as_os_str().is_empty() {
            io(fs::create_dir_all(parent))?;
        }
    }
    io(fs::write(out_path, &out_bytes))?;

    // === Post-fill 검증: 출력을 재파싱하여 plan 의 모든 셀 값이 정확히 보존됐는지 확인 ===
    let mismatches: Vec<String> = if verify {
        let verify_core = map_err(DocumentCore::from_bytes(&out_bytes))?;
        let verify_tables = discover_all_tables(&verify_core);
        let mut errs: Vec<String> = Vec::new();
        for (loc, fills) in &planned {
            let vt = verify_tables.iter().find(|t| {
                t.section == loc.section
                    && t.parent_para == loc.parent_para
                    && t.control == loc.control
            });
            let Some(vt) = vt else {
                errs.push(format!(
                    "재파싱 후 표를 찾을 수 없음 (sec={} para={} ctrl={})",
                    loc.section, loc.parent_para, loc.control
                ));
                continue;
            };
            for fill in fills {
                let actual = vt
                    .cells
                    .iter()
                    .find(|(r, c, _)| *r == fill.row && *c == fill.col)
                    .map(|(_, _, t)| t.as_str())
                    .unwrap_or("(?)");
                // trim_end: HWP 셀은 내용 끝에 \n/공백을 자동 추가하는 관례가 있음.
                // leading whitespace 는 의도적일 수 있으므로 trim_end 만 적용.
                if actual.trim_end() != fill.value.trim_end() {
                    errs.push(format!(
                        "셀 (row={}, col={}) 기대='{}' 실제='{}'",
                        fill.row, fill.col, fill.value, actual
                    ));
                }
            }
        }
        errs
    } else {
        Vec::new()
    };

    let status = if !verify {
        "applied (verify=false)"
    } else if mismatches.is_empty() {
        "applied + verified"
    } else {
        "applied but verification FAILED"
    };

    let result = build_result_dict(
        py,
        out_path,
        out_bytes.len(),
        &planned,
        status,
        Some(&mismatches),
    )?;
    result.set_item("preserved_streams", preserved_streams)?;

    if verify && !mismatches.is_empty() {
        return Err(pyo3::exceptions::PyRuntimeError::new_err(format!(
            "post-fill 검증 실패 ({}건): 첫 항목 = {}",
            mismatches.len(),
            mismatches[0]
        )));
    }

    Ok(result)
}

fn build_result_dict<'py>(
    py: Python<'py>,
    out_path: &str,
    bytes_len: usize,
    planned: &[(TableLocation, Vec<CellFill>)],
    status: &str,
    mismatches: Option<&[String]>,
) -> PyResult<Bound<'py, PyDict>> {
    let result = PyDict::new_bound(py);
    result.set_item("path", out_path)?;
    result.set_item("bytes", bytes_len)?;
    result.set_item("status", status)?;
    if let Some(ms) = mismatches {
        let ml = PyList::empty_bound(py);
        for m in ms {
            ml.append(m)?;
        }
        result.set_item("mismatches", ml)?;
    }
    let ops_list = PyList::empty_bound(py);
    for (loc, fills) in planned {
        let entry = PyDict::new_bound(py);
        let table = PyDict::new_bound(py);
        table.set_item("section", loc.section)?;
        table.set_item("parent_para", loc.parent_para)?;
        table.set_item("control", loc.control)?;
        table.set_item("rows", loc.rows)?;
        table.set_item("cols", loc.cols)?;
        entry.set_item("table", table)?;
        let applied_list = PyList::empty_bound(py);
        for f in fills {
            let af = PyDict::new_bound(py);
            af.set_item("row", f.row)?;
            af.set_item("col", f.col)?;
            af.set_item("value", &f.value)?;
            applied_list.append(af)?;
        }
        entry.set_item("applied", applied_list)?;
        ops_list.append(entry)?;
    }
    result.set_item("operations", ops_list)?;
    Ok(result)
}

/// 단일 표 단일 컬럼 편의 함수 — 내부적으로 fill_template 한 번 호출.
///
/// mapping = {
///   "header_match": str,    # 표 식별
///   "column": str,          # 컬럼 헤더
///   "values": {int: str},   # 행 인덱스 → 값
/// }
#[pyfunction]
#[pyo3(signature = (template_path, out_path, mapping, dry_run=false, verify=true, preserve_images=true))]
fn fill_template_table<'py>(
    py: Python<'py>,
    template_path: &str,
    out_path: &str,
    mapping: Bound<'py, PyDict>,
    dry_run: bool,
    verify: bool,
    preserve_images: bool,
) -> PyResult<Bound<'py, PyDict>> {
    let operations = PyList::empty_bound(py);
    operations.append(mapping)?;
    fill_template(
        py,
        template_path,
        out_path,
        operations,
        dry_run,
        verify,
        preserve_images,
    )
}

#[pymodule]
fn hwp_automate(_py: Python, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(analyze_template, m)?)?;
    m.add_function(wrap_pyfunction!(fill_template, m)?)?;
    m.add_function(wrap_pyfunction!(fill_template_table, m)?)?;
    m.add_function(wrap_pyfunction!(preserve_images_from_source, m)?)?;
    Ok(())
}
