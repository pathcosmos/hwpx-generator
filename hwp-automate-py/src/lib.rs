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
use std::fs;

fn map_err<T>(r: Result<T, HwpError>) -> PyResult<T> {
    r.map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("HwpError: {}", e)))
}

fn io<T>(r: std::io::Result<T>) -> PyResult<T> {
    r.map_err(|e| pyo3::exceptions::PyIOError::new_err(e.to_string()))
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
#[derive(Debug, Clone)]
struct CellFill {
    row: u16,
    col: u16,
    value: String,
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
            fills.push(CellFill {
                row: row as u16,
                col: target_col,
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
            fills.push(CellFill {
                row: row as u16,
                col: col as u16,
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
        let header: Vec<&str> = t
            .cells
            .iter()
            .filter(|(r, _, _)| *r == 0)
            .map(|(_, _, txt)| txt.as_str())
            .collect();
        entry.set_item("header", header)?;
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
///   - dry_run: bool = False   — true 면 적용·저장 없이 plan 만 반환 (검증 전용)
///   - verify: bool = True     — true 면 저장 후 재파싱하여 모든 셀 값이 보존됐는지 확인
#[pyfunction]
#[pyo3(signature = (template_path, out_path, operations, dry_run=false, verify=true))]
fn fill_template<'py>(
    py: Python<'py>,
    template_path: &str,
    out_path: &str,
    operations: Bound<'py, PyList>,
    dry_run: bool,
    verify: bool,
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
        return build_result_dict(py, out_path, 0, &planned, "dry_run", None);
    }

    // === Apply: 모든 op 가 유효함이 확인됐으니 batch 모드로 일괄 적용 ===
    map_err(core.begin_batch_native())?;
    for (location, fills) in &planned {
        for fill in fills {
            let cell_idx = (fill.row as usize) * (location.cols as usize) + (fill.col as usize);
            map_err(core.insert_text_in_cell_native(
                location.section,
                location.parent_para,
                location.control,
                cell_idx,
                0,
                0,
                &fill.value,
            ))?;
        }
    }
    map_err(core.end_batch_native())?;

    // === Save ===
    let out_bytes = map_err(core.export_hwp_native())?;
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
                if actual != fill.value {
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
#[pyo3(signature = (template_path, out_path, mapping, dry_run=false, verify=true))]
fn fill_template_table<'py>(
    py: Python<'py>,
    template_path: &str,
    out_path: &str,
    mapping: Bound<'py, PyDict>,
    dry_run: bool,
    verify: bool,
) -> PyResult<Bound<'py, PyDict>> {
    let operations = PyList::empty_bound(py);
    operations.append(mapping)?;
    fill_template(py, template_path, out_path, operations, dry_run, verify)
}

#[pymodule]
fn hwp_automate(_py: Python, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(analyze_template, m)?)?;
    m.add_function(wrap_pyfunction!(fill_template, m)?)?;
    m.add_function(wrap_pyfunction!(fill_template_table, m)?)?;
    Ok(())
}
