//! 기존 양식의 표 채우기 데모.
//!
//! 패턴: 사용자 양식(.hwp)을 로드해 헤더 매칭으로 표/컬럼을 자동 식별하고,
//! 빈 셀에만 값을 삽입한다. 양식의 구조/스타일/병합/테두리는 무손상 보존.
//!
//! 시나리오: biz_plan.hwp 의 5×6 인력 명단 표 — 자격증 컬럼 빈 셀 4개 채우기.
//!
//! 사용:
//!   cargo run                                  # 기본 템플릿 + 기본 출력
//!   cargo run -- <template.hwp> <output.hwp>   # 경로 지정
//!
//! 주의: from-scratch 로 새 문서를 만드는 패턴은 사용하지 않는다.
//! 항상 사용자가 제공한 양식을 베이스로 한다.

use rhwp::document_core::DocumentCore;
use rhwp::error::HwpError;
use rhwp::model::control::Control;
use std::fs;
use std::path::Path;

type DynErr = Box<dyn std::error::Error>;

fn h(e: HwpError) -> DynErr {
    format!("HwpError: {}", e).into()
}

#[derive(Debug, Clone)]
struct TableLocation {
    section: usize,
    parent_para: usize,
    control: usize,
    rows: u16,
    cols: u16,
    /// (row, col, joined_text) — 행/열 주소와 셀 텍스트 (문단 사이는 "|" 결합)
    cells: Vec<(u16, u16, String)>,
}

fn main() -> Result<(), DynErr> {
    let mut args = std::env::args();
    let _bin = args.next();
    let template_path = args
        .next()
        .unwrap_or_else(|| "../../codebase/rhwp/samples/biz_plan.hwp".to_string());
    let out_path = args
        .next()
        .unwrap_or_else(|| "output/poc_v3.hwp".to_string());

    println!("=== 1. 템플릿 로드: {} ===", template_path);
    let bytes = fs::read(&template_path)?;
    let mut core = DocumentCore::from_bytes(&bytes).map_err(h)?;

    println!("\n=== 2. 양식 내 표 발견 (검색) ===");
    let tables = discover_tables(&core);
    println!("  총 표 {}개 발견:", tables.len());
    for (i, t) in tables.iter().enumerate() {
        let header_row: Vec<&str> = t
            .cells
            .iter()
            .filter(|(r, _, _)| *r == 0)
            .map(|(_, _, txt)| txt.as_str())
            .collect();
        println!(
            "    [{}] sec={} para={} ctrl={} : {}x{} | 헤더: {:?}",
            i, t.section, t.parent_para, t.control, t.rows, t.cols, header_row
        );
    }

    println!("\n=== 3. 채울 표 선택: 5×6 인력 명단 (heuristic — '성명' 헤더 포함) ===");
    let target = tables
        .iter()
        .find(|t| {
            t.rows == 5
                && t.cols == 6
                && t.cells
                    .iter()
                    .any(|(r, _, txt)| *r == 0 && txt.contains("성명"))
        })
        .ok_or("'성명' 헤더가 있는 5×6 표를 찾을 수 없음")?;
    println!(
        "  선택: sec={} para={} ctrl={}",
        target.section, target.parent_para, target.control
    );

    println!("\n=== 4. '자격증' 컬럼 인덱스 탐색 ===");
    let cert_col = target
        .cells
        .iter()
        .find(|(r, _, txt)| *r == 0 && txt == "자격증")
        .map(|(_, c, _)| *c)
        .ok_or("'자격증' 헤더 컬럼을 찾을 수 없음")?;
    println!("  '자격증' 컬럼 = {}", cert_col);

    println!("\n=== 5. 각 인원의 자격증 셀에 값 삽입 ===");
    // 행별 채울 자격증 (row 1~4 = 데이터 행)
    let certifications = [
        (1u16, "정보처리기사"),
        (2u16, "정보보안기사"),
        (3u16, "네트워크관리사"),
        (4u16, "컴활 1급"),
    ];

    core.begin_batch_native().map_err(h)?;
    for (row, value) in &certifications {
        // 셀 인덱스 (row-major: row*cols + col)
        let cell_idx = (*row as usize) * (target.cols as usize) + (cert_col as usize);

        // 현재 셀 텍스트 확인
        let current = target
            .cells
            .iter()
            .find(|(r, c, _)| *r == *row && *c == cert_col)
            .map(|(_, _, t)| t.as_str())
            .unwrap_or("?");

        // 사람 이름도 함께 출력 (컬럼 c=2 = '성명')
        let name = target
            .cells
            .iter()
            .find(|(r, c, _)| *r == *row && *c == 2)
            .map(|(_, _, t)| t.as_str())
            .unwrap_or("?");

        println!(
            "  행 {} '{}' 의 자격증: '{}' -> '{}'  (cell_idx={})",
            row, name, current, value, cell_idx
        );

        let r = core
            .insert_text_in_cell_native(
                target.section,
                target.parent_para,
                target.control,
                cell_idx,
                0, // cell_para_idx (셀 내 첫 문단)
                0, // char_offset
                value,
            )
            .map_err(h)?;
        if !r.contains("\"ok\":true") {
            println!("    경고: {}", r);
        }
    }
    core.end_batch_native().map_err(h)?;

    println!("\n=== 6. HWP 5.0 binary 저장 ===");
    let output_bytes = core.export_hwp_native().map_err(h)?;
    let out_dir = Path::new(&out_path).parent().unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(out_dir)?;
    fs::write(&out_path, &output_bytes)?;
    println!("  저장: {} ({} bytes)", out_path, output_bytes.len());

    println!("\n=== 7. 라운드트립 검증: 채워진 셀의 텍스트 확인 ===");
    let verify_core = DocumentCore::from_bytes(&output_bytes).map_err(h)?;
    let verify_tables = discover_tables(&verify_core);
    let verify_target = verify_tables
        .iter()
        .find(|t| {
            t.section == target.section
                && t.parent_para == target.parent_para
                && t.control == target.control
        })
        .ok_or("출력 파일에서 대상 표 재발견 실패")?;

    for (row, expected) in &certifications {
        let actual = verify_target
            .cells
            .iter()
            .find(|(r, c, _)| *r == *row && *c == cert_col)
            .map(|(_, _, t)| t.as_str())
            .unwrap_or("(?)");
        let ok = actual == *expected;
        println!(
            "  행 {}: 기대='{}' 실제='{}'  {}",
            row,
            expected,
            actual,
            if ok { "✓" } else { "✗ FAIL" }
        );
    }

    println!("\n✅ PoC v3 (기존 양식 표 채우기) 완료");
    Ok(())
}

/// Document 안의 모든 표를 순회하며 위치+셀 정보를 수집.
fn discover_tables(core: &DocumentCore) -> Vec<TableLocation> {
    let mut out = Vec::new();
    let doc = core.document();
    for (sec_idx, section) in doc.sections.iter().enumerate() {
        for (para_idx, para) in section.paragraphs.iter().enumerate() {
            for (ctrl_idx, ctrl) in para.controls.iter().enumerate() {
                if let Control::Table(table) = ctrl {
                    let cells: Vec<(u16, u16, String)> = table
                        .cells
                        .iter()
                        .map(|cell| {
                            let text = cell
                                .paragraphs
                                .iter()
                                .map(|p| p.text.as_str())
                                .collect::<Vec<_>>()
                                .join("|");
                            (cell.row, cell.col, text)
                        })
                        .collect();
                    out.push(TableLocation {
                        section: sec_idx,
                        parent_para: para_idx,
                        control: ctrl_idx,
                        rows: table.row_count,
                        cols: table.col_count,
                        cells,
                    });
                }
            }
        }
    }
    out
}
