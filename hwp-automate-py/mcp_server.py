"""hwp-automate MCP 서버 — FastMCP stdio.

Claude Desktop / Claude Code / Cursor 등 MCP 호환 클라이언트가 자연어로
HWP 양식 자동 채우기 도구를 사용할 수 있게 함.

실행:
    /path/to/.venv/bin/python /path/to/hwp-automate-py/mcp_server.py
또는 uv:
    uv run --directory /path/to/hwp-automate-py python mcp_server.py

Claude Desktop 설정 예 (claude_desktop_config.json):
{
  "mcpServers": {
    "hwp-automate": {
      "command": "/path/to/hwp-automate-py/.venv/bin/python",
      "args": ["/path/to/hwp-automate-py/mcp_server.py"]
    }
  }
}

Claude Code 등록:
    claude mcp add hwp-automate -- /path/to/.venv/bin/python /path/to/mcp_server.py

5 개 tool 노출:
    analyze_form           — 양식 구조 + 빈 셀 + 라벨 추론 (AI 가 즉시 의미 파악)
    preview_form_structure — 가벼운 markdown 요약
    fill_form              — operations 로 양식 채우기 (dry_run, verify 지원)
    fill_form_from_data    — field_map.json + data.json 호환 입력
    verify_output          — 결과 파일 셀 값 라운드트립 검증
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

# stdout 은 MCP 프로토콜 전용. 모든 진단 로그는 stderr.
# 진입점에서 한 줄만 출력 — import 시점에는 침묵.

from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]

import hwp_automate

# field_map 어댑터는 hwp_automate_cli 패키지에 있음
sys.path.insert(0, str(Path(__file__).parent))
from hwp_automate_cli.field_map import field_map_to_operations  # noqa: E402

mcp = FastMCP("hwp-automate")


def _resolve(path: str) -> str:
    """사용자가 ~ 또는 상대 경로를 줘도 절대 경로로 정규화."""
    return str(Path(path).expanduser().resolve())


@mcp.tool()
def analyze_form(template_path: str) -> dict:
    """HWP/HWPX 양식 분석 — 표 구조, 빈 셀, 라벨 추론 결과를 반환.

    각 표 entry 의 ``empty_cells`` 와 ``suggested_fields`` 가 핵심:
      - empty_cells: 빈 셀 목록 + 인접 셀에서 추론한 ``neighbor_label``
      - suggested_fields: 라벨이 추론된 빈 셀만 — AI 가 사용자에게 물어볼 후보

    Args:
        template_path: HWP/HWPX 파일 경로 (절대, 상대, 또는 ~/...)

    Returns:
        분석 결과 dict — section_count, table count, tables 리스트 (cells/empty_cells/suggested_fields 포함)
    """
    return hwp_automate.analyze_template(_resolve(template_path))


@mcp.tool()
def preview_form_structure(template_path: str) -> dict:
    """양식 구조의 가벼운 markdown 요약 — Claude 가 사용자에게 보여줄 표 인벤토리.

    analyze_form 보다 컨텍스트 윈도우 절약. 큰 양식(40MB·표 16개) 첫 검토에 적합.

    Args:
        template_path: HWP/HWPX 파일 경로

    Returns:
        {"markdown": "표 목록 markdown", "table_count": int}
    """
    info = hwp_automate.analyze_template(_resolve(template_path))
    name = Path(template_path).name
    lines = [
        f"# {name}",
        "",
        f"- 섹션: {info['section_count']} / 문단: {info['paragraph_count']}",
        f"- 스타일: {info['style_count']} / 표: {len(info['tables'])}",
        "",
        "## 표 목록",
        "",
        "| # | 위치(sec,para,ctrl) | 크기 | 빈셀 | 헤더 |",
        "|---|---|---|---|---|",
    ]
    for i, t in enumerate(info["tables"]):
        loc = f"{t['section']},{t['parent_para']},{t['control']}"
        size = f"{t['rows']}×{t['cols']}"
        empty = len(t.get("empty_cells", []))
        header_short = ", ".join(t["header"][:6])
        if len(t["header"]) > 6:
            header_short += " ..."
        lines.append(f"| {i} | [{loc}] | {size} | {empty} | {header_short} |")
    return {"markdown": "\n".join(lines), "table_count": len(info["tables"])}


@mcp.tool()
def fill_form(
    template_path: str,
    output_path: str,
    operations: list,
    dry_run: bool = False,
    verify: bool = True,
    preserve_images: bool = True,
) -> dict:
    """양식을 operations 에 따라 채워 출력 파일로 저장.

    operations 의 각 dict 형식:
        표 식별 (둘 중 하나 필수):
          - "header_match": str    (헤더 행에 포함될 텍스트)
          - "table_at": [sec, para, ctrl]
        채우기 모드 (둘 중 하나 필수):
          - "column": str + "values": {row: str}     (컬럼 헤더 자동 탐색)
          - "cells": [{"row": int, "col": int, "value": str}, ...]

    Pre-flight 검증 + post-fill 라운드트립 검증 + BinData 보존 모두 자동.

    Args:
        template_path: 원본 양식 경로
        output_path: 결과 저장 경로
        operations: 채우기 작업 list (위 형식)
        dry_run: True 면 plan 만 검증, 파일 무수정
        verify: True 면 저장 후 라운드트립 검증 (권장 — 값 보존 자동 확인)
        preserve_images: True 면 원본 이미지·미리보기 stream 보존 (권장 — false 시 한컴이 손상으로 인식할 수 있음)

    Returns:
        {status, path, bytes, operations: [...], mismatches: [...]?}
    """
    return hwp_automate.fill_template(
        _resolve(template_path),
        _resolve(output_path),
        operations,
        dry_run=dry_run,
        verify=verify,
        preserve_images=preserve_images,
    )


@mcp.tool()
def fill_form_from_data(
    template_path: str,
    output_path: str,
    field_map_path: str,
    data_path: str,
    table_locator: dict,
    dry_run: bool = False,
    skip_empty: bool = True,
) -> dict:
    """기존 field_map.json + data.json 형식으로 양식 채우기.

    hwpx-generator 의 ``templates/.../field_map.json`` 형식 (entity_blocks, company_lists)
    을 ``cells`` operations 로 자동 변환하여 fill_form 호출.

    Args:
        template_path: 양식 경로
        output_path: 결과 경로
        field_map_path: field_map.json 경로
        data_path: data.json 경로 (사업자 정보·인력·예산 등)
        table_locator: 대상 표 식별 — {"header_match": str} 또는 {"table_at": [s,p,c]}
        dry_run: True 면 plan 만
        skip_empty: True (기본) 면 빈 값 셀 안 채움

    Returns:
        fill_form 과 동일
    """
    field_map = json.loads(Path(_resolve(field_map_path)).read_text(encoding="utf-8"))
    data = json.loads(Path(_resolve(data_path)).read_text(encoding="utf-8"))
    operations = field_map_to_operations(
        field_map, data, table_locator=table_locator, skip_empty=skip_empty
    )
    if not operations:
        raise ValueError(
            "변환 결과 operations 가 0개. field_map.json 의 entity_blocks/"
            "company_lists 가 비었거나 data.json 과 매칭 안 됨."
        )
    return hwp_automate.fill_template(
        _resolve(template_path),
        _resolve(output_path),
        operations,
        dry_run=dry_run,
    )


@mcp.tool()
def verify_output(output_path: str, expected_cells: list) -> dict:
    """출력 파일의 특정 셀 값을 라운드트립으로 재확인.

    fill_form 의 verify=True 가 자동 검증하지만, 별도로 특정 셀만 검증하고 싶을 때 사용.

    Args:
        output_path: 검증 대상 HWP 경로
        expected_cells: [{header_match 또는 table_at, row, col, expected_value}, ...]

    Returns:
        {"all_match": bool, "mismatches": [str, ...]}
    """
    info = hwp_automate.analyze_template(_resolve(output_path))
    mismatches: list[str] = []
    for exp in expected_cells:
        target_table = None
        if "header_match" in exp:
            for t in info["tables"]:
                if any(exp["header_match"] in h for h in t["header"]):
                    target_table = t
                    break
        elif "table_at" in exp:
            sec, para, ctrl = exp["table_at"]
            for t in info["tables"]:
                if (
                    t["section"] == sec
                    and t["parent_para"] == para
                    and t["control"] == ctrl
                ):
                    target_table = t
                    break
        if target_table is None:
            mismatches.append(f"표 식별 실패: {exp}")
            continue
        cell = next(
            (
                c
                for c in target_table["cells"]
                if c["row"] == exp["row"] and c["col"] == exp["col"]
            ),
            None,
        )
        if cell is None:
            mismatches.append(
                f"셀 (r={exp['row']}, c={exp['col']}) 없음 (병합 또는 범위 외)"
            )
            continue
        actual = cell["text"].rstrip()
        expected = exp["expected_value"].rstrip()
        if actual != expected:
            mismatches.append(
                f"셀 (r={exp['row']}, c={exp['col']}) 기대='{expected}' 실제='{actual}'"
            )
    return {"all_match": len(mismatches) == 0, "mismatches": mismatches}


def main() -> None:
    """진입점. stdio 전송으로 MCP 서버 실행.

    주의: stdout 은 MCP 프로토콜 전용. 디버그 출력은 모두 stderr 로.
    """
    print("hwp-automate MCP 서버 시작 (stdio)", file=sys.stderr)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
