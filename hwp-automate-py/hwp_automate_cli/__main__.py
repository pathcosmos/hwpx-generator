"""CLI 진입점 — `python -m hwp_automate_cli <subcommand>`

서브커맨드:
  analyze   템플릿 양식의 표·스타일 인벤토리 출력
  fill      field_map.json + 데이터로 양식의 여러 셀을 한 번에 채움
  cell      표 1개·셀 몇 개 빠르게 채우는 단축 명령

예시:
  python -m hwp_automate_cli analyze --template ./양식.hwp

  python -m hwp_automate_cli fill \\
    --template ./양식.hwp \\
    --field-map ./templates/.../field_map.json \\
    --data ./data/sample_input.json \\
    --output ./out/채워진.hwp \\
    --header-match "성명"           # 또는 --table-at 0 70 0

  python -m hwp_automate_cli cell \\
    --template ./양식.hwp \\
    --output ./out/quick.hwp \\
    --header-match "성명" \\
    --cell 1,5,정보처리기사  --cell 2,5,정보보안기사
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import hwp_automate

from .field_map import field_map_to_operations


def _table_locator_from_args(args) -> dict:
    if args.table_at is not None:
        return {"table_at": list(args.table_at)}
    if args.header_match:
        return {"header_match": args.header_match}
    raise SystemExit("ERROR: --header-match 또는 --table-at 중 하나는 필수")


def cmd_analyze(args) -> int:
    info = hwp_automate.analyze_template(args.template)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
        return 0
    print(f"파일:      {info['path']}")
    print(f"크기:      {info['file_size']:,} bytes")
    print(f"섹션:      {info['section_count']}")
    print(f"문단:      {info['paragraph_count']}")
    print(
        f"스타일:    {info['style_count']} (char={info['char_shape_count']}, "
        f"para={info['para_shape_count']}, numbering={info['numbering_count']}, "
        f"border_fill={info['border_fill_count']})"
    )
    print(f"\n표 {len(info['tables'])} 개:")
    for i, t in enumerate(info["tables"]):
        print(
            f"  [{i}] sec={t['section']} para={t['parent_para']} ctrl={t['control']}"
            f"  {t['rows']}x{t['cols']}  헤더: {t['header']}"
        )
    return 0


def cmd_fill(args) -> int:
    locator = _table_locator_from_args(args)
    fm_path = Path(args.field_map)
    data_path = Path(args.data)
    if not fm_path.exists():
        raise SystemExit(f"ERROR: field-map 없음: {fm_path}")
    if not data_path.exists():
        raise SystemExit(f"ERROR: data 없음: {data_path}")

    field_map = json.loads(fm_path.read_text(encoding="utf-8"))
    data = json.loads(data_path.read_text(encoding="utf-8"))
    operations = field_map_to_operations(
        field_map, data, table_locator=locator, skip_empty=not args.include_empty
    )
    if not operations:
        raise SystemExit(
            "ERROR: 변환 결과 operation 이 0 개. field_map / data 가 비었거나 매칭 안 됨."
        )

    if args.print_operations:
        print(json.dumps(operations, ensure_ascii=False, indent=2))
        if args.dry_run:
            return 0

    result = hwp_automate.fill_template(
        args.template,
        args.output,
        operations,
        dry_run=args.dry_run,
        verify=not args.no_verify,
    )
    print(f"status:  {result['status']}")
    if result.get("mismatches"):
        for m in result["mismatches"]:
            print(f"  - {m}", file=sys.stderr)
    print(f"path:    {result['path']}")
    print(f"bytes:   {result['bytes']:,}")
    print(f"ops:     {len(result['operations'])}")
    total_cells = sum(len(o["applied"]) for o in result["operations"])
    print(f"cells:   {total_cells}")
    return 0


def cmd_cell(args) -> int:
    locator = _table_locator_from_args(args)
    cells = []
    for raw in args.cell or []:
        parts = raw.split(",", 2)
        if len(parts) != 3:
            raise SystemExit(f"ERROR: --cell 형식: row,col,value (받음: {raw!r})")
        try:
            row, col = int(parts[0]), int(parts[1])
        except ValueError as e:
            raise SystemExit(f"ERROR: --cell row/col 정수 아님: {raw!r} ({e})")
        cells.append({"row": row, "col": col, "value": parts[2]})
    if not cells:
        raise SystemExit("ERROR: --cell 최소 1개 필요")

    operations = [{**locator, "cells": cells}]
    if args.print_operations:
        print(json.dumps(operations, ensure_ascii=False, indent=2))
        if args.dry_run:
            return 0

    result = hwp_automate.fill_template(
        args.template,
        args.output,
        operations,
        dry_run=args.dry_run,
        verify=not args.no_verify,
    )
    print(f"status:  {result['status']}")
    if result.get("mismatches"):
        for m in result["mismatches"]:
            print(f"  - {m}", file=sys.stderr)
    print(f"path:    {result['path']}")
    print(f"bytes:   {result['bytes']:,}")
    print(f"cells:   {len(result['operations'][0]['applied'])}")
    return 0


def _add_locator_args(p: argparse.ArgumentParser) -> None:
    g = p.add_mutually_exclusive_group(required=False)
    g.add_argument("--header-match", help="표 식별: 헤더 행에 포함될 텍스트")
    g.add_argument(
        "--table-at",
        nargs=3,
        type=int,
        metavar=("SEC", "PARA", "CTRL"),
        help="표 식별: 직접 좌표 (sec, parent_para, ctrl)",
    )


def _add_fill_options(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="적용·저장 없이 plan 만 검증·반환",
    )
    p.add_argument(
        "--no-verify",
        action="store_true",
        help="저장 후 라운드트립 검증 생략 (기본은 verify on)",
    )
    p.add_argument(
        "--print-operations",
        action="store_true",
        help="실행 전 변환된 operations JSON 출력",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="hwp_automate_cli",
        description="hwp_automate (rhwp 기반 한컴 HWP 자동화) 명령행 진입점",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_analyze = sub.add_parser("analyze", help="양식의 표·스타일 인벤토리")
    p_analyze.add_argument("--template", required=True)
    p_analyze.add_argument("--json", action="store_true", help="JSON 형식 출력")
    p_analyze.set_defaults(func=cmd_analyze)

    p_fill = sub.add_parser("fill", help="field_map.json + 데이터로 일괄 채우기")
    p_fill.add_argument("--template", required=True)
    p_fill.add_argument("--field-map", required=True, help="field_map.json 경로")
    p_fill.add_argument("--data", required=True, help="data JSON 경로")
    p_fill.add_argument("--output", required=True)
    p_fill.add_argument(
        "--include-empty",
        action="store_true",
        help="빈 값도 셀에 쓰기 (기본은 빈 값 셀 제외)",
    )
    _add_locator_args(p_fill)
    _add_fill_options(p_fill)
    p_fill.set_defaults(func=cmd_fill)

    p_cell = sub.add_parser("cell", help="단일 표 셀 몇 개 빠르게 채우기")
    p_cell.add_argument("--template", required=True)
    p_cell.add_argument("--output", required=True)
    p_cell.add_argument(
        "--cell",
        action="append",
        metavar="ROW,COL,VALUE",
        help="셀 1개 (반복 가능)",
    )
    _add_locator_args(p_cell)
    _add_fill_options(p_cell)
    p_cell.set_defaults(func=cmd_cell)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
