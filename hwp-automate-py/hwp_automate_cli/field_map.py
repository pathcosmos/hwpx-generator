"""기존 hwpx-generator 의 field_map.json 형식을 hwp_automate.fill_template 의
operations 리스트로 변환하는 어댑터.

field_map.json 형식 (요약):

  {
    "entity_blocks": [
      {
        "name": "...",                # 사람용 라벨
        "data_path": "기관명",          # 입력 data dict 의 시작점 (dot-notated)
        "start_row": 6,                # 표 안 시작 행
        "fields": [
          { "offset": 0,
            "left":  {"col": 3, "field": "기업명"},
            "right": {"col": 9, "field": "사업자등록번호"}
          },
          ...
        ]
      }
    ],
    "company_lists": [
      {
        "name": "...",
        "data_path": "참여공급기업",
        "data_start_row": 26,
        "max_items": 3,
        "columns": {"1": "번호", "2": "기업명", ...}
      }
    ]
  }

이 형식을 우리의 fill_template operations (cells 모드) 리스트로 1:1 매핑한다.
같은 양식의 같은 표를 채우므로 모든 operation 은 같은 table_locator 를 공유.
"""

from __future__ import annotations

from typing import Any, Iterable


def resolve_data_path(data: Any, path: str) -> Any:
    """`담당자.성명` 같은 dot-notated path 를 따라 data 를 내려간다.
    각 단계에서 찾을 수 없으면 None 반환."""
    if path == "" or path is None:
        return data
    cur = data
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur


def _format_value(value: Any) -> str:
    """fill 에 들어갈 값을 문자열로 정규화. None 은 빈 문자열."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        # 리스트는 줄바꿈 결합 (예: 주요솔루션)
        return "\n".join(str(v) for v in value)
    return str(value)


def _entity_block_cells(
    block: dict, data: dict, skip_empty: bool
) -> Iterable[dict]:
    """entity_blocks 의 한 항목 → cells (row, col, value) 시퀀스."""
    block_data = resolve_data_path(data, block["data_path"]) or {}
    if not isinstance(block_data, dict):
        return
    start_row = int(block["start_row"])
    for f in block.get("fields", []):
        row = start_row + int(f["offset"])
        for side in ("left", "right"):
            spec = f.get(side)
            if not spec:
                continue
            value = resolve_data_path(block_data, spec["field"])
            value_str = _format_value(value)
            if skip_empty and value_str == "":
                continue
            yield {
                "row": row,
                "col": int(spec["col"]),
                "value": value_str,
            }


def _company_list_cells(
    cl: dict, data: dict, skip_empty: bool
) -> Iterable[dict]:
    """company_lists 의 한 항목 → cells (row, col, value) 시퀀스."""
    items = resolve_data_path(data, cl["data_path"]) or []
    if not isinstance(items, list):
        return
    max_items = int(cl.get("max_items", len(items)))
    data_start_row = int(cl["data_start_row"])
    columns = cl["columns"]
    for i, item in enumerate(items[:max_items]):
        row = data_start_row + i
        for col_str, field_name in columns.items():
            col = int(col_str)
            if field_name == "번호":
                # 번호 컬럼은 1-indexed 자동 채움 (data 의 필드 무시)
                value_str = str(i + 1)
            else:
                value = resolve_data_path(item, field_name)
                value_str = _format_value(value)
            if skip_empty and value_str == "":
                continue
            yield {"row": row, "col": col, "value": value_str}


def field_map_to_operations(
    field_map: dict,
    data: dict,
    table_locator: dict,
    skip_empty: bool = True,
) -> list[dict]:
    """field_map.json 형식 → hwp_automate.fill_template operations 리스트.

    Args:
        field_map: 파싱된 field_map.json (dict).
        data: 양식에 채워 넣을 데이터 (dict).  field_map 의 data_path 들이 여기서 해석됨.
        table_locator: 대상 표 식별 — {"header_match": str} 또는 {"table_at": [s, p, c]}.
            모든 entity_block / company_list 가 이 같은 표를 채운다고 가정한다 (기존
            cover_table_index 패턴과 동일).  여러 표 대상이면 호출을 나누거나 결과를 결합.
        skip_empty: True 면 값이 빈 문자열인 셀은 cells 에서 빼고, 양식의 기존 빈 셀 유지.

    Returns:
        fill_template 의 operations 인자에 그대로 전달 가능한 dict 리스트.
        각 operation 은 {**table_locator, "cells": [{row, col, value}, ...]} 형태.
        entity_blocks / company_lists 의 항목별로 1 개씩 생성된다 (같은 표라도 분리해 두면
        operations 결과 dict 에서 어느 블록 / 리스트가 채워졌는지 추적 가능).
    """
    if not isinstance(field_map, dict):
        raise TypeError("field_map 은 dict")
    if "header_match" not in table_locator and "table_at" not in table_locator:
        raise ValueError("table_locator 에 'header_match' 또는 'table_at' 필요")

    operations: list[dict] = []

    for block in field_map.get("entity_blocks", []):
        cells = list(_entity_block_cells(block, data, skip_empty))
        if cells:
            operations.append({**table_locator, "cells": cells})

    for cl in field_map.get("company_lists", []):
        cells = list(_company_list_cells(cl, data, skip_empty))
        if cells:
            operations.append({**table_locator, "cells": cells})

    return operations
