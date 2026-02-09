"""Field mapper: maps input JSON data to cover table cell addresses using field_map.json."""

import json
import os


def load_field_map(template_dir):
    """Load field_map.json from template directory."""
    path = os.path.join(template_dir, "field_map.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_data_path(data, path):
    """Resolve dot-notation path like '담당자.성명' -> data['담당자']['성명'].
    Returns None if any key is missing."""
    parts = path.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return None
    return current


def build_cell_data(input_data, field_map):
    """Convert input JSON data to {(row, col): text} dict using field_map.

    Processes:
    1. entity_blocks: 대표공급기업, 클라우드사업자, 협력기관
    2. company_lists: 참여공급기업, 도입실증기업

    Skips empty strings and None values.
    Returns dict of {(row_addr, col_addr): text_value}
    """
    cell_data = {}

    # Process entity blocks
    for block in field_map.get("entity_blocks", []):
        data_path = block["data_path"]
        entity_data = input_data.get(data_path)
        if entity_data is None:
            continue

        start_row = block["start_row"]

        for field_def in block["fields"]:
            row = start_row + field_def["offset"]

            # Left side
            if "left" in field_def:
                left = field_def["left"]
                value = resolve_data_path(entity_data, left["field"])
                if value is not None and value != "":
                    cell_data[(row, left["col"])] = str(value)

            # Right side
            if "right" in field_def:
                right = field_def["right"]
                value = resolve_data_path(entity_data, right["field"])
                if value is not None and value != "":
                    cell_data[(row, right["col"])] = str(value)

    # Process company lists
    for comp_list in field_map.get("company_lists", []):
        data_path = comp_list["data_path"]
        items = input_data.get(data_path)
        if items is None:
            continue

        data_start_row = comp_list["data_start_row"]
        max_items = comp_list["max_items"]
        columns = comp_list["columns"]

        for i, item in enumerate(items[:max_items]):
            row = data_start_row + i
            for col_str, field_name in columns.items():
                col = int(col_str)
                value = item.get(field_name)
                if value is not None and value != "":
                    cell_data[(row, col)] = str(value)

    return cell_data
