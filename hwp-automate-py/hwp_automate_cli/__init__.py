"""hwp_automate_cli — hwp_automate (Rust 확장) 의 사용 편의를 위한 Python 헬퍼 모음.

이 모듈은 wheel 에 번들되지 않은 보조 도구다. 사용자는 다음 두 가지를 함께 임포트:
  - hwp_automate          (Rust 확장: analyze_template, fill_template, fill_template_table)
  - hwp_automate_cli      (Python: field_map.json 어댑터, CLI 진입점)
"""

from .field_map import field_map_to_operations, resolve_data_path

__all__ = ["field_map_to_operations", "resolve_data_path"]
