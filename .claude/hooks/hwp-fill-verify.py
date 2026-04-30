#!/usr/bin/env python3
"""PostToolUse hook — fill_template 호출 직후 결과 파일을 자동 검증·보고.

활성 조건: Bash tool 의 command 가 'fill_template' 또는 'hwp_automate_cli ... fill'
또는 'fill_form' (MCP tool 명) 패턴을 포함.

stdin 으로 hook 입력 JSON 을 받고, stdout 에 Claude 가 다음 turn 에서 참고할
additionalContext 를 출력.
"""

from __future__ import annotations

import json
import os
import re
import sys


def matches_fill_command(command: str) -> bool:
    if not command:
        return False
    triggers = (
        "fill_template",
        "fill_form",
        "hwp_automate_cli fill",
        "hwp_automate_cli cell",
    )
    return any(t in command for t in triggers)


def extract_output_path(command: str) -> str | None:
    """command 에서 출력 파일 경로를 휴리스틱으로 추출.

    1) `--output PATH` 또는 `-o PATH`
    2) `out_path="PATH"` Python 키워드 인자
    3) 두 번째 위치 인자 (CLI fill 의 양식 다음)
    """
    # CLI 인자
    m = re.search(r"--output[= ]+([^\s'\"]+\.hwp[x]?)", command)
    if m:
        return m.group(1)
    m = re.search(r"-o[= ]+([^\s'\"]+\.hwp[x]?)", command)
    if m:
        return m.group(1)
    # Python 키워드
    m = re.search(r"out_path\s*=\s*['\"]([^'\"]+\.hwp[x]?)", command)
    if m:
        return m.group(1)
    m = re.search(r"output_path\s*=\s*['\"]([^'\"]+\.hwp[x]?)", command)
    if m:
        return m.group(1)
    return None


def main() -> int:
    try:
        data = json.load(sys.stdin)
    except (ValueError, OSError):
        return 0  # hook 은 조용히 실패 — Bash 결과는 그대로 반환

    tool_name = data.get("tool_name", "")
    if tool_name != "Bash":
        return 0  # 우리는 Bash 호출 후만 처리

    tool_input = data.get("tool_input", {})
    command = tool_input.get("command", "")
    if not matches_fill_command(command):
        return 0  # fill 관련 명령이 아니면 조기 종료

    out_path = extract_output_path(command)
    if not out_path:
        return 0  # 경로 추출 실패면 조용히 종료

    # ~ 확장
    out_path = os.path.expanduser(out_path)

    if not os.path.exists(out_path):
        msg = f"⚠️ HWP fill 결과 파일이 디스크에 없습니다: {out_path}"
        print(json.dumps({"hookSpecificOutput": {"additionalContext": msg}}))
        return 0

    size = os.path.getsize(out_path)
    abs_path = os.path.abspath(out_path)

    # 간단한 무결성 체크: HWP 5.0 binary 매직 시작
    is_hwp5 = False
    try:
        with open(abs_path, "rb") as f:
            magic = f.read(8)
        is_hwp5 = magic == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # CFB 매직
    except OSError:
        pass

    note = (
        f"📄 HWP fill 결과 파일 자동 확인\n"
        f"  경로: {abs_path}\n"
        f"  크기: {size:,} bytes\n"
        f"  HWP 5.0 CFB: {'✓' if is_hwp5 else '? (HWPX 또는 비표준 컨테이너)'}\n"
        f"  사용자에게 한컴/모바일 한글에서 시각 확인을 안내하세요."
    )
    print(json.dumps({"hookSpecificOutput": {"additionalContext": note}}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
