"""SVG 기반 시각 회귀 테스트 — 한컴 없이 Mac 에서 자동 검증.

전략 (char-level 회귀):
  1. 원본 양식(biz_plan.hwp)의 SVG 를 페이지별로 렌더 → 모든 글자 멀티셋 (Counter).
  2. fill_template 으로 자격증 컬럼 4 셀 채움 → filled SVG 페이지별로 렌더 → 글자 멀티셋.
  3. 글자 멀티셋 diff:
     - filled − baseline = 추가된 글자 (정확히 우리 fill 값 합과 일치해야 함)
     - baseline − filled = 사라진 글자 (0 이어야 함)

좌표 비교는 의도적으로 하지 않음 — fill_template 이 셀 한두 개 채울 때 인접 문단의
LineSeg 가 미세 재계산되어 다른 글자들도 sub-px 좌표 변화가 생길 수 있음 (라운드트립 IR
검증은 통과). 시각 회귀의 본질은 "텍스트 내용이 의도대로 변했는가" 이므로 char-level 만
검증해도 충분히 강한 회귀 검출력을 가진다.

이 테스트가 통과하면:
  - 양식의 모든 텍스트 글자가 그대로 보존 (8 개 표 + 본문 + 캡션)
  - fill 동작이 의도한 글자만 정확히 추가
  - rhwp 직렬화 → 재파싱 → 재렌더 사이클 무결

요구사항:
  - rhwp CLI 가 빌드되어 있어야 함 (`../../codebase/rhwp/target/debug/rhwp`).
"""

from __future__ import annotations

import re
import subprocess
from collections import Counter
from pathlib import Path

import pytest

import hwp_automate

REPO_ROOT = Path(__file__).resolve().parent.parent  # hwp-automate-py/
TEMPLATE_PATH = REPO_ROOT / ".." / ".." / "codebase" / "rhwp" / "samples" / "biz_plan.hwp"
RHWP_BIN = REPO_ROOT / ".." / ".." / "codebase" / "rhwp" / "target" / "debug" / "rhwp"

# 채울 값 (test_svg_regression 내에서만 사용)
EXPECTED_FILLS = {
    1: "정보처리기사",
    2: "정보보안기사",
    3: "네트워크관리사",
    4: "컴활 1급",
}

# SVG 의 <text x=".." y=".." ...>X</text> 패턴 — rhwp 는 글자별로 분리 렌더
TEXT_RE = re.compile(
    r'<text\s+x="([0-9.\-]+)"\s+y="([0-9.\-]+)"[^>]*>([^<]+)</text>'
)


def _extract_chars(svg_text: str) -> Counter:
    """SVG 의 모든 <text> 글자 멀티셋 (Counter). 좌표 무시."""
    return Counter(m.group(3) for m in TEXT_RE.finditer(svg_text))


def _render_svg_pages(hwp_path: Path, out_dir: Path) -> list[Path]:
    """rhwp CLI 로 .hwp 파일을 페이지별 SVG 로 출력."""
    out_dir.mkdir(parents=True, exist_ok=True)
    res = subprocess.run(
        [str(RHWP_BIN), "export-svg", str(hwp_path), "-o", str(out_dir)],
        capture_output=True,
        text=True,
        check=True,
    )
    pages = sorted(out_dir.glob("*.svg"))
    assert pages, f"SVG 페이지 없음: {res.stdout}\n{res.stderr}"
    return pages


def _expected_added_chars() -> Counter:
    """EXPECTED_FILLS 의 모든 비공백 글자 멀티셋."""
    out: Counter = Counter()
    for v in EXPECTED_FILLS.values():
        out.update(c for c in v if not c.isspace())
    return out


@pytest.fixture(scope="module")
def baseline_chars(tmp_path_factory) -> Counter:
    """원본 biz_plan.hwp 의 모든 페이지 글자 멀티셋."""
    if not TEMPLATE_PATH.exists():
        pytest.skip(f"템플릿 없음: {TEMPLATE_PATH}")
    if not RHWP_BIN.exists():
        pytest.skip(f"rhwp CLI 없음: {RHWP_BIN}. `cargo build --bin rhwp` 필요.")
    out_dir = tmp_path_factory.mktemp("baseline_svg")
    pages = _render_svg_pages(TEMPLATE_PATH, out_dir)
    counter: Counter = Counter()
    for p in pages:
        counter.update(_extract_chars(p.read_text(encoding="utf-8")))
    return counter


@pytest.fixture(scope="module")
def filled_chars(tmp_path_factory) -> tuple[Counter, Path]:
    """자격증 컬럼 4 셀을 채운 후 SVG 렌더 결과의 글자 멀티셋."""
    if not TEMPLATE_PATH.exists():
        pytest.skip(f"템플릿 없음: {TEMPLATE_PATH}")
    if not RHWP_BIN.exists():
        pytest.skip(f"rhwp CLI 없음: {RHWP_BIN}.")
    work = tmp_path_factory.mktemp("filled")
    filled_hwp = work / "filled.hwp"
    r = hwp_automate.fill_template(
        str(TEMPLATE_PATH),
        str(filled_hwp),
        [{
            "header_match": "성명",
            "column": "자격증",
            "values": EXPECTED_FILLS,
        }],
    )
    assert r["status"] == "applied + verified", r["status"]
    pages = _render_svg_pages(filled_hwp, work / "svg")
    counter: Counter = Counter()
    for p in pages:
        counter.update(_extract_chars(p.read_text(encoding="utf-8")))
    return counter, filled_hwp


def test_no_chars_removed(baseline_chars: Counter, filled_chars: tuple[Counter, Path]):
    """양식의 어떤 글자도 사라지면 안 됨 — 양식 무손상 보장 (가장 중요한 회귀)."""
    filled, _ = filled_chars
    removed = baseline_chars - filled
    assert not removed, (
        f"양식에서 {sum(removed.values())} 개 글자가 사라짐. "
        f"첫 10 개: {list(removed.elements())[:10]}"
    )


def test_added_chars_match_intended(
    baseline_chars: Counter, filled_chars: tuple[Counter, Path]
):
    """추가된 글자의 멀티셋이 우리가 채운 4 개 셀의 글자와 정확히 일치."""
    filled, _ = filled_chars
    added = filled - baseline_chars
    expected = _expected_added_chars()
    assert added == expected, (
        f"추가된 글자 합 ≠ 의도한 글자 합\n"
        f"  추가:        {dict(sorted(added.items()))}\n"
        f"  의도:        {dict(sorted(expected.items()))}\n"
        f"  추가-의도:    {added - expected}\n"
        f"  의도-추가:    {expected - added}"
    )


def test_total_char_count_difference(
    baseline_chars: Counter, filled_chars: tuple[Counter, Path]
):
    """전체 글자 수 차이가 정확히 우리가 채운 비공백 글자 수와 일치."""
    filled, _ = filled_chars
    diff = sum(filled.values()) - sum(baseline_chars.values())
    expected_diff = sum(_expected_added_chars().values())
    assert diff == expected_diff, (
        f"글자 수 차이 {diff} ≠ 의도한 비공백 글자 수 {expected_diff}\n"
        f"baseline={sum(baseline_chars.values())}, filled={sum(filled.values())}"
    )
