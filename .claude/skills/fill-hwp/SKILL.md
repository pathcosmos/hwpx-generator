---
name: fill-hwp
description: HWP/HWPX 양식 파일을 AI 가 분석하여 사용자가 어떤 정보를 입력해야 할지 제안하고, 콘텐츠를 생성·삽입하여 결과 파일을 만드는 자동화. 사용자가 "양식 채워줘", "HWP 작성", "사업신청서 작성" 같은 요청을 하거나 .hwp/.hwpx 파일을 가리키며 채우기를 원할 때 사용.
---

# HWP 양식 자동 채우기 — playbook

이 skill 은 사용자의 HWP/HWPX 양식 파일을 분석해 빈 셀과 라벨을 식별하고, 사용자와 멀티턴 대화로 필요한 정보를 수집하여 양식을 채운다. rhwp 기반 `hwp_automate` Python 모듈을 활용.

## 환경 준비 (한 번)

이 프로젝트의 hwp-automate-py venv 를 사용한다:

```
HWP_VENV=/Volumes/minim42tbtmm/temp_git/hwpx-generator/hwp-automate-py/.venv
```

venv 가 없거나 hwp_automate import 실패 시:

```bash
cd /Volumes/minim42tbtmm/temp_git/hwpx-generator/hwp-automate-py
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade maturin
maturin develop --release
```

## 단계 1 — 양식 파일 특정

사용자가 `/fill-hwp` 인자로 경로를 주거나 메시지에 .hwp 경로가 있으면 그것을 사용. 없으면 현재 작업 폴더와 ~/Downloads 에서 .hwp/.hwpx 파일 목록을 보여주고 사용자에게 선택 요청.

```bash
# 후보 탐색
ls *.hwp *.hwpx 2>/dev/null
ls ~/Downloads/*.hwp 2>/dev/null | head -5
```

## 단계 2 — 양식 구조 분석 (AI 가 의미 추론하는 핵심)

```bash
source $HWP_VENV/bin/activate
python -m hwp_automate_cli analyze --template "$TEMPLATE" --json > /tmp/hwp_analyze.json
```

또는 inline Python:

```python
import hwp_automate, json
info = hwp_automate.analyze_template("/path/to/form.hwp")
print(json.dumps(info, ensure_ascii=False, indent=2)[:3000])  # 큰 양식은 잘라서
```

**해석 가이드** (이게 핵심):

`info["tables"]` 의 각 표 entry 에:
- `header`: row 0 텍스트 list — 표의 의미 식별 ("성명", "기업명" 등이 보이면 어떤 표인지 추정 가능)
- `cells`: 모든 셀 (row, col, text, is_empty, neighbor_label)
- `empty_cells`: 빈 셀만 + neighbor_label (라벨 추론) — **사용자에게 물어볼 후보**
- `suggested_fields`: 라벨 추론 성공한 빈 셀 — `[{"label":"매출액(백만원)","row":4,"col":2}, ...]` 형태

각 표의 `suggested_fields` 가 곧 "이 양식에 채워야 할 항목 목록" 이다.

## 단계 3 — 채울 내용 파악 (사용자에게 한 번에 묻기)

분석 결과의 `suggested_fields` 를 종합해 사용자에게 친화적으로 정리해서 묻는다.

**좋은 예** (한 번에 묶어서):
```
이 양식 분석 결과 다음 8개 항목이 필요합니다:

표 1 (기본정보):
  • 업종명 / 주생산품 / 매출액 / 영업이익 / 수출액 / 부채비율

표 2 (인력 명단, 5x6):
  • (헤더에 "성명", "자격증" 컬럼 있음 — 인원 수와 각자 정보 알려주세요)

회사 기본정보를 알려주시거나, "회사 프로필 파일 있어" 같이 답해 주세요.
```

**원칙**:
- `empty_cells.neighbor_label` 그대로 사용자 친화적 한국어 — 추가 번역 불필요
- 사용자가 "큰 틀만" 주는 시나리오: 자세한 콘텐츠는 LLM(Claude 본인)이 컨텍스트로 생성. 사용자에게 "이런 회사 / 이런 사업" 정도만 묻고 나머지는 작성.
- 사용자가 회사 프로필 JSON 을 가지고 있으면 받아서 매핑.

## 단계 4 — 채우기 실행 (dry_run 먼저, 그 다음 실제)

`hwp_automate.fill_template` 호출. operations 는 표별로 분리:

```python
import hwp_automate

ops = [
    {
        "table_at": [0, 6, 0],  # 표 1 (기본정보) 좌표 — analyze 결과에서 가져옴
        "cells": [
            {"row": 1, "col": 6, "value": "철강 특수강 제조"},  # 업종명
            {"row": 4, "col": 2, "value": "12,500"},           # 매출액
            # ...
        ],
    },
    {
        "header_match": "성명",  # 표 2 (인력명단) — 헤더로 식별
        "cells": [
            {"row": 1, "col": 5, "value": "정보처리기사"},
            # ...
        ],
    },
]

# Dry run 먼저
result = hwp_automate.fill_template(
    template_path="/path/to/form.hwp",
    out_path="/path/to/output.hwp",
    operations=ops,
    dry_run=True,
)
print(f"Dry run: {result['status']}")  # "dry_run"

# 사용자에게 plan 확인 받은 후 실제 실행
result = hwp_automate.fill_template(
    template_path="/path/to/form.hwp",
    out_path="/path/to/output.hwp",
    operations=ops,
    # dry_run=False (default), verify=True (default), preserve_images=True (default)
)
print(f"Final: {result['status']}")  # "applied + verified"
```

**중요 옵션**:
- `dry_run=True` — 실제 적용·저장 없이 plan 만 검증. AI 가 plan 이상 없는지 먼저 확인.
- `verify=True` (기본) — 저장 후 라운드트립 검증 자동.
- `preserve_images=True` (기본) — 원본의 이미지(BinData/Preview) 그대로 보존. **반드시 true 유지** (false 시 한컴이 손상으로 인식할 수 있음).

## 단계 5 — 결과 보고

`fill_template` 결과 dict 를 사용자 친화적으로 요약:

```
✅ 양식 채우기 완료 (applied + verified)
  • 출력: /path/to/output.hwp (28,160 bytes)
  • 채운 셀: 6개 (표 기본정보 6/6)
  • 검증: 라운드트립 완전 일치, mismatches 없음

다음으로 한컴/모바일 한글에서 출력 파일을 열어 확인해 주세요.
```

## 오류 처리

| 증상 | 원인 | 대응 |
|---|---|---|
| `ValueError: header_match='X' 으로 표를 찾을 수 없음` | analyze 결과의 헤더 텍스트와 다름 | analyze 결과 다시 보여주고 사용자에게 정확한 헤더 확인 |
| `ValueError: (row=N, col=M) 위치에 셀이 없음 (병합되었거나 범위 외)` | 셀 병합으로 그 위치가 표에 없음 | analyze 의 `cells` 목록에서 실제 존재하는 좌표 확인 후 재시도 |
| `RuntimeError: post-fill 검증 실패` | 라운드트립 mismatch (드뭄) | mismatches 메시지 확인. 보통 trailing 공백 차이라 `verify=False` 로 우회 가능 (단 결과 무결성은 사용자 본인 확인 필요) |
| `ImportError: hwp_automate` | venv 활성화 안 됨 | 환경 준비 절 명령으로 venv 재구성 |

## 진행 흐름 요약

```
사용자: "사업계획서양식.hwp 채워줘"
  ↓
[analyze] → 16 표, suggested_fields 종합
  ↓
AI: "다음 정보가 필요합니다: 업종명, 주생산품, 매출액, ..."
  ↓
사용자: "철강 특수강 제조, 스테인리스, 12500..."
  ↓
[fill dry_run] → plan 검증 OK
  ↓
[fill 실제] → applied + verified
  ↓
AI: "✅ 완료. /tmp/사업계획서양식_filled.hwp 생성"
```

## 핵심 원칙

1. **사용자 양식 무손상 보장** — 항상 새 출력 파일에 저장, 원본 절대 덮어쓰기 금지.
2. **분석 → 제안 → 확인 → 실행** 순서 — 사용자 확인 없이 fill 적용하지 말 것.
3. **dry_run 활용** — 큰 양식은 plan 먼저 보여주고 적용.
4. **빈 셀에만 채움** — `is_empty=True` 인 셀만 대상. 기존 채워진 셀은 사용자가 명시적으로 "수정해" 라고 할 때만.
5. **시각 검증 한계 인지** — rhwp v0.7.x 의 SVG 렌더는 outline 자동 번호를 안 그림. 한컴/모바일 한글에서 최종 확인 안내.
