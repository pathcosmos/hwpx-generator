# hwp-automate-py — Python 바인딩 (PyO3)

`hwpx-generator` 의 **경로 B (Rust + rhwp)** 를 Python 에서 호출 가능하게 노출하는 PyO3 + maturin 기반 Python 확장 모듈. abi3-py39 로 빌드되어 **Python 3.9 ~ 3.14** 모두에서 동일한 wheel 사용.

상위 컨텍스트는 `../CLAUDE.md` 의 [Rust + rhwp 경로](../CLAUDE.md#rust--rhwp-경로-크로스플랫폼-com-불필요) 섹션 참조.

## 한 줄 요약

```python
import hwp_automate

# 분석 — AI 가 양식 의미를 즉시 파악
info = hwp_automate.analyze_template("양식.hwp")
# info["tables"][i]["suggested_fields"] → [{"label":"업종명","row":1,"col":6}, ...]

# 채우기 — 다중 표·다중 셀, dry_run+verify+preserve_images 자동
hwp_automate.fill_template(
    "양식.hwp", "결과.hwp",
    [
        {"header_match": "성명", "column": "자격증",
         "values": {1: "정보처리기사", 2: "정보보안기사"}},
        {"header_match": "성명", "cells": [{"row": 3, "col": 5, "value": "직접지정"}]},
    ],
)
```

## 디렉토리

```
hwp-automate-py/
├── Cargo.toml                   pyo3 0.22 + abi3-py39 + rhwp path 의존
├── pyproject.toml               maturin + mcp optional + pytest config
├── src/
│   └── lib.rs                   ★ Rust 확장 모듈 — 4 함수 노출 (~700줄)
│                                  analyze_template / fill_template /
│                                  fill_template_table / preserve_images_from_source
├── mcp_server.py                ★ FastMCP stdio 서버 (5 tools)
├── hwp_automate_cli/            Python 보조 도구 (wheel 비번들)
│   ├── __init__.py
│   ├── field_map.py             field_map.json → operations 어댑터
│   └── __main__.py              CLI 진입점 (analyze / fill / cell)
├── tests/
│   └── test_svg_regression.py   ★ SVG 글자 멀티셋 회귀 (3 testcase)
├── .venv/                       격리 venv (gitignore)
├── target/                      cargo build (gitignore)
└── output/                      테스트 출력 / V1_*.hwp (gitignore)
```

## API 함수 4 개 + MCP tools 5 개 + CLI 3 명령

| 진입점 | 함수/명령 | 위치 |
|---|---|---|
| Rust 확장 (Python 직접 호출) | `analyze_template`, `fill_template`, `fill_template_table`, `preserve_images_from_source` | `src/lib.rs` |
| MCP 서버 (stdio) | `analyze_form`, `preview_form_structure`, `fill_form`, `fill_form_from_data`, `verify_output` | `mcp_server.py` |
| Python CLI (`python -m hwp_automate_cli`) | `analyze`, `fill`, `cell` | `hwp_automate_cli/__main__.py` |
| (이 저장소 외부) Claude Code Skill | `/fill-hwp 양식.hwp` | `../.claude/skills/fill-hwp/SKILL.md` |

## 사전 요구사항

- **Rust 1.75+** — `brew install rust` (검증: 1.95.0)
- **maturin 1.13+** — `brew install maturin` 또는 venv 안에서 `pip install maturin`
- **Python 3.9+** — abi3 wheel 이라 마이너 버전 무관
- **`../../codebase/rhwp/`** 위치에 [edwardkim/rhwp](https://github.com/edwardkim/rhwp) git clone

## 설치 (빌드 + venv 등록)

### 방법 1 — 격리 venv 에 develop 모드 설치 (개발용, 권장)

```bash
cd hwp-automate-py
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade maturin
maturin develop --release        # ~30초 (rhwp 컴파일 + abi3 wheel + venv 설치)
```

이후 venv 안에서 `import hwp_automate` 즉시 가능.

### 방법 2 — wheel 빌드 후 다른 venv (예: ai-env) 에 설치

```bash
# 1) wheel 빌드 (abi3-py39 = 모든 3.9+ 호환)
cd hwp-automate-py
maturin build --release          # dist/hwp_automate-0.1.0-cp39-abi3-*.whl 생성

# 2) ai-env 같은 다른 venv 에 설치
source /path/to/ai-env/bin/activate
pip install dist/hwp_automate-0.1.0-cp39-abi3-macosx_11_0_arm64.whl
```

> **플랫폼 wheel 명명** — Mac arm64 에서 빌드한 wheel 은 `macosx_11_0_arm64` 태그. Windows/Linux 에서 사용하려면 해당 OS 에서 별도 빌드 필요.

### 방법 3 — pip 가 maturin 호출하도록 (PEP 517)

```bash
# pyproject.toml 의 [build-system] 이 maturin 을 자동 사용
pip install /path/to/hwp-automate-py/
```

## API 레퍼런스

### `analyze_template(path: str) -> dict`

양식 HWP 의 구조·스타일·표 인벤토리·**셀 단위 상세 + 라벨 추론** 을 dict 로 반환. 부작용 없음. AI 가 양식 의미를 즉시 추론할 수 있도록 강화됨.

**리턴 dict 최상위 키:**

| 키 | 타입 | 설명 |
|----|------|------|
| `path` | str | 입력 경로 (그대로 echo) |
| `file_size` | int | 바이트 크기 |
| `section_count` | int | 구역 수 |
| `paragraph_count` | int | 전체 문단 합계 |
| `char_shape_count` | int | CharShape 풀 크기 |
| `para_shape_count` | int | ParaShape 풀 크기 |
| `style_count` | int | Style 풀 크기 |
| `numbering_count` | int | Numbering 정의 수 |
| `bullet_count` | int | Bullet 정의 수 |
| `border_fill_count` | int | BorderFill 정의 수 |
| `style_names` | list[str] | 모든 스타일 한국어 이름 |
| `tables` | list[dict] | 표 인벤토리 (아래) |

**각 `tables[i]` dict:**

| 키 | 타입 | 설명 |
|----|------|------|
| `section`, `parent_para`, `control` | int | rhwp 의 위치 식별자 |
| `rows`, `cols` | int | 표 행/열 수 |
| `header` | list[str] | 헤더 행 (row 0) 의 셀 텍스트 list (`"|"` 로 다문단 결합) |
| `cells` | list[dict] | **★ NEW** 모든 셀 (rows×cols 가 아니라 병합 후 실제 셀 수) |
| `empty_cells` | list[dict] | **★ NEW** 빈 셀만 + neighbor_label 추론 |
| `suggested_fields` | list[dict] | **★ NEW** 라벨 추론 성공한 빈 셀 — AI 가 사용자에게 물어볼 후보 |

**`cells[i]` 형식:**

| 키 | 타입 | 설명 |
|----|------|------|
| `row`, `col` | int | 셀 좌표 (0-indexed) |
| `text` | str | 셀 텍스트 (다문단은 `"|"` 결합) |
| `is_empty` | bool | `text.trim()` 이 비었는지 |
| `neighbor_label` | str (선택) | `is_empty=True` 일 때 인접 셀에서 추론한 라벨. 같은 행 왼쪽 우선, 같은 열 위쪽 차순위. |

**`empty_cells[i]` 형식:**

| 키 | 타입 | 설명 |
|----|------|------|
| `row`, `col` | int | 빈 셀 좌표 |
| `neighbor_label` | str (선택) | 라벨 추론 결과 (실패 시 키 없음) |

**`suggested_fields[i]` 형식 — `empty_cells` 중 라벨 추론 성공한 항목만:**

| 키 | 타입 | 설명 |
|----|------|------|
| `label` | str | 추론된 라벨 (사용자에게 물어볼 친화적 한국어) |
| `row`, `col` | int | 채울 셀 좌표 |

**예시 — 실 양식 (YCP 사업신청서) 의 기본정보 표:**

```python
info = hwp_automate.analyze_template("/path/to/사업신청서.hwp")

basic_info = next(t for t in info["tables"] if t["parent_para"] == 6)
print(f"표 크기: {basic_info['rows']}x{basic_info['cols']}")
# 7x8

print(f"빈 셀 수: {len(basic_info['empty_cells'])}")
# 6

for sf in basic_info["suggested_fields"]:
    print(f"  '{sf['label']}' → (row={sf['row']}, col={sf['col']})")
# '업종명' → (row=1, col=6)
# '주생산품' → (row=2, col=6)
# '매출액(백만원)' → (row=4, col=2)
# '영업이익(백만원)' → (row=4, col=3)
# '수출액(백만원)' → (row=4, col=5)
# '부채비율(%)' → (row=4, col=7)
```

이 결과만으로 AI 는 사용자에게 "**업종명, 주생산품, 매출액(백만원), 영업이익(백만원), 수출액(백만원), 부채비율(%) 을 알려주세요**" 라고 자연스러운 질문을 만들 수 있다.

**라벨 추론 규칙 (`find_neighbor_label`):**
1. **1순위 — 같은 행 왼쪽 셀** (한국 양식의 라벨-값 가로 페어 패턴: "기업명 | (값)")
2. **2순위 — 같은 열 위쪽 셀** (헤더 행 또는 위 라벨)
3. 둘 다 비어 있으면 추론 실패 → `neighbor_label` 키 없음, `suggested_fields` 에서 제외

### `fill_template(template_path, out_path, operations, *, dry_run=False, verify=True, preserve_images=True) -> dict` ★ (주력)

여러 표·여러 컬럼·여러 셀을 한 번의 호출로 채움. **Pre-flight 검증** (모든 op 유효성 확인) → batch 적용 → **post-fill 검증** (라운드트립 재파싱) → **BinData/Preview 보존** (rhwp 라운드트립 손실 우회) 의 4 단계 자동.

**`operations` 는 dict 의 list. 각 dict 형식:**

표 식별 (둘 중 하나 필수):
- `"header_match": str` — 대상 표의 헤더 행(row 0) 어딘가에 포함되어야 할 텍스트
- `"table_at": [sec, para, ctrl]` — 직접 좌표

채우기 모드 (둘 중 하나 필수):
- `"column": str` + `"values": {row: str}` — 컬럼 헤더로 위치 찾고 행별 값
- `"cells": [{"row": int, "col": int, "value": str}, ...]` — (row, col) 직접 지정

**옵션:**
- `dry_run=True` — 실제 적용·저장 없이 plan 만 검증·반환 (status="dry_run", bytes=0)
- `verify=False` — 저장 후 라운드트립 검증 생략 (기본 verify=True 권장)
- `preserve_images=True` (기본) — 원본 양식의 BinData/Preview stream 을 byte-for-byte 보존하여 rhwp v0.7.x 의 이미지 라운드트립 손실 회피. **반드시 true 유지** (false 시 한컴이 양식을 "손상"으로 판정하거나 그림 일부 누락 가능). 이미지를 추가/삭제하는 변경이 아니면 안전.

**리턴 dict:**

| 키 | 설명 |
|----|------|
| `path`, `bytes` | 출력 경로/크기 (dry_run 이면 bytes=0) |
| `status` | `"applied + verified"` / `"applied (verify=false)"` / `"dry_run"` / 실패 시 에러 |
| `mismatches` | verify 모드에서 라운드트립 불일치 발견 시 메시지 list (정상이면 `[]`) |
| `preserved_streams` | preserve_images=True 일 때 입력에서 보존된 stream 수 (BinData/Preview/etc 합) |
| `operations` | input op 별 `{table: {section, parent_para, control, rows, cols}, applied: [{row, col, value}]}` |

**예시 — 다중 op:**

```python
r = hwp_automate.fill_template(
    "biz_plan.hwp",
    "filled.hwp",
    [
        {"header_match": "성명", "column": "자격증",
         "values": {1: "정보처리기사", 2: "정보보안기사"}},
        {"header_match": "성명",
         "cells": [{"row": 3, "col": 4, "value": "5년"}]},
    ],
)
assert r["status"] == "applied + verified"
```

**예시 — dry_run 으로 사전 검증만:**

```python
r = hwp_automate.fill_template(
    "biz_plan.hwp", "ignored.hwp",
    [{"header_match": "없는헤더", "column": "X", "values": {1: "Y"}}],
    dry_run=True,
)  # ValueError: operations[0].header_match='없는헤더' 으로 표를 찾을 수 없음
```

**예외:**

- `KeyError` — operation 에 필수 키 누락 (`header_match`/`table_at`, `column`/`cells`, `value` 등)
- `TypeError` — operations 또는 그 안 dict 가 잘못된 타입
- `ValueError` — pre-flight 단계 매칭/범위 실패 (양식 무수정으로 안전)
- `RuntimeError` — `verify=True` 인데 라운드트립 불일치 발견 (drift 감지)
- `IOError` — 파일 read/write 실패

### `fill_template_table(template_path, out_path, mapping, *, dry_run=False, verify=True, preserve_images=True) -> dict`

단일 표·단일 컬럼 편의 wrapper. 내부적으로 `fill_template` 한 번 호출. 기존 한 줄 호출 코드 호환성 보존.

```python
hwp_automate.fill_template_table(
    "biz_plan.hwp", "filled.hwp",
    {"header_match": "성명", "column": "자격증",
     "values": {1: "정보처리기사", 2: "정보보안기사"}},
)
```

리턴은 `fill_template` 과 동일.

### `preserve_images_from_source(source_hwp, target_hwp, out_hwp) -> int`

`source_hwp` 의 BinData/Preview 등 stream 을 byte-for-byte 보존하면서 `target_hwp` 의 BodyText/DocInfo/FileHeader 만 살린 새 CFB 를 `out_hwp` 로 저장.

`fill_template` 의 `preserve_images=True` 가 자동 호출하지만, 별도 후처리가 필요할 때 단독 사용 가능.

```python
hwp_automate.preserve_images_from_source(
    source_hwp="원본.hwp",       # 이미지 출처
    target_hwp="rhwp_출력.hwp",  # rhwp 가 만든 셀 변경된 출력
    out_hwp="최종.hwp",          # 머지된 결과
)
# → 보존된 stream 수 (int)
```

내부 동작: rhwp 의 `LenientCfbReader` 로 양 CFB 를 read, `mini_cfb::build_cfb` 로 새 컨테이너 작성. 외부 `cfb` crate 의 strict 검증 회피.

## MCP 서버 (`mcp_server.py`)

Claude Desktop / Claude Code / Cursor / 기타 MCP 호환 클라이언트가 자연어로 본 라이브러리를 사용할 수 있게 하는 FastMCP stdio 서버.

### 설치

```bash
cd hwp-automate-py
source .venv/bin/activate
pip install --upgrade 'mcp[cli]>=1.2.0'   # Python 3.10+ 필요
maturin develop --release                   # rhwp wheel 설치 (이미 했다면 건너뜀)
```

### 실행

```bash
# 직접 실행 (테스트용)
python mcp_server.py

# 또는 클라이언트가 자동 spawn (Claude Desktop config 등록 후)
```

### 노출 5 tools

| tool | 인자 | 용도 |
|------|------|------|
| `analyze_form` | `template_path` | 양식 구조·빈 셀·라벨 추론 (analyze_template 그대로) |
| `preview_form_structure` | `template_path` | 가벼운 markdown 요약 (큰 양식 첫 검토 시 컨텍스트 절약) |
| `fill_form` | `template_path, output_path, operations, dry_run, verify, preserve_images` | 양식 채우기 (fill_template 그대로) |
| `fill_form_from_data` | `template_path, output_path, field_map_path, data_path, table_locator, dry_run, skip_empty` | field_map.json + data.json 형식 자동 변환 후 채우기 |
| `verify_output` | `output_path, expected_cells` | 결과 파일 셀 라운드트립 검증 |

### Claude Desktop 등록

`~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) — 절대 경로 필수:

```json
{
  "mcpServers": {
    "hwp-automate": {
      "command": "/abs/path/to/hwp-automate-py/.venv/bin/python",
      "args": ["/abs/path/to/hwp-automate-py/mcp_server.py"]
    }
  }
}
```

### Claude Code 등록

```bash
claude mcp add hwp-automate -- /abs/path/to/.venv/bin/python /abs/path/to/mcp_server.py
```

또는 프로젝트 루트의 `.mcp.json` 파일에:

```json
{
  "mcpServers": {
    "hwp-automate": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "--directory", "/abs/path/to/hwp-automate-py", "python", "mcp_server.py"]
    }
  }
}
```

### 사용자 시나리오 (자연어 → MCP tool 호출)

```
사용자: /Users/me/사업신청서.hwp 양식 분석하고 우리 회사 데이터로 채워줘.
        회사명=테크스타트, 대표자=김철수

Claude: [analyze_form 호출]
        → 16 표 발견, 표 [2] 기본정보 의 빈 셀 6개 발견:
          업종명, 주생산품, 매출액(백만원), 영업이익(백만원), 수출액(백만원), 부채비율(%)

        업종명·주생산품·재무현황 4개 항목도 알려주세요. (회사명·대표자는 이미 채워져 있음)

사용자: 철강 특수강 제조 / 스테인리스 / 12500 / 1200 / 8300 / 45.2

Claude: [fill_form 호출 with operations] → status: applied + verified, 6 셀 적용
        결과: /Users/me/사업신청서_filled.hwp (35MB)
        한컴/모바일 한글에서 열어 확인하세요.
```

### Smoke test (in-process)

```bash
cd hwp-automate-py && source .venv/bin/activate
python -c "
import asyncio
from mcp_server import mcp
async def t():
    tools = await mcp.list_tools()
    print(f'tools: {len(tools)}')
    for x in tools: print(f'  - {x.name}')
asyncio.run(t())
"
# tools: 5
#   - analyze_form
#   - preview_form_structure
#   - fill_form
#   - fill_form_from_data
#   - verify_output
```

### 보안 / 안정성

- **stdout 오염 방지** — 모든 진단 로그를 `sys.stderr` 로. import 시점에는 침묵.
- **파일 경로 기반 설계** — 40MB+ 큰 HWP 도 메시지 크기 제약 없음 (path 만 전달).
- **Pre-flight 검증** — 잘못된 fill 요청은 양식 무수정으로 안전하게 거부.
- **Post-fill verify** — `preserve_images=True` 와 결합해 한컴이 받는 valid HWP 5.0 보장.

## CLI — `python -m hwp_automate_cli`

`hwp_automate_cli` 보조 모듈은 wheel 에 번들되지 않는다 (소스 트리에서만 사용). 사용자 양식·field_map.json·data 를 인자로 받아 `fill_template` 을 호출하는 명령행 래퍼다.

### `analyze` — 양식 인벤토리

```bash
python -m hwp_automate_cli analyze --template ./양식.hwp
python -m hwp_automate_cli analyze --template ./양식.hwp --json    # 기계판독용
```

### `fill` — field_map.json + data 로 일괄 채우기

```bash
python -m hwp_automate_cli fill \
  --template ./양식.hwp \
  --field-map ./templates/cloud_integrated/field_map.json \
  --data ./data/sample_input.json \
  --output ./out/채워진.hwp \
  --header-match "성명"          # 또는 --table-at 0 70 0

# 옵션
  --dry-run              # 검증만, 파일 안 만듦
  --no-verify            # 라운드트립 검증 생략
  --print-operations     # 변환된 operations JSON 표시
  --include-empty        # 빈 값도 셀에 쓰기 (기본은 빈 값 셀 제외)
```

### `cell` — 표 1개 셀 몇 개 빠르게

```bash
python -m hwp_automate_cli cell \
  --template ./양식.hwp \
  --output ./out/quick.hwp \
  --header-match "성명" \
  --cell 1,5,정보처리기사 \
  --cell 2,5,정보보안기사
```

### `field_map.json` 어댑터 직접 사용 (Python)

```python
import json
from hwp_automate_cli import field_map_to_operations
import hwp_automate

field_map = json.load(open("templates/cloud_integrated/field_map.json"))
data = json.load(open("data/sample_input.json"))

ops = field_map_to_operations(
    field_map, data,
    table_locator={"header_match": "표 헤더 텍스트"},
    skip_empty=True,
)
hwp_automate.fill_template("양식.hwp", "결과.hwp", ops)
```

`entity_blocks` (offset+left/right 2열 블록) 와 `company_lists` (행 단위 리스트) 두 형식이 모두 cells 모드 operation 으로 변환됨.

## 동작 검증 — 빠른 회귀

```bash
source .venv/bin/activate
python3 - <<'PY'
import hwp_automate

# 1. analyze
info = hwp_automate.analyze_template("../../codebase/rhwp/samples/biz_plan.hwp")
assert info["style_count"] >= 20
assert any(t["rows"] == 5 and t["cols"] == 6 for t in info["tables"])

# 2. fill_template (다중 op + verify)
r = hwp_automate.fill_template(
    "../../codebase/rhwp/samples/biz_plan.hwp",
    "output/regression.hwp",
    [{"header_match": "성명", "column": "자격증",
      "values": {1: "테A", 2: "테B"}}],
)
assert r["status"] == "applied + verified", r["status"]
assert len(r["operations"][0]["applied"]) == 2

# 3. dry_run 안전성
r = hwp_automate.fill_template(
    "../../codebase/rhwp/samples/biz_plan.hwp",
    "output/dry.hwp",
    [{"header_match": "성명", "column": "자격증", "values": {1: "X"}}],
    dry_run=True,
)
assert r["status"] == "dry_run"
assert r["bytes"] == 0

# 4. pre-flight 보호 (잘못된 컬럼 → 양식 무수정)
import os
fail_path = "output/should_not_exist.hwp"
if os.path.exists(fail_path):
    os.remove(fail_path)
try:
    hwp_automate.fill_template(
        "../../codebase/rhwp/samples/biz_plan.hwp", fail_path,
        [{"header_match": "성명", "column": "없는컬럼", "values": {1: "X"}}],
    )
except ValueError:
    pass
assert not os.path.exists(fail_path), "pre-flight 실패인데 파일 만들어짐"

print("✅ 4단계 회귀 모두 통과")
PY
```

## 다른 venv (ai-env 등) 에 통합

`ai-env` 같은 공통 venv 에 wheel 을 설치해두면, 다른 자동화 스크립트가 곧장 `import hwp_automate` 로 사용 가능.

```bash
# 1) 한 번 빌드
cd hwp-automate-py && maturin build --release

# 2) ai-env 에 설치
source /path/to/ai-env/bin/activate
pip install --force-reinstall dist/hwp_automate-0.1.0-cp39-abi3-*.whl

# 3) ai-env 어디서든 사용
python3 -c "import hwp_automate; print(hwp_automate.__doc__)"
```

추후 rhwp / 본 모듈을 업데이트했다면 `maturin build` → `pip install --force-reinstall` 반복.

## 개발 사이클 (라이브러리 수정 시)

```bash
cd hwp-automate-py
source .venv/bin/activate
# src/lib.rs 수정 후
maturin develop --release    # ~5초 (rhwp 캐시 활용)
python3 -c "import hwp_automate; ..."   # 즉시 반영 확인
```

## 기존 경로 A 와의 입력 형식 대조

기존 `templates/<name>/field_map.json` 은 행 단위 entity_blocks 형식이고, 본 모듈의 `fill_template_table` 은 컬럼 단위 `header_match` + `column` + `values` 형식. 두 형식은 의미상 **일대일 매핑 가능** — 향후 어댑터 (`field_map.json` → `mapping` dict) 한 함수로 두 자동화 경로의 입력을 통일 가능.

## 한계

- 현재 `fill_template_table` 은 **한 번에 한 컬럼**만 채움. 여러 컬럼 동시 채움은 호출 반복 또는 향후 확장으로 처리.
- **컬럼 헤더가 정확히 일치**해야 함 (부분 매치 미지원). 향후 옵션 추가 가능.
- 표 행 추가/삭제·셀 병합·스타일 적용 등은 본 함수에 없음. 필요하면 lib.rs 에 추가 함수 노출.
- macOS arm64 wheel 은 Mac 만, Windows/Linux 는 별도 빌드 필요 (abi3 라 Python 버전은 무관).

## Acknowledgement (감사·출처)

본 Python 바인딩은 두 개의 외부 오픈소스 프로젝트 위에 만들어졌습니다. HWP 처리의 핵심 가치는 모두 다음 프로젝트들의 결과물이며, 본 모듈은 그 위에 PyO3 + Python 편의 레이어를 더한 얇은 래퍼입니다.

### 🦀 rhwp — 핵심 엔진 (직접 의존)

- **저자:** Edward Kim ([@edwardkim](https://github.com/edwardkim))
- **저장소:** https://github.com/edwardkim/rhwp
- **라이선스:** MIT
- **설명:** Rust + WebAssembly 기반 오픈소스 HWP/HWPX 뷰어/에디터. v0.7.x 시점 891+ 테스트, hyper-waterfall 방법론으로 개발.

**본 모듈이 사용하는 rhwp 의 모듈·기능:**

| rhwp 모듈/기능 | 본 모듈이 사용하는 방법 |
|---|---|
| `rhwp::document_core::DocumentCore` | `analyze_template`, `fill_template` 의 모든 IR 조작 — `from_bytes`, `create_table_native`, `insert_text_in_cell_native`, `apply_cell_style_native`, `apply_style_native`, `split_paragraph_native`, `begin_batch_native`, `end_batch_native`, `export_hwp_native` 등 |
| `rhwp::parser::parse_document` | HWP 5.0 / HWPX 자동 포맷 감지 파싱 |
| `rhwp::error::HwpError` | 에러 타입 — Python `RuntimeError("HwpError: ...")` 로 매핑 |
| `rhwp::model::control::Control::Table`, `model::table::{Table, Cell}` | 표 IR — `discover_all_tables` 가 직접 순회하여 표 인벤토리 구축 |
| `rhwp::parser::cfb_reader::LenientCfbReader` | 비표준 CFB 메타 가진 양식의 lenient 파싱 — `merge_cfb_preserving_input` 에서 BinData 보존 우회용 |
| `rhwp::serializer::mini_cfb::build_cfb` | rhwp 의 자체 CFB v3 writer — `merge_cfb_preserving_input` 결과를 한컴 호환 CFB 로 작성 |

**의존 방식:** 본 모듈의 `Cargo.toml` 에 `rhwp = { path = "../../codebase/rhwp", default-features = false }` 로 path 의존. **rhwp 코드는 일절 수정하지 않음** — upstream 그대로 사용.

**왜 rhwp 인가:** Mac/Linux/Windows 어디서든 한컴오피스 설치 없이 .hwp 처리가 가능한 유일한 성숙한 오픈소스 엔진. 본 모듈의 모든 신뢰성(라운드트립 무결성, 한컴 호환 출력, IR 정확도)은 rhwp 의 결과물입니다.

**우회 우리 코드:** rhwp v0.7.x 의 BinData (이미지 BMP 등) 라운드트립 충실도 한계가 있어 (이미지가 미세하게 변형되거나 한컴이 "손상" 으로 판정 가능), 본 모듈은 `LenientCfbReader` 와 `mini_cfb` 를 사용한 stream-level 머지 우회법(`preserve_images=True`, 기본 활성)으로 BinData/Preview 를 입력 양식에서 byte-for-byte 보존합니다. 이 우회는 rhwp 의 lenient/builder 모듈을 그대로 활용하며, rhwp 자체의 개선이 이루어지면 자연스럽게 단순해질 수 있습니다.

### 🪝 hop — 패턴 출처 (참고만, 직접 의존 안 함)

- **저자:** golbin ([@golbin](https://github.com/golbin))
- **저장소:** https://github.com/golbin/hop
- **라이선스:** MIT
- **설명:** Tauri 2 기반 macOS/Windows/Linux 데스크톱 HWP 뷰어·에디터. rhwp 를 third_party 서브모듈로 통합한 운영 환경 사례.

**본 모듈이 hop 에서 흡수한 설계 패턴:**

| hop 의 코드 | 본 모듈이 영향받은 부분 |
|---|---|
| `apps/desktop/src-tauri/src/state.rs::editable_core_from_bytes` | `analyze_template` / `fill_template` 의 표준 진입점 — `DocumentCore::from_bytes(bytes)` 한 줄로 양식 로드 |
| `commands.rs::mutate_document(doc_id, operation, args, ...)` JSON dispatcher | `fill_template(operations: list[dict])` 의 다중 op 디자인 — 한 호출에 여러 fill 작업 묶기 |
| `commands.rs::query_document(doc_id, query, args)` JSON 쿼리 dispatcher | `analyze_template` 의 read-only 인벤토리 반환 패턴 |
| Hop 의 `expected_revision` 동시성 보호 | 향후 동시 편집 안전망 (현재 미구현, 패턴만 인지) |
| `apps/desktop/src-tauri/src/commands.rs::commit_staged_hwp_save` 의 atomic save | 운영 자동화 시 사용자 원본 파일 보호 (현재 본 모듈은 별도 출력 경로 사용으로 우회) |

**의존 방식:** **직접 코드 의존 없음.** hop 의 코드를 읽고 좋은 패턴을 본 모듈 설계에 흡수. hop 자체는 GUI 데스크톱 앱이라 우리 Python 헤드리스 자동화와 다른 영역.

**왜 hop 패턴인가:** rhwp 를 production 환경에서 어떻게 쓰는지 보여주는 가장 완성도 높은 사례. 그 코드를 읽으며 "DocumentCore.document 는 `pub(crate)` — 외부에선 만지지 말 것", "JSON 디스패처로 다중 op 를 한 트랜잭션으로", "외부 수정 감지 후 atomic rename" 같은 운영 정책을 자연스럽게 채택했습니다.

### ⚙️ Python · Rust 빌드 라이브러리

| 라이브러리 | 용도 | 라이선스 |
|---|---|---|
| [PyO3](https://github.com/PyO3/pyo3) | Rust ↔ Python FFI 바인딩 | Apache-2.0 / MIT |
| [maturin](https://github.com/PyO3/maturin) | PyO3 wheel 빌드 | MIT |
| [pytest](https://github.com/pytest-dev/pytest) | 테스트 (`tests/test_svg_regression.py`) | MIT |
| Rust toolchain (cargo, rustc) | 컴파일 | MIT/Apache-2.0 |

### 한글 / 한컴 상표 안내

- **"한글", "한컴", "HWP", "HWPX"** 는 주식회사 한글과컴퓨터의 등록 상표입니다.
- 본 프로젝트는 한글과컴퓨터와 제휴, 후원, 승인 관계가 없는 **독립적인 오픈소스 작업**입니다.
- HWP 5.0 바이너리 포맷 처리는 rhwp 가 한글과컴퓨터의 공개 문서를 참고하여 구현한 결과를 활용합니다.

### 본 모듈의 위치·범위

- **목적:** rhwp 엔진을 Python 에서 호출 가능하게 노출 + 양식 자동 채우기 편의 API
- **상태:** PoC 단계 (proof-of-concept). 사용자 내부 업무 자동화 목적
- **외부 배포·재배포 시:** rhwp / hop / Hangul/Hancom 상표 표기 의무 준수 필수
- **wheel 의 abi3 보장 범위:** Python 3.9 ~ 3.14 동일 wheel 호환 (각 OS 별 별도 빌드 필요)
