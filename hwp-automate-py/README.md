# hwp-automate-py — Python 바인딩 (PyO3)

`hwpx-generator` 의 **경로 B (Rust + rhwp)** 를 Python 에서 호출 가능하게 노출하는 PyO3 + maturin 기반 Python 확장 모듈. abi3-py39 로 빌드되어 **Python 3.9 ~ 3.14** 모두에서 동일한 wheel 사용.

상위 컨텍스트는 `../CLAUDE.md` 의 [Rust + rhwp 경로](../CLAUDE.md#rust--rhwp-경로-크로스플랫폼-com-불필요) 섹션 참조.

## 한 줄 요약

```python
import hwp_automate
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
├── Cargo.toml                  pyo3 0.22 + abi3-py39 + rhwp path 의존
├── pyproject.toml              maturin 빌드 설정
├── src/
│   └── lib.rs                  Rust 확장 모듈 — 3개 함수 노출 (~430줄)
├── hwp_automate_cli/           Python 보조 도구 (wheel 비번들)
│   ├── __init__.py             field_map_to_operations 등 export
│   ├── field_map.py            field_map.json → operations 어댑터
│   └── __main__.py             CLI 진입점 (analyze / fill / cell)
├── .venv/                      격리 venv (gitignore 됨)
├── target/                     cargo build (gitignore 됨)
└── output/                     테스트 출력 (gitignore 됨)
```

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

양식 HWP 의 구조·스타일·표 인벤토리를 dict 로 반환. 부작용 없음.

**리턴 dict 키:**

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

각 `tables[i]` dict:

| 키 | 설명 |
|----|------|
| `section`, `parent_para`, `control` | rhwp 의 위치 식별자 |
| `rows`, `cols` | 표 크기 |
| `header` | 헤더 행 (row 0) 의 셀 텍스트 list (`"|"` 로 다문단 결합) |

**예시:**

```python
info = hwp_automate.analyze_template("biz_plan.hwp")
print(info["style_count"])           # 26
for t in info["tables"]:
    print(f"{t['rows']}x{t['cols']} 헤더={t['header']}")
# 5x6 헤더=['구분', '분야별', '성명', '기술등급', '경력(년)', '자격증']
# ...
```

### `fill_template(template_path, out_path, operations, *, dry_run=False, verify=True) -> dict` ★ (주력)

여러 표·여러 컬럼·여러 셀을 한 번의 호출로 채움. **Pre-flight 검증** (모든 op 가 유효한지 적용 전에 확인) 후 batch 모드로 일괄 적용, **post-fill 검증** (저장한 결과를 재파싱하여 모든 셀 값 보존 확인).

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

**리턴 dict:**

| 키 | 설명 |
|----|------|
| `path`, `bytes` | 출력 경로/크기 (dry_run 이면 bytes=0) |
| `status` | `"applied + verified"` / `"applied (verify=false)"` / `"dry_run"` / 실패 시 에러 |
| `mismatches` | verify 모드에서 라운드트립 불일치 발견 시 메시지 list (정상이면 `[]`) |
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

### `fill_template_table(template_path, out_path, mapping, *, dry_run=False, verify=True) -> dict`

단일 표·단일 컬럼 편의 wrapper. 내부적으로 `fill_template` 한 번 호출. 기존 한 줄 호출 코드 호환성 보존.

```python
hwp_automate.fill_template_table(
    "biz_plan.hwp", "filled.hwp",
    {"header_match": "성명", "column": "자격증",
     "values": {1: "정보처리기사", 2: "정보보안기사"}},
)
```

리턴은 `fill_template` 과 동일.

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
