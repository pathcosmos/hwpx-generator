# CLAUDE.md

이 파일은 Claude Code 및 AI 에이전트가 이 저장소에서 작업할 때 반드시 숙지해야 할 사항을 정리한다.

## 프로젝트 개요

한컴오피스 한글(Hangul)의 XML 기반 문서 형식인 HWPX 파일을 프로그래밍 방식으로 생성하는 도구. 정부 사업계획서, 주간/월간 보고서 등 복잡한 양식 문서의 자동 생성을 목표로 한다.

**핵심 전략**: 두 자동화 경로를 보유하며 시나리오에 따라 선택한다.

| 경로 | 위치 | 환경 | 강점 |
|------|------|------|------|
| **A. Python + lxml + COM** (기존 하이브리드) | 루트 + `src/` | WSL/Ubuntu + Windows 한컴 | HWPX XML 직접 편집, 한컴 PDF 변환 충실도 |
| **B. Rust + rhwp** (크로스플랫폼, COM 불필요) | `hwp-automate-poc/`, `hwp-automate-py/` | macOS/Linux/Windows 모두 | HWP 5.0 binary 처리, 한컴 설치 없이 동작, IR 안전 변형 |

경로 B 의 상세는 본 문서 [Rust + rhwp 경로](#rust--rhwp-경로-크로스플랫폼-com-불필요) 절 참고.

## 실행 환경

| 구성 | 상세 |
|------|------|
| WSL | Ubuntu, Python 3.12.3 (`--break-system-packages` 필요, venv 미설치) |
| Windows | Python 3.13.12 (`/mnt/c/Users/lanco/AppData/Local/Microsoft/WindowsApps/python.exe`) |
| 한컴오피스 | 2024 (Windows) |
| WSL 패키지 | `lxml`, `pymupdf`, `Pillow`, `scikit-image`, `numpy` |
| Windows 패키지 | `pywin32` |

## 아키텍처 & 모듈

```
JSON Input → field_mapper.py → hwpx_editor.py (XML 직접 수정, WSL)
                                     ↓
                                bridge.py (WSL→Windows 브릿지)
                                     ↓
                                hwp_com.py (COM 텍스트 교체 + PDF 저장, Windows)
                                     ↓
                                pdf_compare.py (SSIM 검증, WSL)
```

| 모듈 | 환경 | 역할 |
|------|------|------|
| `src/generate_hwpx.py` | WSL | 메인 CLI 파이프라인 |
| `src/bridge.py` | WSL | WSL↔Windows 브릿지. 인라인 스크립트를 Windows Python으로 실행 |
| `src/hwp_com.py` | **Windows only** | 한컴오피스 COM 자동화 (pywin32). WSL에서 import 불가 |
| `src/hwpx_editor.py` | WSL | HWPX ZIP 내부 section0.xml 수정 (lxml) |
| `src/field_mapper.py` | WSL | JSON 입력 → (row, col) 셀 좌표 매핑 |
| `src/pdf_compare.py` | WSL | PDF SSIM + 텍스트 비교 |
| `src/extract_template.py` | WSL | HWPX 구조 분석 + 템플릿 설정 초안 생성 |

## Rust + rhwp 경로 (크로스플랫폼, COM 불필요)

기존 Python+lxml+COM 하이브리드와 **나란히** 동작하는 두 번째 자동화 경로. macOS/Linux/Windows 모두에서 한컴오피스 설치 없이 .hwp(HWP 5.0 binary) 처리 가능.

### 언제 이 경로를 선택?

| 시나리오 | 권장 경로 |
|----------|----------|
| Mac/Linux dev 환경, 한컴 미설치 | **B. Rust + rhwp** |
| 표/스타일 IR 직접 조작 (헤더 매칭, 컬럼 자동 탐색) | **B. Rust + rhwp** |
| HWP 5.0 binary 직접 다루기 | **B. Rust + rhwp** |
| HWPX 직접 XML 편집이 이미 동작 중 | A. Python + lxml |
| 한컴 PDF 변환 충실도 필요 | A. Python + COM |
| 폼 필드 (양식 입력 필드 채우기) | A. Python + COM (또는 향후 hwpctl Field API) |

### 위치

```
hwpx-generator/
├── hwp-automate-poc/      Rust PoC (3 시나리오 검증된 데모, binary)
└── hwp-automate-py/       Python 바인딩 (PyO3 abi3-py39 wheel)
../codebase/rhwp/          rhwp 엔진 (gh repo clone edwardkim/rhwp; 별도 codebase)
../codebase/hop/           Tauri 데스크톱 앱 (참고 패턴 출처)
```

두 서브프로젝트의 `Cargo.toml` 은 `rhwp = { path = "../../codebase/rhwp" }` 로 path 의존. 동료 머신으로 옮길 때 `codebase/rhwp` 가 같은 상대 위치에 있어야 함. 운영용은 향후 hop 처럼 git submodule(third_party/rhwp) 패턴으로 vendor 고정 가능.

### 빠른 시작

#### Rust 측 — PoC binary 직접 실행

```bash
cd hwp-automate-poc
cargo run                                       # 기본: biz_plan.hwp 의 자격증 컬럼 자동 채우기
cargo run -- <template.hwp> <output.hwp>        # 다른 양식 지정
# 결과: hwp-automate-poc/output/poc_v3.hwp
```

> **원칙**: 새 문서를 from-scratch 로 빌드하는 패턴은 본 경로에서 사용하지 않는다. 항상 사용자가 제공한 양식(.hwp) 을 베이스로 한다.

#### Python 측 — venv 에 wheel 설치 후 import

```bash
cd hwp-automate-py
python3 -m venv .venv && source .venv/bin/activate
pip install --upgrade maturin
maturin develop --release       # rhwp 컴파일 + abi3 wheel 빌드 + venv 설치 (~30초)
```

```python
import hwp_automate

# 1. 분석 — 양식의 표/스타일/번호 인벤토리 (read-only)
info = hwp_automate.analyze_template("template.hwp")
# {'tables': [...], 'style_count': 26, 'numbering_count': 8, ...}

# 2. 양식 표 자동 채우기 — 헤더 매칭으로 컬럼 자동 탐색
hwp_automate.fill_template_table(
    "template.hwp", "filled.hwp",
    {"header_match": "성명", "column": "자격증",
     "values": {1: "정보처리기사", 2: "정보보안기사", 3: "네트워크관리사", 4: "컴활 1급"}}
)
```

### 검증된 능력 (PoC v1~v3 + Python 바인딩)

| # | 능력 | 검증 방법 |
|---|------|----------|
| 1 | 양식 로드 + 표 인벤토리 | `DocumentCore::from_bytes` 로 사용자 양식 로드, 8개 표 자동 발견 |
| 2 | 헤더 매칭으로 표 식별 | "성명" 텍스트 포함된 5×6 표 자동 선택 (인덱스 하드코딩 없음) |
| 3 | 컬럼 헤더로 위치 자동 탐색 | "자격증" 헤더 → col=5 자동 산출 (양식 변경에도 견딤) |
| 4 | 빈 셀에 값 정확 삽입 | row 단위 4명 자격증 채움, 라운드트립 4/4 일치 |
| 5 | 양식의 다른 부분 무손상 보존 | 8개 표 중 7개 + 모든 텍스트·스타일·이미지 그대로 |

### 한계 (rhwp v0.7.x 기준)

- **outline 자동 번호 SVG 미렌더** — IR (`head=Outline`) 은 정확하나 rhwp v0.7.x SVG 렌더러가 자동 번호 텍스트를 안 그림. **한컴/모바일 한글에서 열면 정상 표시.**
- **HWPX 직렬화 표/그림 부분 미완** — Stage 3~5 이월. HWP 5.0 binary 저장은 안정적이므로 출력은 .hwp 권장.
- **HWPX 출처 문서 저장 비활성화** (rhwp #196) — HWPX→HWP 완전 변환기(#197) 미완. 입력도 가급적 HWP 5.0 binary.
- **Document IR 외부 직접 변형 불가** — `DocumentCore.document` 가 `pub(crate)`. 모든 변경은 `*_native` 메서드 경유.
- **`@rhwp/core` npm 미배포** — Node 직접 사용 비추. Python(PyO3) 또는 Rust 라이브러리 경유.

### 빌드 환경

| 도구 | 최소 버전 | 설치 (macOS) |
|------|----------|-------------|
| Rust | 1.75+ | `brew install rust` (검증: 1.95.0) |
| maturin | 1.13+ | `brew install maturin` 또는 venv 안에서 `pip install maturin` |
| Xcode CLT | clang 21+ | `xcode-select --install` (이미 설치됨) |
| Python | 3.9+ | abi3-py39 wheel 이라 3.9~3.14 호환 |

### 함정 (Python+lxml 경로와 다름)

이 경로는 **rhwp 가 IR 단계에서 검증** (891+ 테스트) 하므로 다음 함정들에 노출되지 않음:
- XML 선언부 형식 민감성 → rhwp 직렬화기가 처리
- ZIP 압축 방식 보존 → HWP 5.0 binary 는 ZIP 아님 (CFB)
- charPrIDRef 보존 → IR 의 char_shape_id 가 자동 매핑
- 네임스페이스 보존 → rhwp 가 OOXML 처리 시 자동
- 좀비 한글 프로세스 → COM 미사용
- PrintMethod=4 PDF 버그 → COM 미사용

대신 새로 인지할 함정은 다음 절 참고.

### Rust+rhwp 경로 함정

1. **양식 채우기만 사용 — from-scratch 빌드 금지**
   - `create_blank_document_native` / `Document::default` 같은 from-scratch 진입점은 본 경로에서 **절대 사용 금지**.
   - 이유: 사용자 도메인(공공/기업 양식 자동화) 은 기관 표준 템플릿의 구조·스타일·로고를 충실히 보존해야 하며, from-scratch 빌드는 이를 만족할 수 없음.
   - 항상 사용자가 제공한 .hwp 양식을 `DocumentCore::from_bytes` 로 로드해 빈 셀에만 값을 삽입하는 패턴을 사용.

2. **셀 인덱스는 row-major linear** — `cell_idx = row * cols + col`. 사용 전 표 헤더 매칭으로 cols 확인 필수.

3. **빈 cell 의 paragraphs 는 항상 ≥1** — HWP 표는 모든 셀에 최소 1개 문단 보장. 빈 텍스트 셀이라도 `cell.paragraphs[0]` 접근 가능.

4. **path 의존 절대경로 함정** — Cargo.toml 의 `path = "../../codebase/rhwp"` 는 상대 경로. CI/배포 시 디렉토리 레이아웃 보장 필요.

### 진입점 — Python (PyO3 노출 + CLI)

**Rust 확장 모듈 함수 3개** (`hwp_automate.*`):

| 함수 | 용도 |
|------|------|
| `analyze_template(path)` | 양식 표·스타일·번호 인벤토리 (read-only) |
| `fill_template(template, out, operations, dry_run=False, verify=True)` ★ | 여러 표·여러 컬럼·여러 셀 일괄 채우기. Pre-flight + post-fill 검증 |
| `fill_template_table(template, out, mapping, ...)` | 단일 표·단일 컬럼 편의 wrapper |

**Python 보조 모듈** (`hwp_automate_cli/`, wheel 비번들):

| 모듈 | 용도 |
|------|------|
| `hwp_automate_cli.field_map_to_operations(field_map, data, table_locator)` | 기존 `field_map.json` (entity_blocks + company_lists) → operations 어댑터 |
| `python -m hwp_automate_cli analyze \| fill \| cell` | CLI 진입점 — 양식 인벤토리, field_map+data 일괄, 빠른 셀 채우기 |

상세 사용법은 `hwp-automate-py/README.md` 참조.

> **의도적으로 노출하지 않는 것**: from-scratch 로 새 문서를 만드는 함수 (예: `create_blank_document_native` 래퍼). 본 경로의 자동화 원칙 — "사용자 양식만 채운다" — 을 강제하기 위함.

## AI 통합 진입점

경로 B 의 Python API 위에 세 가지 AI 통합 진입점이 layered. 사용자 환경에 따라 골라 쓰면 됨.

### A. Claude Code Skill (`fill-hwp`)

`.claude/skills/fill-hwp/SKILL.md` — Claude Code 사용자가 `/fill-hwp 양식.hwp` 한 명령으로 양식 채우기 시작. AI 가 자동 분석 → 사용자에게 필요 정보 질문 → 콘텐츠 생성 → fill_template 호출 → 검증 의 multi-turn playbook.

`.claude/hooks/hwp-fill-verify.py` (PostToolUse) — fill 명령 후 출력 파일 자동 확인 (크기·CFB 매직).

`.claude/settings.json` — hook 등록.

**가장 빠른 진입점.** Claude Code 안에서 즉시 동작. 별도 설치 없음.

### B. MCP 서버

`hwp-automate-py/mcp_server.py` — FastMCP stdio 서버. Claude Desktop / Claude Code / Cursor 등 모든 MCP 호환 클라이언트에서 사용 가능.

5 개 tool 노출:

| tool | 용도 |
|------|------|
| `analyze_form` | 양식 구조·빈 셀·라벨 추론 결과 (AI 가 의미 파악) |
| `preview_form_structure` | 가벼운 markdown 요약 (큰 양식 첫 검토) |
| `fill_form` | operations 로 채우기 (dry_run, verify 지원) |
| `fill_form_from_data` | 기존 field_map.json + data.json 호환 |
| `verify_output` | 결과 파일 셀 값 라운드트립 검증 |

설치 + 실행:

```bash
cd hwp-automate-py
pip install --upgrade 'mcp[cli]>=1.2.0'   # Python 3.10+
maturin develop --release                  # rhwp wheel 설치
python mcp_server.py                       # 또는 클라이언트가 자동 spawn
```

Claude Desktop 설정 (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "hwp-automate": {
      "command": "/path/to/hwp-automate-py/.venv/bin/python",
      "args": ["/path/to/hwp-automate-py/mcp_server.py"]
    }
  }
}
```

Claude Code 등록:

```bash
claude mcp add hwp-automate -- /path/to/.venv/bin/python /path/to/mcp_server.py
```

### C. (선택) Standalone CLI Agent

`hwp-automate-py/hwp_automate_agent/` — 미작성. Claude Agent SDK 기반 자체 CLI. batch / cron / CI 시나리오에서만 가치. 1·2 단계 충분 검증 후 도입.

### 분석 결과 강화 (모든 진입점 공통 전제)

`analyze_template` 가 각 표의 `cells`, `empty_cells`, `suggested_fields` 를 노출.

- `cells`: 모든 셀 (row, col, text, is_empty, neighbor_label)
- `empty_cells`: 빈 셀 + neighbor_label 추론 (왼쪽 셀 우선, 위쪽 셀 차순위)
- `suggested_fields`: 라벨 추론 성공한 빈 셀만 — `[{"label":"매출액(백만원)", "row":4, "col":2}, ...]` AI 가 즉시 사용자에게 "어떤 정보가 필요한가" 물어볼 수 있는 단서.

## HWPX 파일 수정 시 필수 주의사항

> **적용 범위**: 이 절은 **경로 A (Python + lxml)** 에 한정. 경로 B (Rust + rhwp) 는 IR 단계에서 rhwp 가 처리하므로 이 함정들에 노출되지 않음.

### 1. XML 선언부 (가장 빈번한 오류 원인)

한컴오피스는 XML 선언부 형식에 **매우 민감**하다.

```xml
<!-- 올바른 형식 (한컴오피스 호환) -->
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ...>

<!-- 잘못된 형식 (lxml 기본 출력 — 한컴오피스에서 파일이 열리지 않음) -->
<?xml version='1.0' encoding='UTF-8'?>
<hs:sec ...>
```

**차이점**:
- 큰따옴표(`"`) vs 작은따옴표(`'`) — 한컴오피스는 큰따옴표만 인식
- `standalone="yes"` 필수 — 생략 시 인식 불가
- 선언부와 루트 요소 사이에 **줄바꿈 없어야** 함 — lxml은 기본적으로 줄바꿈 추가

**해결**: `hwpx_editor.py`의 `serialize_xml()` 메서드가 원본 XML 선언부를 보존하여 직접 구성한다. **절대 `etree.tostring(xml_declaration=True)`를 사용하지 말 것**.

### 2. ZIP 압축 방식 보존

HWPX는 ZIP 아카이브이며, 엔트리별 압축 방식이 다르다:

| 엔트리 | 압축 방식 | 변경 시 |
|--------|----------|--------|
| `mimetype` | `ZIP_STORED` (비압축) | **파일 인식 불가** |
| `version.xml` | `ZIP_STORED` | 경우에 따라 오류 |
| `BinData/*.png` | `ZIP_STORED` | 이미지 깨짐 가능 |
| `Preview/*.png` | `ZIP_STORED` | 미리보기 깨짐 |
| `Contents/*.xml` 등 | `ZIP_DEFLATED` | 정상 |

**해결**: `hwpx_editor.py`가 `__init__`에서 원본 ZIP의 `compress_type`을 모두 기록하고, `save()` 시 그대로 복원한다.

### 3. PrintMethod=4 버그

`settings.xml`에 `PrintMethod=4`(2페이지/장 인쇄)가 있으면 PDF 변환 시 페이지 수가 반으로 줄어든다.

- **반드시 COM이 파일을 열기 전에** `fix_hwpx_for_pdf()`로 `PrintMethod=0`으로 변경
- COM은 파일을 열 때 settings.xml을 읽으므로, 열린 후 수정하면 효과 없음
- `bridge.py`의 `fix_hwpx_for_pdf()` 함수가 이를 처리

### 4. 네임스페이스 보존

HWPX XML은 12개 이상의 네임스페이스를 사용한다. `xml.etree.ElementTree`는 프리픽스를 `ns0:`, `ns1:`로 변환하여 한컴오피스가 파일을 인식하지 못한다. **반드시 `lxml`을 사용**해야 원본 프리픽스(`hp:`, `hs:`, `hh:` 등)가 보존된다.

### 5. charPrIDRef (글꼴 스타일 참조)

셀 텍스트를 수정할 때 `<hp:run>`의 `charPrIDRef` 속성을 변경하면 안 된다. 이 값은 `header.xml`에 정의된 글자 속성(폰트, 크기, 굵기 등)을 참조하는 ID이다. `hwpx_editor.py`의 `set_cell_text()`는 기존 `<hp:run>`의 `charPrIDRef`를 그대로 유지한다.

**참고: 문서별 charPrIDRef 값이 다르다** — 같은 "본문 텍스트"라도 문서마다 ID가 다를 수 있으므로, 새 문서 작업 시 반드시 해당 문서의 header.xml에서 확인해야 한다.

## COM 자동화 주의사항

> **적용 범위**: 이 절은 **경로 A (Python + COM)** 에 한정. 경로 B (Rust + rhwp) 는 COM 을 사용하지 않음.

### 좀비 프로세스

COM 자동화 중 오류 발생 시 `Hwp.exe`가 백그라운드에 남을 수 있다. 파일 잠금으로 이어져 `PermissionError`를 유발한다.

```bash
# 좀비 프로세스 정리
taskkill.exe /F /IM Hwp.exe
```

### 타임아웃

- PDF SaveAs는 119페이지 기준 약 200초 소요
- `timeout`은 최소 300초 이상 설정
- `Print.Execute` 호출 시 RPC 연결 끊김 — **사용 금지**

### PageSetup 한계

- COM `PageSetup`은 HWP 파일에 대해 모든 값이 0을 반환 (Width=0, Height=0, Landscape=0)

## HWPX 포맷 참고

### ZIP 구조

```
mimetype                    # "application/hwp+zip" (반드시 첫 번째, 비압축)
version.xml                 # HWPML 버전 (현재 1.5)
META-INF/                   # container.xml, container.rdf, manifest.xml
Contents/
  content.hpf               # OPF 패키지 매니페스트
  header.xml                # 폰트, 스타일 (charPr, paraPr), 테두리
  section0.xml              # 본문: 문단, 표, 이미지, 레이아웃
BinData/                    # 임베디드 이미지 (PNG, BMP)
Preview/                    # 미리보기 (PrvText.txt, PrvImage.png)
settings.xml                # 앱 설정 (인쇄, 줌, 캐럿 위치)
```

### XML 네임스페이스

| 프리픽스 | URI | 용도 |
|---------|-----|------|
| `hp:` | `http://www.hancom.co.kr/hwpml/2011/paragraph` | 문단, 런, 텍스트, 표, 셀 |
| `hs:` | `http://www.hancom.co.kr/hwpml/2011/section` | 섹션 루트, 페이지 속성 |
| `hh:` | `http://www.hancom.co.kr/hwpml/2011/head` | 폰트, 글자/문단 속성, 스타일 |
| `hc:` | `http://www.hancom.co.kr/hwpml/2011/core` | 코어 타입 |
| `ha:` | `http://www.hancom.co.kr/hwpml/2011/app` | 앱 설정 |

### 단위 & 구조

- **HWPUNIT**: 1/7200 인치. A4 = 59528 x 84188
- **스타일 참조**: `charPrIDRef`, `paraPrIDRef`, `borderFillIDRef` (숫자 ID)
- **표 구조**: `<hp:tbl>` → `<hp:tr>` → `<hp:tc>` → `<hp:subList>` → `<hp:p>` → `<hp:run>` → `<hp:t>`
- **셀 병합**: `<hp:cellSpan colSpan="N" rowSpan="N"/>`
- **이미지**: `binaryItemIDRef`로 `content.hpf`의 ID 참조

## 템플릿 시스템

### 디렉토리 구조

```
templates/
  cloud_integrated/        # 기본 템플릿 (test_01.hwpx용)
    template.json          # 메타 정보 + 찾아바꾸기 매핑
    field_map.json         # 커버 표 셀 좌표 매핑
  새_양식/                 # 새 템플릿 추가 시
    template.json
    field_map.json
```

### template.json

COM 찾아바꾸기 매핑을 정의. `find`에 원문 텍스트, `data_key`에 입력 JSON 키를 지정.

- **단순 치환**: `data_key`로 값을 가져와 `find` 텍스트를 교체
- **포맷 치환**: `format` 키가 있으면 `data_key`의 상위 객체 필드로 포맷팅

### field_map.json

XML 셀 채우기 좌표 매핑. `entity_blocks`(기관 정보)와 `company_lists`(기업 목록)로 구성. `(start_row + offset, col)` 형식으로 셀 위치를 지정.

## 테스트

```bash
# HwpxEditor 단위 테스트 (11개, WSL에서 실행)
python3 -m pytest tests/test_hwpx_editor.py -v

# COM 모듈 테스트 (Windows Python에서만 실행 가능)
# tests/test_hwp_com_module.py — WSL에서는 sys.exit(1)로 중단됨

# 전체 통합 테스트 (WSL, Windows 한컴오피스 필요)
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --data data/sample_input.json \
  --output output/milestone1 \
  --compare ref/test_01.pdf
```

**주의**: `python3 -m pytest tests/ -v`로 전체 테스트를 실행하면 `test_hwp_com_module.py`가 Windows 전용이므로 `sys.exit(1)`로 프로세스가 종료된다. **WSL에서는 반드시 테스트 파일을 개별 지정**해야 한다.

## 빠른 명령어 참고

```bash
# HWPX ZIP 구조 확인
python3 -c "import zipfile; z = zipfile.ZipFile('ref/test_01.hwpx'); print(z.namelist())"

# 좀비 한글 프로세스 정리
taskkill.exe /F /IM Hwp.exe

# 새 템플릿 구조 분석
python3 src/extract_template.py --hwpx ref/새양식.hwpx --all-tables

# 템플릿 설정 초안 자동 생성
python3 src/extract_template.py --hwpx ref/새양식.hwpx --generate-template-config -o templates/새양식/
```
