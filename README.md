# HWPX Generator

한컴오피스 한글(Hangul) 문서를 **프로그래밍·AI 로 자동 생성**하는 도구. 정부 사업계획서 같은 복잡한 양식부터 단순 양식 채우기까지 두 자동화 경로를 모두 지원하며, Claude Code · Claude Desktop · CLI · MCP 클라이언트 등 다양한 AI/도구 환경에서 사용 가능.

## 두 자동화 경로

| 경로 | 환경 | 강점 | 사용처 |
|---|---|---|---|
| **A. Python + lxml + COM** (기존) | WSL (Ubuntu) + Windows + 한컴오피스 2024 | HWPX XML 직접 편집, 한컴 PDF 충실도, 119페이지 마크다운 → HWPX 자동 변환 | 정밀 서식 재현, PDF 출력 |
| **B. Rust + rhwp** (PR #1~#3) | macOS / Linux / Windows 어디서든 한컴오피스 설치 없이 동작 | HWP 5.0 binary 직접 처리, 다중 셀 일괄 채우기, AI 친화 API | 양식 자동 채우기, AI 통합 |

두 경로는 **배타적이지 않고 한 저장소 안에 공존**합니다. 시나리오에 따라 골라 쓰거나 조합하여 사용.

## AI 통합 진입점 (경로 B)

경로 B 위에 세 가지 AI 통합 진입점이 layered. 사용자 환경에 따라 가장 편한 길로 시작:

| 진입점 | 위치 | 시작 명령 | 적합한 사용자 |
|---|---|---|---|
| **Claude Code Skill** | `.claude/skills/fill-hwp/SKILL.md` | `/fill-hwp 양식.hwp` (Claude Code 안) | Claude Code 사용자 (즉시) |
| **MCP 서버** | `hwp-automate-py/mcp_server.py` | `python mcp_server.py` (또는 클라이언트 자동) | Claude Desktop / Cursor / 기타 MCP 클라이언트 |
| **(선택) Standalone CLI Agent** | 미구현 (3 단계 옵션) | — | batch · CI · cron |

세 진입점 모두 같은 `hwp_automate` 라이브러리 위에 layered. 코드 중복 없음.

## 주요 기능

### 경로 A — Python + lxml + COM
- **마크다운 → HWPX 자동 변환** — `.md` 파일을 파싱하여 HWPX 양식의 지정 위치에 서식 있는 본문 삽입
- **Two-Pass 하이브리드 파이프라인** — Pass 1(XML 직접 편집) + Pass 2(COM 서식 삽입)
- **다중 템플릿 지원** — 임의의 HWPX 파일을 템플릿으로 등록, 설정 기반 문서 생성
- **COM 포스트 포맷 패턴** — InsertText → 선택 → 서식 적용 → 해제로 정확한 폰트/크기 렌더링
- **PDF 자동 변환** — 한컴오피스 COM API 를 통한 정확한 PDF 출력
- **PDF 비교 검증** — SSIM(구조적 유사도) + 텍스트 일치도 자동 비교
- **감사/디버깅 도구** — 교차참조 검사, 콘텐츠 무결성 감사, COM 크래시 격리, XML 직렬화 진단

### 경로 B — Rust + rhwp (cross-platform)
- **HWP 5.0 binary 직접 처리** — 한컴오피스 설치 없이 macOS/Linux/Windows 동작
- **양식 분석** — `analyze_template` 가 모든 표·셀·라벨 추론 (`cells`, `empty_cells`, `suggested_fields`, `neighbor_label`) 노출 → AI 가 즉시 양식 의미 파악
- **다중 셀 일괄 채우기** — `fill_template` 한 호출로 여러 표·여러 컬럼·여러 셀 처리, 표 식별은 헤더 매칭 또는 직접 좌표
- **Pre-flight + post-fill 검증** — 적용 전 모든 op 유효성 확인, 적용 후 라운드트립 자동 검증
- **dry_run 모드** — 실제 적용 없이 plan 만 검증
- **BinData 보존 우회** — `preserve_images=True` 기본, 원본 이미지 byte-for-byte 보존 (rhwp 라운드트립 손실 회피)
- **셀 병합 자동 처리** — `find_cell_idx` 가 (row, col) 위치 검색으로 병합된 표에서도 정확한 인덱싱
- **AI 친화 API** — JSON 호환, abi3-py39 wheel (Python 3.9~3.14), MCP 서버·Claude Code Skill 노출
- **field_map.json 어댑터** — 경로 A 의 기존 매핑 형식 그대로 재사용 가능

## 아키텍처

### 전체 구성 (두 경로 + AI 통합 layer)

```
┌─────────────────────────────────────────────────────────────────────────┐
│  사용자 환경                                                              │
│  ┌──────────────┐  ┌──────────────────┐  ┌──────────────────┐          │
│  │ Claude Code  │  │ Claude Desktop / │  │ 터미널 / CLI /    │          │
│  │ (Skill)      │  │ Cursor / 기타 MCP │  │ generate_hwpx.py │          │
│  └──────┬───────┘  └────────┬─────────┘  └─────────┬────────┘          │
│         │                   │                       │                    │
└─────────┼───────────────────┼───────────────────────┼────────────────────┘
          │                   │                       │
          ▼                   ▼                       ▼
   ┌──────────────────────────────────┐   ┌──────────────────────────┐
   │  AI 통합 layer (경로 B)           │   │  경로 A (CLI)             │
   │  • SKILL.md (multi-turn playbook)│   │  • generate_hwpx.py       │
   │  • mcp_server.py (5 tools)       │   │  • form_filler.py         │
   │  • field_map.json adapter        │   │  • md_parser → md_to_ops  │
   └──────────────┬───────────────────┘   └──────────────┬───────────┘
                  │                                       │
                  ▼                                       ▼
   ┌──────────────────────────────────┐   ┌──────────────────────────┐
   │  hwp_automate (PyO3 abi3 wheel)   │   │  bridge.py (WSL→Win)      │
   │  • analyze_template               │   │  + hwpx_editor.py (lxml)  │
   │  • fill_template                  │   │  + hwp_com.py (pywin32)   │
   │  • preserve_images_from_source    │   │  + pdf_compare.py (SSIM)  │
   └──────────────┬───────────────────┘   └──────────────┬───────────┘
                  │                                       │
                  ▼                                       ▼
   ┌──────────────────────────────────┐   ┌──────────────────────────┐
   │  rhwp (외부 의존, MIT)            │   │  한컴오피스 한글 2024      │
   │  Rust HWP 5.0 binary engine       │   │  COM API (Windows)        │
   │  ../codebase/rhwp                 │   │                           │
   └──────────────────────────────────┘   └──────────────────────────┘
```

세 AI 진입점은 같은 `hwp_automate` Python wheel 위에 layered (코드 중복 0). 그 아래 rhwp 가 HWP 5.0 binary 의 모든 파싱·직렬화·IR 변형을 책임짐. 경로 A 는 HWPX (XML) 와 COM 자동화로 별도 운영.

### Two-Pass 파이프라인 (경로 A 상세)

```
┌─────────────────────────────────────────────────────────────────┐
│  Pass 1: XML 직접 편집 (WSL, lxml)                               │
│                                                                  │
│  ┌──────────┐   ┌──────────────┐   ┌──────────────┐             │
│  │  JSON    │──>│ field_mapper │──>│ hwpx_editor  │             │
│  │  Input   │   │  (셀 매핑)   │   │ (184셀 채움)  │             │
│  └──────────┘   └──────────────┘   └──────┬───────┘             │
│                                           │                      │
│  ┌──────────┐   ┌────────────────┐        ▼                      │
│  │ Markdown │──>│  md_parser     │   ┌──────────────┐            │
│  │ Sections │   │ (구조화 파싱)   │   │ form_filler  │            │
│  └──────────┘   └───────┬────────┘   │ (오케스트레이터)│            │
│                         │            └──────┬───────┘            │
│                         ▼                   │                     │
│                 ┌────────────────┐           │                     │
│                 │ section_mapper │           │                     │
│                 │ (마커 매핑)     │           │                     │
│                 └───────┬────────┘           │                     │
│                         │                   │                     │
│                         ▼                   │                     │
│                 ┌────────────────┐           │                     │
│                 │  md_to_ops     │           │                     │
│                 │ (COM 명령 생성) │           │                     │
│                 └───────┬────────┘           │                     │
│                         │                   │                     │
├─────────────────────────┼───────────────────┼─────────────────────┤
│  Pass 2: COM 자동화 (Windows, pywin32)      │                     │
│                         │                   │                     │
│                         ▼                   ▼                     │
│                 ┌────────────────────────────────┐                │
│                 │         bridge.py               │                │
│                 │  (WSL→Windows 브릿지)            │                │
│                 │  포스트 포맷: Insert→Select→     │                │
│                 │  Format→Deselect                │                │
│                 └───────────┬────────────────────┘                │
│                             │                                     │
│                        ┌────┴────┐                                │
│                        ▼         ▼                                │
│                     .hwpx      .pdf                               │
├───────────────────────────────────────────────────────────────────┤
│  검증 (WSL)                                                       │
│                 ┌─────────────────┐                                │
│                 │  pdf_compare.py │                                │
│                 │  (SSIM 검증)    │                                │
│                 └─────────────────┘                                │
└───────────────────────────────────────────────────────────────────┘
```

### 모듈 구성

| 모듈 | 실행 환경 | 역할 |
|------|----------|------|
| `src/form_filler.py` | WSL | **파이프라인 오케스트레이터**. Pass 1 + Pass 2 순차 실행 |
| `src/md_parser.py` | WSL | 마크다운 파서. `.md` → 구조화된 블록(헤딩/문단/표/리스트) |
| `src/md_to_ops.py` | WSL | 마크다운 블록 → COM 자동화 명령 시퀀스 변환 |
| `src/section_mapper.py` | WSL | 마크다운 섹션 → HWPX 마커(##SEC_CONTENT##) 매핑 |
| `src/generate_hwpx.py` | WSL | 메인 CLI 파이프라인. 전체 흐름 제어 |
| `src/bridge.py` | WSL | WSL↔Windows Python 브릿지. 포스트 포맷 패턴 구현 |
| `src/hwp_com.py` | Windows | 한컴오피스 COM 자동화 (pywin32) |
| `src/hwpx_editor.py` | WSL | HWPX ZIP 내부 section0.xml 직접 수정 (lxml) |
| `src/field_mapper.py` | WSL | JSON 입력 데이터 → 셀 좌표 매핑 |
| `src/pdf_compare.py` | WSL | PDF 페이지별 SSIM + 텍스트 비교 |
| `src/extract_template.py` | WSL | HWPX 파일 구조 분석/추출 |

### 데이터 흐름

```
sample_input.json
       │
       ▼
 field_mapper.py ──(field_map.json)──> {(row, col): text} 딕셔너리
       │
       ▼
 hwpx_editor.py ──> section0.xml 내 빈 셀에 텍스트 삽입 (93개 셀)
       │
       ▼
   bridge.py ──> Windows Python으로 COM 스크립트 전달
       │
       ▼
  hwp_com.py ──> 한컴오피스 COM: 텍스트 교체 + HWPX/PDF 저장
       │
       ▼
 pdf_compare.py ──> 참조 PDF와 SSIM/텍스트 비교 검증
```

## 환경 요구사항

### 경로 A — Python + lxml + COM (Windows 필수)

한컴오피스 COM API 만이 보장하는 가치:
- 119페이지, 456개 표, 63개 이미지, 370개 글자 속성의 **100% 서식 재현**
- 내장 렌더링 엔진을 통한 **정확한 PDF 변환** (`SaveAs "PDF"`)
- 기존 문서를 열어서 수정하는 **템플릿 기반 워크플로우**

**필수 구성:**

- **Windows 측**: Windows 10/11, 한컴오피스 2024 (한글), Python 3.13+, `pywin32`
- **WSL 측**: Ubuntu (WSL2 권장), Python 3.12+, `lxml`, `PyMuPDF(fitz)`, `scikit-image`, `Pillow`, `numpy`

**WSL ↔ Windows 브릿지:** WSL Python 이 `bridge.py` 로 Windows Python(`python.exe`) 을 subprocess 호출. Windows Python 이 한컴오피스 COM API 로 문서를 열고·수정·PDF 저장. 결과는 `/mnt/d/` 등 공유 드라이브로 WSL 에서 접근.

### 경로 B — Rust + rhwp (크로스플랫폼, 한컴 불필요)

**필수 구성 (모든 OS 동일):**
- Rust 1.75+ (`brew install rust` 또는 rustup, 검증 시점 1.95.0)
- Python 3.9+ (abi3 wheel — 단일 wheel 이 3.9~3.14 모두 호환)
- maturin 1.13+ (`brew install maturin` 또는 `pip install maturin`)
- (Mac) Xcode CLT — clang 21+ 이미 시스템에 있어야 함
- `../codebase/rhwp` 위치에 [edwardkim/rhwp](https://github.com/edwardkim/rhwp) git clone (저장소 외부 의존)

**MCP 서버 사용 시 추가 (선택):**
- Python 3.10+ (mcp SDK 가 3.10+ 요구 — abi3 wheel 자체는 3.9 호환 유지)
- `mcp[cli]>=1.2.0` (`pip install 'hwp-automate[mcp]'` 또는 직접 설치)

**Claude Code Skill 사용 시 추가 (선택):**
- Claude Code 가 이미 설치되어 있어야 함 (이 README 가 작성된 환경)
- 별도 설치 명령 없음 — `.claude/skills/fill-hwp/SKILL.md` 가 자동 인식

### 외부 의존 디렉토리 레이아웃

경로 B 는 `../codebase/rhwp` 위치에 rhwp 가 있어야 동작 (본 저장소에 미포함):

```
temp_git/                  (또는 임의 작업 루트)
├── hwpx-generator/         ← 이 저장소
└── codebase/               ← 별도 git clone (저장소 외부)
    └── rhwp/               ← gh repo clone edwardkim/rhwp
```

운영 시 vendor 고정이 필요하면 hop 의 `third_party/rhwp` git submodule 패턴을 따라할 수 있음 (현재 PoC 단계에선 path 의존이 더 빠른 반복을 위해 유지).

## 설치 및 실행

### 경로 A 설치 (Python + lxml + COM)

#### 1. 전제조건 확인

```bash
# WSL에서 확인
python3 --version        # 3.12+
# Windows Python 경로 확인
/mnt/c/Users/<username>/AppData/Local/Microsoft/WindowsApps/python.exe --version
```

#### 2. WSL Python 패키지 설치

```bash
pip3 install --break-system-packages lxml pymupdf Pillow scikit-image numpy
```

#### 3. Windows Python 패키지 설치

```powershell
# Windows PowerShell에서
pip install pywin32
```

#### 4. 실행 (`generate_hwpx.py`)

```bash
# 데이터를 적용하여 문서 생성 + PDF 변환 (기본 cloud_integrated 템플릿)
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --data data/sample_input.json \
  --output output/

# 다른 템플릿 설정으로 문서 생성
python3 src/generate_hwpx.py \
  --template ref/새양식.hwpx \
  --template-dir templates/새양식/ \
  --data data/새양식_input.json \
  --output output/

# PDF 비교 검증 포함
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --output output/ \
  --pdf-only \
  --compare ref/test_01.pdf
```

CLI 옵션:

| 옵션 | 설명 |
|------|------|
| `--template, -t` | (필수) 템플릿 HWPX 파일 경로 |
| `--template-dir` | 템플릿 설정 디렉토리 (기본: `templates/cloud_integrated/`) |
| `--data, -d` | JSON 입력 데이터 파일 경로 |
| `--output, -o` | 출력 디렉토리 (기본: `output`) |
| `--pdf-only` | 데이터 없이 템플릿을 그대로 PDF로 변환 |
| `--no-pdf` | PDF 생성 건너뛰기 |
| `--compare, -c` | 비교할 참조 PDF 경로 |

### 경로 B 설치 (Rust + rhwp, cross-platform)

#### 1. 외부 의존 (rhwp) 준비

```bash
mkdir -p ../codebase
cd ../codebase
gh repo clone edwardkim/rhwp        # 또는 git clone https://github.com/edwardkim/rhwp.git
cd -
```

#### 2. Rust + maturin 설치 (이미 있으면 생략)

```bash
# macOS
brew install rust maturin

# Linux / Windows — rustup.rs 또는 패키지 매니저
```

#### 3. Python venv + wheel 빌드

```bash
cd hwp-automate-py
python3 -m venv .venv
source .venv/bin/activate          # Linux/Mac
# Windows: .venv\Scripts\activate

pip install --upgrade maturin pytest
maturin develop --release          # ~30초, rhwp 컴파일 + abi3 wheel 빌드 + venv 설치
```

#### 4. (선택) MCP 서버 의존 추가

```bash
# Python 3.10+ 환경에서만
pip install 'mcp[cli]>=1.2.0'
```

#### 5. 실행 — 3 가지 진입점

##### 5.1 — Python 직접 호출 (라이브러리)

```python
import hwp_automate

# 양식 분석 — AI 가 양식 의미를 파악할 수 있는 모든 정보 노출
info = hwp_automate.analyze_template("/path/to/form.hwp")
for t in info["tables"]:
    print(t["header"], t["empty_cells"], t["suggested_fields"])

# 양식 채우기 — 다중 표·다중 셀, dry_run+verify+preserve_images 자동
result = hwp_automate.fill_template(
    template_path="/path/to/form.hwp",
    out_path="/path/to/output.hwp",
    operations=[{
        "header_match": "성명",
        "cells": [{"row": 1, "col": 5, "value": "정보처리기사"}],
    }],
)
print(result["status"])  # "applied + verified"
```

##### 5.2 — CLI (`hwp_automate_cli`)

```bash
source hwp-automate-py/.venv/bin/activate

# 분석
python -m hwp_automate_cli analyze --template /path/to/form.hwp --json

# 빠른 셀 채우기
python -m hwp_automate_cli cell \
  --template /path/to/form.hwp \
  --output /path/to/output.hwp \
  --header-match 성명 \
  --cell 1,5,정보처리기사 \
  --cell 2,5,정보보안기사

# field_map.json + data.json 으로 일괄 채우기
python -m hwp_automate_cli fill \
  --template /path/to/form.hwp \
  --field-map templates/cloud_integrated/field_map.json \
  --data data/sample_input.json \
  --output /path/to/output.hwp \
  --header-match "기관명"
```

##### 5.3 — Claude Code Skill (사용자가 이 환경에서 즉시 사용)

별도 설치 없음. `.claude/skills/fill-hwp/SKILL.md` 가 자동 인식되어:

```
사용자: /fill-hwp /Users/lanco/Downloads/사업신청서.hwp
Claude:  [analyze 실행] 양식 분석 완료. 다음 정보가 필요합니다:
         업종명, 주생산품, 매출액(백만원), 영업이익(백만원), 수출액(백만원), 부채비율(%)
         알려주시면 채워서 저장하겠습니다.

사용자:  철강 특수강 제조 / 스테인리스 강재 / 12,500 / 1,200 / 8,300 / 45.2

Claude:  [fill_template 실행]
         ✅ /Users/lanco/Downloads/사업신청서_filled.hwp (35MB) 완료
         applied + verified, 6 셀 적용. 한컴/모바일 한글에서 확인해 주세요.
```

PostToolUse hook (`.claude/hooks/hwp-fill-verify.py`) 이 fill 명령 후 출력 파일 자동 확인 (크기, HWP 5.0 CFB 매직).

##### 5.4 — MCP 서버 (Claude Desktop / Cursor / 기타 MCP 클라이언트)

```bash
cd hwp-automate-py
source .venv/bin/activate
pip install 'mcp[cli]>=1.2.0'      # Python 3.10+
python mcp_server.py                # 또는 클라이언트가 자동 spawn
```

**Claude Desktop 등록** (`~/Library/Application Support/Claude/claude_desktop_config.json`):

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

**Claude Code 등록**:

```bash
claude mcp add hwp-automate -- /abs/path/to/.venv/bin/python /abs/path/to/mcp_server.py
```

노출 5 tools:

| tool | 용도 |
|---|---|
| `analyze_form` | 양식 구조·빈 셀·라벨 추론 (AI 가 의미 파악) |
| `preview_form_structure` | 가벼운 markdown 요약 (큰 양식 첫 검토) |
| `fill_form` | operations 로 양식 채우기 (dry_run, verify) |
| `fill_form_from_data` | field_map.json + data.json 호환 입력 |
| `verify_output` | 결과 셀 라운드트립 검증 |

상세 문서: [`hwp-automate-py/README.md`](hwp-automate-py/README.md), [`hwp-automate-poc/README.md`](hwp-automate-poc/README.md)

## 다중 템플릿 지원 (경로 A)

> **적용 범위**: 이 절의 `templates/<name>/` 디렉토리 + `template.json` + `field_map.json` 패턴은 **경로 A** 의 자산이다. 경로 B 는 어댑터를 통해 같은 `field_map.json` 을 재사용할 수 있다 ([hwp-automate-py/README.md](hwp-automate-py/README.md) 의 `fill_template_from_data` 절 참고).

기본 제공되는 `cloud_integrated`(클라우드 종합솔루션 사업계획서) 외에 다른 한글 문서를 템플릿으로 등록하여 사용할 수 있습니다.

### 템플릿 디렉토리 구조

```
templates/
  cloud_integrated/            # 기본 템플릿 (test_01.hwpx용)
    template.json              # 메타 정보 + 찾아바꾸기 매핑
    field_map.json             # 커버 테이블 셀 좌표 매핑
  새_양식_이름/                # 새 템플릿 추가 시
    template.json
    field_map.json
```

### template.json 형식

각 템플릿 설정 디렉토리에는 `template.json`이 있어 찾아바꾸기 매핑과 메타 정보를 정의합니다.

```jsonc
{
  "name": "cloud_integrated",
  "description": "2026 클라우드 종합솔루션 지원사업 사업계획서",
  "cover_table_index": 0,       // XML 셀 채우기 대상 표 인덱스
  "replacements": [
    {
      "find": "원본 문서의 텍스트",    // COM 찾아바꾸기 대상
      "data_key": "사업명"            // 입력 JSON의 키
    },
    {
      "find": "원본의 기간 텍스트",
      "data_key": "수행기간._개발",   // 상위 객체 참조
      "format": " · 개발 : {개발시작} ~ {개발종료} ({개발기간})"  // 포맷 문자열
    }
  ]
}
```

- **단순 치환**: `data_key`로 입력 데이터에서 값을 가져와 `find` 텍스트를 교체
- **포맷 치환**: `format` 키가 있으면 `data_key`의 상위 객체를 사용하여 `{필드명}` 패턴을 채움

### 새 템플릿 등록 워크플로우

```bash
# Step 1: HWPX 파일의 표 구조 분석
python3 src/extract_template.py --hwpx ref/새양식.hwpx --all-tables

# Step 2: 템플릿 설정 초안 자동 생성
python3 src/extract_template.py \
  --hwpx ref/새양식.hwpx \
  --generate-template-config \
  -o templates/새양식/

# Step 3: 생성된 template.json, field_map.json 수동 검토 및 보정
#   - cover_table_index 확인
#   - replacements에 찾아바꾸기 항목 추가
#   - field_map.json의 셀 좌표 검증

# Step 4: 입력 데이터 JSON 작성 후 문서 생성
python3 src/generate_hwpx.py \
  --template ref/새양식.hwpx \
  --template-dir templates/새양식/ \
  --data data/새양식_input.json \
  --output output/
```

### extract_template.py 옵션

| 옵션 | 설명 |
|------|------|
| `--hwpx` | HWPX 파일 경로 (기본: `ref/test_01.hwpx`) |
| `--cover` | 커버 페이지 추출 |
| `--sections` | 본문 섹션 추출 |
| `--styles` | 스타일 정보 추출 |
| `--tables` | 주요 표 추출 |
| `--all-tables` | 모든 표 요약 목록 출력 |
| `--sample-data` | sample_input.json 생성 |
| `--generate-template-config` | template.json + field_map.json 초안 자동 생성 |
| `--output, -o` | 출력 파일/디렉토리 경로 |

## 구현 접근방식 비교

3가지 접근방식을 실제로 테스트하고 평가한 결과입니다. 대상 문서는 119페이지, 456개 표의 정부 사업계획서입니다.

### 평가 결과

| 평가 기준 | COM 자동화 | python-hwpx | 직접 XML |
|-----------|:---------:|:----------:|:-------:|
| 서식 재현도 | ★★★★★ | ★★☆☆☆ | ★★★★☆ |
| PDF 변환 | ★★★★★ | ☆☆☆☆☆ | ☆☆☆☆☆ |
| 개발 용이성 | ★★★★☆ | ★★★☆☆ | ★☆☆☆☆ |
| 실행 속도 | ★★☆☆☆ | ★★★★★ | ★★★★☆ |
| 이식성 | ★☆☆☆☆ | ★★★★★ | ★★★★★ |
| 유지보수성 | ★★★★☆ | ★★★☆☆ | ★☆☆☆☆ |
| 기존 문서 수정 | ★★★★★ | ★★★☆☆ | ★☆☆☆☆ |
| **총점** | **28/35** | **18/35** | **16/35** |

### 최종 결정: 하이브리드 COM 중심 방식

**XML 직접 수정 + COM 교체/PDF**를 조합한 하이브리드 방식을 채택했습니다.

- **XML 직접 수정** (`hwpx_editor.py`): 빈 셀에 데이터를 채우는 단순 작업. lxml로 네임스페이스를 보존하면서 section0.xml을 수정합니다. 서식(charPrIDRef, paraPrIDRef)은 템플릿의 것을 그대로 유지합니다.
- **COM 텍스트 교체** (`hwp_com.py`): 사업명, 과제명 등 기존 텍스트를 찾아바꾸는 작업. 한컴오피스의 AllReplace를 사용하여 정확한 서식을 유지합니다.
- **COM PDF 변환**: `SaveAs("PDF")`로 한컴오피스 렌더링 엔진이 직접 PDF를 생성합니다.

상세 비교 보고서: [`analysis/approach_comparison.md`](analysis/approach_comparison.md)

## 핵심 기술 이슈 및 해결

### XML 선언부 호환성

**증상**: 생성된 HWPX 파일이 한컴오피스에서 열리지 않음

**원인**: lxml의 기본 XML 직렬화가 한컴오피스와 호환되지 않는 형식을 생성:
- 작은따옴표(`'`) 사용 (한컴오피스는 큰따옴표 `"` 필요)
- `standalone="yes"` 누락
- XML 선언부와 루트 요소 사이에 줄바꿈 삽입

**해결**: `hwpx_editor.py`의 `serialize_xml()` 메서드가 원본 XML 선언부를 보존하여 직접 구성. `etree.tostring(xml_declaration=True)` 대신 원본 선언부를 문자열로 접합

### ZIP 압축 방식 보존

**증상**: HWPX 파일이 열리지 않거나 이미지가 깨짐

**원인**: HWPX ZIP 엔트리별로 압축 방식이 다름 (`mimetype`, `version.xml`, 이미지 파일은 `ZIP_STORED`, 나머지는 `ZIP_DEFLATED`). 모든 엔트리를 동일하게 압축하면 한컴오피스가 인식하지 못함

**해결**: `hwpx_editor.py`가 원본 ZIP의 각 엔트리별 `compress_type`을 기록하고, 저장 시 그대로 복원

### PrintMethod=4 버그

**증상**: 119페이지 문서가 PDF로 변환하면 60페이지(가로 방향)로 출력됨

**원인**: HWPX의 `settings.xml`에 `PrintMethod=4`(2페이지/장 인쇄) 설정이 있어, COM이 PDF 생성 시 이 설정을 적용

**해결**: `fix_hwpx_for_pdf()` 함수로 COM이 파일을 열기 **전에** `PrintMethod=0`으로 변경. COM은 파일을 열 때 settings.xml을 읽으므로, 열린 후에 수정해도 효과 없음

### COM 좀비 프로세스 관리

COM 자동화 중 오류가 발생하면 `Hwp.exe`가 백그라운드에 남을 수 있습니다. 항상 `try/finally`로 `hwp.quit()`을 호출하고, 필요 시 수동 정리:

```bash
taskkill.exe /F /IM Hwp.exe
```

### 네임스페이스 보존

HWPX의 XML은 12개 이상의 네임스페이스를 사용합니다. `xml.etree.ElementTree`는 네임스페이스 프리픽스를 `ns0:`, `ns1:`로 변환하여 한글이 파일을 인식하지 못합니다. **lxml** 사용이 필수입니다 — 원본 프리픽스(`hp:`, `hs:`, `hh:` 등)를 그대로 보존합니다.

### (경로 B) 셀 병합 표의 cell_idx 어긋남

**증상**: rhwp 의 `insert_text_in_cell_native(... cell_idx ...)` 가 "셀 인덱스 N 범위 초과" 로 실패. 작은 표(병합 없음)에서는 동작하다가 큰 양식의 7×8 같은 병합 표에서 실패.

**원인**: HWP 표는 셀 병합으로 인해 `Table.cells` Vec 의 길이가 `rows × cols` 보다 작음. `cell_idx = row × cols + col` 공식이 어긋남.

**해결**: `find_cell_idx()` 가 `Table.cells` 에서 (row, col) 위치를 직접 검색. pre-flight 단계에서 한 번만 산출하여 캐시.

### (경로 B) HWP 셀 텍스트 trailing whitespace

**증상**: 사용자 입력 "단원구" 가 라운드트립 후 "단원구 " 로 보여 `verify=True` 가 실패.

**원인**: 한컴 셀 content 끝에 \n/공백을 자동 추가하는 관례.

**해결**: post-fill verify 비교를 `trim_end()` 로. leading whitespace 는 의도적일 수 있어 한쪽만 trim.

### (경로 B) rhwp BinData 라운드트립 손실

**증상**: 35MB YCP 양식을 fill 후 한컴이 "손상" 으로 판정. 54MB 코리녹스는 그림 일부 누락.

**원인**: rhwp v0.7.x 가 BinData stream (BMP 이미지 등) 을 재직렬화하면서 미세하게 변형 — 같은 stream 수지만 byte 가 ~5MB 손실.

**해결**: `merge_cfb_preserving_input()` — rhwp 출력 베이스 + BinData/Preview 만 입력 양식의 raw bytes 로 byte-for-byte 보존. 외부 `cfb` crate 의 strict 검증이 rhwp `mini_cfb` 출력과 비호환이라, rhwp 자체의 `LenientCfbReader` + `mini_cfb::build_cfb` 로 stream 단위 머지. **`preserve_images=True` 가 기본** — 사용자가 끄지 않는 한 자동 적용.

### (경로 B) HWP 표준 layout 으로 leaf 이름 → storage path 재구성

**증상**: `LenientCfbReader.list_entries()` 가 `BIN0001.bmp` 같이 leaf 이름만 반환 (parent storage 정보 없음). `mini_cfb::build_cfb` 에 그대로 넘기면 모든 stream 이 root 에 평탄화되어 한컴이 인식 못 함.

**해결**: `leaf_to_hwp_path()` 가 leaf 이름 패턴으로 storage path 추론 — `BIN****.*` → `/BinData/`, `Section{N}` → `/BodyText/`, `PrvText/PrvImage` → root 또는 `/Preview/` (양식별).

## 프로젝트 구조

```
hwpx-generator/
├── README.md
├── CHANGELOG.md                        # 변경 이력 (Milestone 3, 경로 B, V1, AI 통합 등)
├── CLAUDE.md                           # Claude Code · AI 에이전트 가이드 (두 경로 + AI 통합 진입점)
├── .gitignore                          # .venv, output, .claude 의 로컬 파일만 무시 (skill/hook 은 트래킹)
│
├── 🔵 경로 A — Python + lxml + COM (기존)
│
├── analysis/                           # 접근방식 평가 보고서
│   ├── approach_comparison.md          #   3가지 방식 종합 비교
│   ├── com_evaluation.md               #   COM API 테스트 결과
│   ├── pyhwpx_evaluation.md            #   python-hwpx 평가
│   ├── direct_xml_evaluation.md        #   직접 XML 평가
│   └── hwpx_structure_analysis.md      #   HWPX 구조 분석
├── data/
│   ├── sample_input.json               # 샘플 입력 데이터
│   ├── schema.json                     # 입력 데이터 JSON Schema
│   └── form_content_map.json           # 마크다운 섹션 → HWPX 마커 매핑
├── ref/                                # 참조 파일 (gitignore 대상)
├── src/                                # 경로 A 메인 코드
│   ├── form_filler.py                  # ★ 파이프라인 오케스트레이터 (Pass 1 + Pass 2)
│   ├── md_parser.py                    # ★ 마크다운 파서
│   ├── md_to_ops.py                    # ★ 마크다운 → COM 명령 변환
│   ├── section_mapper.py               # ★ 섹션 → 마커 매핑
│   ├── generate_hwpx.py                # 메인 CLI 파이프라인
│   ├── bridge.py                       # WSL↔Windows 브릿지 (포스트 포맷 패턴)
│   ├── hwp_com.py                      # 한컴오피스 COM 자동화 (Windows only)
│   ├── hwpx_editor.py                  # HWPX XML 편집기 (lxml)
│   ├── field_mapper.py                 # JSON→셀 좌표 매핑
│   ├── pdf_compare.py                  # PDF 비교 검증
│   └── extract_template.py             # 템플릿 구조 분석
├── templates/                          # 경로 A 양식별 매핑
│   ├── cloud_integrated/               # 클라우드 종합솔루션 템플릿
│   │   ├── template.json
│   │   └── field_map.json
│   └── gyeongnam_rbd/                  # ★ 경남 R&BD 사업계획서 (36표·184셀)
│       └── field_map.json
├── tools/
│   └── make_rawcopy.py                 # HWPX 클린 카피 유틸리티
├── tests/
│   ├── test_hwpx_editor.py             # HwpxEditor 단위 테스트
│   ├── test_hwp_com_module.py          # COM 모듈 (Windows only)
│   └── test_integration.py             # 통합 테스트
├── audit_crossrefs.py                  # 교차참조 유효성 검사
├── audit_hwpx_content.py               # 콘텐츠 무결성 감사
├── audit_section0.py                   # section0.xml 상세 감사
├── compare_section0{,_v2}.py           # section0 비교
├── debug_crash_isolate.py              # COM 크래시 격리
├── diagnose_xml_serialization.py       # XML 직렬화 진단
│
├── 🟢 경로 B — Rust + rhwp + AI 통합
│
├── hwp-automate-poc/                   # Rust PoC (binary)
│   ├── README.md
│   ├── Cargo.toml                      # rhwp = ../../codebase/rhwp
│   ├── src/main.rs                     # 양식 표 자동 채우기 데모
│   ├── output/                         # 생성된 .hwp / .svg (gitignore)
│   └── target/                         # cargo build (gitignore)
│
├── hwp-automate-py/                    # Python 바인딩 + AI 통합 진입점
│   ├── README.md
│   ├── Cargo.toml                      # PyO3 + rhwp path 의존
│   ├── pyproject.toml                  # maturin abi3-py39, mcp optional
│   ├── src/lib.rs                      # ★ analyze_template, fill_template,
│   │                                   #    fill_template_table, preserve_images_from_source
│   ├── mcp_server.py                   # ★ FastMCP stdio 서버 (5 tools)
│   ├── hwp_automate_cli/               # Python 보조 도구 (wheel 비번들)
│   │   ├── __init__.py
│   │   ├── __main__.py                 # CLI: analyze / fill / cell
│   │   └── field_map.py                # field_map.json 어댑터
│   ├── tests/
│   │   └── test_svg_regression.py      # SVG 기반 시각 회귀 (한컴 없이 자동)
│   ├── .venv/                          # 격리 venv (gitignore)
│   ├── target/                         # cargo build (gitignore)
│   └── output/                         # V1_*.hwp 등 (gitignore)
│
├── .claude/                            # Claude Code 통합 (skill/hook/settings 트래킹)
│   ├── skills/
│   │   └── fill-hwp/SKILL.md           # ★ /fill-hwp 양식.hwp playbook (multi-turn)
│   ├── hooks/
│   │   └── hwp-fill-verify.py          # PostToolUse: fill 후 출력 파일 자동 확인
│   └── settings.json                   # hook 등록
│
├── .github/
│   └── workflows/build-wheels.yml      # macOS+Linux+Windows wheel 매트릭스 빌드
│
└── output/                             # 경로 A 결과물 (gitignore)
```

**외부 의존 (본 저장소 외부, git clone 별도):**

```
../codebase/rhwp/    ← gh repo clone edwardkim/rhwp  (경로 B 핵심 엔진, MIT)
../codebase/hop/     ← gh repo clone golbin/hop  (참고 패턴 출처, MIT — 선택)
```

## 데이터 입력 형식

입력 데이터는 JSON 형식이며, `data/schema.json`에 정의된 스키마를 따릅니다.

### 주요 필드

```jsonc
{
  "사업명": "OO 종합솔루션 지원사업(통합형 OO화)",
  "과제명": "예시 과제명 — 중소제조기업형 클라우드 통합관리 SaaS 플랫폼",
  "사업개요": "과제내용에 대하여 간단히 요약 기술 (예시 텍스트)",
  "개발솔루션": ["예시 솔루션 A", "예시 솔루션 B", ...], // 최대 5개
  "수행기간": {
    "개발시작": "'26.6.30",
    "개발종료": "'27.6.30",
    "실증시작": "'27.6.30",
    "실증종료": "'27.12.31"
  },
  "대표공급기업": {                                  // 기관 정보 블록
    "기업명": "(주)예시소프트",
    "사업자등록번호": "000-00-00000",
    "대표자명": "홍길동",
    "담당자": { "성명": "김OO", "부서": "...", ... }
  },
  "클라우드사업자": { ... },                         // 동일 구조
  "협력기관": { ... },                               // 동일 구조
  "참여공급기업": [{ "기업명": "(주)예시기업A", ... }], // 최대 3개
  "도입실증기업": [{ "기업명": "(주)실증기업A", ... }]  // 최대 5개
}
```

필수 필드: `사업명`, `과제명`, `개발솔루션`, `수행기간`, `대표공급기업`, `클라우드사업자`

전체 스키마: [`data/schema.json`](data/schema.json) / 샘플 데이터: [`data/sample_input.json`](data/sample_input.json)

## 테스트

### 경로 A — 단위 + 통합 테스트

```bash
# HwpxEditor 단위 테스트 (11개, WSL 에서 실행)
python3 -m pytest tests/test_hwpx_editor.py -v

# 주의: 전체 테스트 (pytest tests/) 는 test_hwp_com_module.py 가
# Windows 전용이라 WSL 에서 sys.exit(1) 로 중단됨.
# 반드시 테스트 파일을 개별 지정할 것.
```

테스트 항목: 표 조회·범위 외 인덱스, 셀 텍스트 설정, 일괄 채우기, 저장 후 재로드, 네임스페이스 보존, charPrIDRef 보존, mimetype 비압축, XML 선언부 보존, ZIP 압축 방식 보존.

**통합 테스트 (전체 파이프라인):**

```bash
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --data data/sample_input.json \
  --output output/milestone1 \
  --compare ref/test_01.pdf
```

검증 기준: SSIM ≥ 0.90 (달성 0.9660), 페이지 수 119/119, 텍스트 일치도 0.9959.

### 경로 B — SVG 시각 회귀 + 실 양식 검증

#### SVG diff 자동 회귀 (한컴 없이 Mac/Linux 자동)

```bash
cd hwp-automate-py
source .venv/bin/activate
python -m pytest tests/ -v
```

3 testcase (`tests/test_svg_regression.py`):
- `test_no_chars_removed` — fill 후 양식의 어떤 글자도 사라지지 않는지
- `test_added_chars_match_intended` — 추가된 글자 멀티셋이 의도한 fill 값과 일치
- `test_total_char_count_difference` — 전체 글자 수 차이가 의도한 비공백 글자 수와 일치

baseline (원본) vs filled (자격증 컬럼 채움) SVG 의 글자 멀티셋 비교. 좌표는 무시 (rhwp 의 LineSeg 미세 재계산 노이즈 회피), 텍스트 보존만 검증.

#### 실 양식 V1 검증 (PR #2)

| 양식 | 크기 | 결과 |
|---|---|---|
| YCP_V0.4 (제조AI 사업신청서) | 35MB → 35MB | ✅ applied + verified, 6 셀 채움 |
| 코리녹스_V0.9 (제조AI 사업신청서) | 54MB → 54MB | ✅ applied + verified, 5 셀 채움 |

`preserve_images=True` 기본 적용으로 **BinData 54/54 동일 크기 보존** — 한컴이 손상으로 인식하지 않음.

#### CI 매트릭스 빌드 (`.github/workflows/build-wheels.yml`)

매 PR / push 시 macOS / Linux / Windows runner 에서 wheel 자동 빌드. Mac+Linux 에서는 SVG 회귀까지 자동 실행. abi3-py39 라 OS 한 곳 빌드한 wheel 이 그 OS 의 모든 Python 3.9~3.14 호환.

#### MCP 서버 smoke test (in-process)

```bash
cd hwp-automate-py
source .venv/bin/activate
python -c "
import asyncio
from mcp_server import mcp

async def t():
    tools = await mcp.list_tools()
    print(f'  {len(tools)} tools 등록')
    for tool in tools:
        print(f'    - {tool.name}')
asyncio.run(t())
"
# → 5 tools (analyze_form, preview_form_structure, fill_form, fill_form_from_data, verify_output)
```

## HWPX 형식 참고

- HWPX는 `application/hwp+zip` MIME 타입의 ZIP 아카이브
- 단위: HWPUNIT (1/7200 인치). A4 = 59528 x 84188
- 주요 네임스페이스: `hp:`(문단), `hs:`(섹션), `hh:`(헤더), `hc:`(코어), `ha:`(앱)
- 상세 구조: [`CLAUDE.md`](CLAUDE.md) 및 [`analysis/hwpx_structure_analysis.md`](analysis/hwpx_structure_analysis.md) 참고

## Acknowledgement (감사·출처)

본 저장소는 두 자동화 경로를 보유합니다. 각 경로가 의존하는 외부 오픈소스 프로젝트들에 대한 감사 표기입니다.

### 경로 A — Python + lxml + COM (기존)

| 라이브러리 | 용도 | 라이선스 |
|---|---|---|
| [lxml](https://lxml.de/) | HWPX (XML 기반) 직접 편집 | BSD |
| [pywin32](https://github.com/mhammond/pywin32) | Windows 한컴오피스 COM 자동화 | PSF |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | PDF 비교 검증 | AGPL |
| [scikit-image](https://scikit-image.org/) | SSIM 픽셀 비교 | BSD |
| 한컴오피스 한글 2024 (Windows) | COM 자동화 대상 | 한글과컴퓨터 (별도 라이선스 필요) |

### 경로 B — Rust + rhwp (크로스플랫폼, COM 불필요)

본 저장소의 `hwp-automate-poc/` 와 `hwp-automate-py/` 서브프로젝트는 다음 두 외부 오픈소스 프로젝트 위에서 만들어졌습니다.

#### 🦀 [edwardkim/rhwp](https://github.com/edwardkim/rhwp) — 핵심 엔진 (직접 의존)

- **저자:** Edward Kim ([@edwardkim](https://github.com/edwardkim))
- **라이선스:** MIT
- **설명:** Rust + WebAssembly 기반 오픈소스 HWP/HWPX 뷰어/에디터. v0.7.x 시점 891+ 테스트, hyper-waterfall 방법론(작업지시자-AI 페어 프로그래밍)으로 개발.
- **본 저장소가 사용하는 방법:**
  - `../codebase/rhwp` 위치에 별도 git clone (본 저장소에 포함되지 않음)
  - `hwp-automate-poc/Cargo.toml` 과 `hwp-automate-py/Cargo.toml` 이 `path = "../../codebase/rhwp"` 로 의존
  - rhwp 코드는 일절 수정하지 않고 upstream 그대로 사용
- **활용 모듈:** `DocumentCore` (IR 빌더), `parse_document`, `serialize_hwp`, `parser::cfb_reader::LenientCfbReader` (비표준 CFB 메타 lenient 파싱), `serializer::mini_cfb::build_cfb` (CFB v3 writer), `model::control::Control::Table`
- **왜 rhwp 인가:** Mac/Linux/Windows 어디서든 한컴오피스 설치 없이 .hwp 처리가 가능한 유일한 성숙한 오픈소스 엔진. 한글과컴퓨터의 공개 문서를 참고하여 구현된 IR 모델·파서·직렬화기를 그대로 활용합니다.

#### 🪝 [golbin/hop](https://github.com/golbin/hop) — 설계 패턴 출처 (참고만, 직접 의존 안 함)

- **저자:** golbin ([@golbin](https://github.com/golbin))
- **라이선스:** MIT
- **설명:** Tauri 2 기반 macOS/Windows/Linux 데스크톱 HWP 뷰어·에디터. rhwp 를 third_party 서브모듈로 통합한 운영 환경 사례.
- **본 저장소가 흡수한 패턴:**
  - `DocumentCore::from_bytes(bytes)` — 양식 로드 표준 진입점 (hop 의 `editable_core_from_bytes`)
  - `mutate_document(operation, args)` JSON 디스패처 → 본 저장소의 `fill_template(operations=[...])` 다중 op 디자인 영감
  - rhwp 의 raw IR 을 외부에서 만지지 않고 `*_native` 메서드 경유하는 boundary 정책
- **의존 방식:** 코드 의존 없음. 단지 hop 의 코드를 읽고 좋은 패턴을 흡수.
- **왜 hop 패턴인가:** rhwp 를 production 환경에서 어떻게 쓰는지 보여주는 가장 완성도 높은 오픈소스 사례. "rhwp 와의 깔끔한 boundary", "JSON 디스패처", "atomic save" 같은 운영 정책을 자연스럽게 채택할 수 있게 해주었습니다.

#### 🦀 추가 Rust 라이브러리 (rhwp 가 transitive 로 사용)

| 라이브러리 | 역할 | 라이선스 |
|---|---|---|
| [PyO3](https://github.com/PyO3/pyo3) | Rust ↔ Python FFI 바인딩 (직접 의존) | Apache-2.0 / MIT |
| [maturin](https://github.com/PyO3/maturin) | abi3 wheel 빌드 도구 | MIT |
| 그 외 rhwp 의 transitive 의존: cfb, byteorder, zip, quick-xml, encoding_rs, image, usvg, pdf-writer, ttf-parser 등 | 각 crate 의 MIT/Apache-2.0 |

#### 의존성 디렉토리 레이아웃

```
temp_git/                              (사용자 작업 루트)
├── hwpx-generator/                    (이 저장소)
│   ├── hwp-automate-poc/              (Rust PoC, rhwp 사용)
│   ├── hwp-automate-py/               (Python 바인딩, rhwp 사용)
│   └── ...
└── codebase/                          (외부 참조, 본 저장소에 포함 안 됨)
    ├── rhwp/                          ← gh repo clone edwardkim/rhwp
    └── hop/                           ← gh repo clone golbin/hop  (참고용)
```

운영 시 vendor 고정이 필요하면 hop 의 `third_party/rhwp` git submodule 패턴을 따라할 수 있습니다. 현재 PoC 단계에선 path 의존이 더 빠른 반복을 위해 유지.

### 한글 / 한컴 상표 안내

- **"한글", "한컴", "HWP", "HWPX"** 는 주식회사 한글과컴퓨터의 등록 상표입니다.
- 본 프로젝트(hwpx-generator)는 한글과컴퓨터와 제휴, 후원, 승인 관계가 없는 **독립적인 오픈소스 작업**입니다.
- HWP/HWPX 포맷 처리는 한글과컴퓨터의 공개 문서를 참고한 다음 도구들을 활용합니다:
  - 경로 A: lxml (XML 직접 편집) + 한컴오피스 한글 2024 (COM 자동화 — 별도 라이선스 필요)
  - 경로 B: rhwp (오픈소스 Rust 엔진, MIT)

### 외부 재배포 / 공개 시

본 저장소를 외부에 재배포하거나 공개 자산화할 경우 다음을 준수하시기 바랍니다.

1. rhwp 와 hop 의 MIT 라이선스 텍스트 동봉 (또는 명시적 링크)
2. "한글", "한컴", "HWP", "HWPX" 상표 안내 유지
3. rhwp / hop 저자(@edwardkim, @golbin) 의 기여를 본 README 와 동일 수준으로 표기

## 라이선스

이 프로젝트는 내부 업무 자동화 목적으로 개발되었습니다. 외부 의존(rhwp, hop, lxml, pywin32 등) 의 각 라이선스 조건은 위 Acknowledgement 섹션을 참고하세요.
