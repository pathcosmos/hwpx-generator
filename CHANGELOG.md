# Changelog

## [2026-06-08] 동국산업 HWP/HWPX 운영 파이프라인 보강

WSL2 + Windows 한글 COM 환경에서 동국산업 자율형공장 산출물을 안정적으로 점검·변환·채우기 위해 경로 A(Python+lxml+COM)를 운영용으로 확장했다. 이번 변경은 `hwpx-generator` 루트 저장소 범위에 한정하며, `hwp-automate-*` 하위 프로젝트 문서는 수정하지 않았다.

### 핵심 판단

| 항목 | 판단 |
|------|------|
| 문서 변환, PDF 생성, 기존 HWP/HWPX 양식 채우기 | 충분 — Windows 한글 COM 중심으로 진행 |
| 복잡한 서식 보존 | 충분 — 한글 렌더링/SaveAs 엔진 활용 |
| 완전한 범용 HWP/HWPX 편집 엔진 | 부족 — 현재 범위에서는 우선순위 제외 |
| 현재 프로젝트 산출물 자동화 | COM 중심으로 충분 |

진행 우선순위는 `COM 안정화 → 전체 문서 일괄 점검/변환 → 3개 양식 필드맵 → batch 셀 채우기 → PDF 검증`으로 확정.

### 신규 운영 CLI — `src/document_pipeline.py`

동국산업 산출물 운영을 위한 새 CLI를 추가했다.

| 명령 | 용도 |
|------|------|
| `health` | Windows Python, pywin32, 한글 COM 사용 가능 여부와 `Hwp.exe` 전후 프로세스 수 확인 |
| `cleanup-com` | 현재 `Hwp.exe` 목록 조회. `--kill` 지정 시에만 강제 종료 |
| `inspect` | 단일 HWP/HWPX 열기, 페이지 수, HWPX 패키지/표 구조 확인 |
| `convert` | 단일 HWP/HWPX 를 HWPX/HWP/PDF/HTML/TEXT 로 변환 |
| `batch-inspect` | 최상위 `*.hwp` 와 `result/*.hwpx` 전체 점검 리포트 생성 |
| `batch-convert` | 전체 문서 세트를 PDF/HWPX 로 일괄 변환. 기본 출력은 `result/batch_check/` |
| `fill-cell` | HWPX 단일 셀 수정 후 선택적으로 PDF 생성 |
| `fill-cells` | 필드맵 JSON + 데이터 JSON 으로 여러 표/여러 셀 일괄 채우기 |
| `verify-pdf` | PDF 존재, 크기, 페이지 수, 텍스트 추출 가능 여부 확인. 선택적으로 참조 PDF 비교 |

기본 문서 루트는 `../동국산업_자율형공장_중간산출물_문서`로 설정했다.

### COM 안정화 (`src/bridge.py`, `src/hwp_com.py`)

`src/hwp_com.py`:
- `_format_for_path()`, `_normalize_save_format()` 추가
- `open_checked()`, `save_as_checked()` 추가
- `inspect_document()`, `convert_document()` 추가
- SaveAs 결과 파일 존재/크기 검증 추가
- 기존 `open()`, `save()`, `save_as()` API는 유지

`src/bridge.py`:
- `check_com_available()` 추가
- `inspect_document()` 추가
- `convert_document()` 추가
- `list_hwp_processes()` 추가
- `cleanup_hwp_processes(kill=False)` 추가
- Windows inline script 실행 결과를 JSON으로 파싱하는 `_run_inline_script_json()` 추가
- COM 작업 전후 `Hwp.exe` 프로세스 목록/개수 기록
- Windows stdout/stderr 한글 경로 디코딩 보정 (`utf-8` → `cp949` fallback)
- `/tmp` 같은 Windows 접근 불가 경로는 traceback 대신 명시적 실패 결과 반환

주의: `cleanup-com`은 기본적으로 조회만 수행한다. `--kill`을 명시한 경우에만 `taskkill /IM Hwp.exe /F`를 실행한다.

### HWPX 편집기 보강 (`src/hwpx_editor.py`)

HWPX 직접 XML 편집은 계속 안전한 셀 텍스트 치환 중심으로 유지하되, 점검 기능을 추가했다.

- `inspect_package(hwpx_path)` 추가
  - `mimetype` 존재/값/첫 엔트리/압축 방식 확인
  - `Contents/section*.xml` 파싱 확인
  - section 수, 표 수, 문단 수 반환
  - 오류/경고 목록 반환
- `validate_package()` 추가
- `inspect_tables(include_empty=True)` 추가
  - 표 인덱스, 행/열 수, 실제 셀 수, 빈 셀 후보 반환
- `set_cell_text_by_index(table_index, row, col, text)` 추가
- `_cell_text()`, `_int_attr()` 보조 함수 추가

### 동국산업 보고서 필드맵 추가

새 디렉토리: `templates/dongkuk_reports/`

| 파일 | 대상 |
|------|------|
| `weekly_report.json` | `★17 SF-PM2 주간보고서 양식_V1.2.hwpx` |
| `monthly_report.json` | `★18 SF-PM3 월간보고서 양식_V1.2.hwpx` |
| `meeting_minutes.json` | `★19 SF-PM4 회의록 양식_V1.1.hwpx` |

필드맵 형식:
- `form_id`
- `template_path`
- `description`
- `fields[]`
  - `name`
  - `table`
  - `row`
  - `col`
  - `description`
  - `required`

현재 필드 범위:
- 주간보고서: 금주 실적, 차주 계획, 위험/이슈 1~2
- 월간보고서: 금월 실적, 차월 계획, 위험/이슈 1~2
- 회의록: 첫 번째 회의록 표의 일시, 장소, 주최자, 안건, 회의내용, 결과, 특이사항, 참석자

### Batch 셀 채우기 정책

`fill-cells`는 필드맵과 데이터 JSON을 사용한다.

입력 데이터 형식:

```json
{
  "current_week_results": "금주 실적 내용",
  "next_week_plan": "차주 계획 내용"
}
```

적용 정책:
- `required=true` 필드가 누락되면 적용 전 실패
- 선택 필드 누락은 skip
- 좌표가 없거나 셀이 존재하지 않으면 적용 전/적용 중 실패
- 출력 HWPX 패키지 검증 후, `--pdf` 지정 시 COM으로 PDF 생성
- `--verify-pdf` 지정 시 PDF 페이지 수/텍스트 추출까지 확인

### PDF 검증

`verify-pdf` 명령 추가:
- PDF 파일 존재 확인
- 파일 크기 확인
- PyMuPDF 기반 페이지 수 확인
- 텍스트 추출 가능 여부 확인
- `--compare` 지정 시 참조 PDF와 페이지 수 비교
- `--ssim` 지정 시 기존 `pdf_compare.py`를 통한 SSIM/text 비교 수행

### 테스트 추가/수정

새 파일: `tests/test_current_hwpx_pipeline.py`

검증 항목:
- `result/*.hwpx` 3개 HWPX 패키지 무결성
- HWPX 표 인벤토리 조회
- 임시 복사본 셀 수정 후 재파싱
- `document_pipeline.py inspect --no-com`
- `batch-inspect --no-com --no-root-hwp`
- `fill-cells` dry-run
- `fill-cells` 임시 HWPX 적용
- 기존 PDF 검증
- `RUN_HWP_COM_TESTS=1`일 때 Windows 한글 COM health, HWP→HWPX, HWP→PDF 변환

`tests/test_hwpx_editor.py`:
- `ref/test_01.hwpx`가 현재 checkout에 없을 때 skip하도록 보정

검증 결과:

```bash
python3 -m py_compile src/hwp_com.py src/bridge.py src/hwpx_editor.py src/document_pipeline.py tests/test_current_hwpx_pipeline.py tests/test_hwpx_editor.py
pytest -q tests/test_current_hwpx_pipeline.py tests/test_hwpx_editor.py
# 12 passed, 14 skipped
```

실제 COM smoke:
- `python3 src/document_pipeline.py health --json` 성공
- `python3 src/document_pipeline.py cleanup-com --json` 조회 성공 (`--kill` 미사용)
- `/mnt/d/...` 임시 폴더에서 `batch-convert --format pdf --verify-pdf` 성공
- HWP 1개 → PDF 변환 성공, PDF 페이지 수/텍스트 추출 검증 성공
- 변환 중 `Hwp.exe` 프로세스 수 증가 없음

### 변경 파일 요약

| 파일/디렉토리 | 변경 |
|---------------|------|
| `src/document_pipeline.py` | 신규 운영 CLI |
| `src/bridge.py` | COM JSON 브릿지, 프로세스 리포트, cleanup, 디코딩 보정 |
| `src/hwp_com.py` | checked open/save/inspect/convert API |
| `src/hwpx_editor.py` | HWPX 패키지/표 점검 API |
| `templates/dongkuk_reports/` | 동국산업 3개 양식 필드맵 |
| `tests/test_current_hwpx_pipeline.py` | 신규 smoke/integration 테스트 |
| `tests/test_hwpx_editor.py` | 누락 fixture skip 보정 |
| `README.md` | 운영 CLI/필드맵/테스트 문서화 |
| `CLAUDE.md` | 작업자용 운영 규칙 업데이트 |
| `CHANGELOG.md` | 본 변경 내역 기록 |

---

## [2026-04-30] AI 통합 진입점 추가 (Skill + MCP 서버 + 분석 강화)

경로 B (Rust+rhwp) 위에 AI 가 양식을 분석하고 사용자에게 필요 정보를 제안한 후 콘텐츠를 자동 생성·삽입하는 통합 진입점을 도입.

### 분석 강화 (`analyze_template` 확장)

`hwp-automate-py/src/lib.rs` 의 `analyze_template` 결과에 표 단위 상세 추가:

- `cells`: 모든 셀 (row, col, text, is_empty, neighbor_label)
- `empty_cells`: 빈 셀 목록 + 라벨 추론
- `suggested_fields`: 라벨 추론 성공한 빈 셀만 — AI 가 즉시 "어떤 정보가 필요한가" 파악

`find_neighbor_label()` 헬퍼: 같은 행 왼쪽 셀 우선(한국 양식의 라벨-값 가로 페어), 같은 열 위쪽 차순위.

### Claude Code Skill — `fill-hwp`

새 파일들:
- `.claude/skills/fill-hwp/SKILL.md` — multi-turn playbook (분석 → 필요 정보 질문 → 채우기 → 검증)
- `.claude/hooks/hwp-fill-verify.py` — PostToolUse hook (fill 명령 후 출력 파일 자동 확인)
- `.claude/settings.json` — hook 등록

Claude Code 사용자가 `/fill-hwp 양식.hwp` 한 명령으로 즉시 양식 채우기 시작.

### MCP 서버 — `mcp_server.py`

새 파일: `hwp-automate-py/mcp_server.py` — FastMCP stdio 서버. Claude Desktop / Claude Code / Cursor 등 MCP 호환 클라이언트에서 사용 가능.

5 개 tool 노출: `analyze_form`, `preview_form_structure`, `fill_form`, `fill_form_from_data`, `verify_output`.

`pyproject.toml` 변경:
- `[project.optional-dependencies]` 에 `mcp = ["mcp[cli]>=1.2.0; python_version >= '3.10'"]` 추가
- `[tool.pytest.ini_options]` 명시 (testpaths)
- requires-python 은 3.9 유지 (mcp 만 3.10+ 조건부)

### 변경 통계

| 영역 | 변경 |
|---|---|
| `hwp-automate-py/src/lib.rs` | +69 / -7 (analyze 강화 + find_neighbor_label) |
| `hwp-automate-py/mcp_server.py` | +245 (신규) |
| `hwp-automate-py/pyproject.toml` | +8 (mcp optional, pytest config) |
| `.claude/skills/fill-hwp/SKILL.md` | +200 (신규 Skill) |
| `.claude/hooks/hwp-fill-verify.py` | +90 (신규 hook) |
| `.claude/settings.json` | +14 (신규) |
| `CLAUDE.md` | +60 (AI 통합 진입점 섹션) |

### 검증

- 기존 SVG 회귀 3 testcase 통과 그대로
- 실 양식 (YCP/코리녹스) fill_template 회귀 통과
- 강화된 analyze_template 가 YCP 기본정보 표의 빈 셀 6 개에 대해 라벨 (업종명/주생산품/매출액/영업이익/수출액/부채비율) 모두 정확 추론
- MCP 서버 5 tool 모두 등록·in-process 호출 검증

---

## [2026-04-29 #2] V1 실 양식 검증 + BinData 보존 우회 + 상세 Acknowledgement

실 사업신청서 양식 (35MB / 54MB) 으로 검증하면서 발견한 두 가지 한계를 수정하고, 상세 출처 표기를 추가.

### 발견·수정한 두 가지 함정

1. **셀 병합으로 `cell_idx = row × cols + col` 어긋남**
   - 5×6 = 30 셀 (병합 없음) 양식에선 우연히 맞았으나 7×8 (병합으로 27셀만) 양식에서 실패
   - 수정: `find_cell_idx()` 가 표의 cells Vec 에서 (row, col) 위치를 직접 검색. pre-flight 에서 한 번 산출, 캐시.

2. **HWP 셀 텍스트 trailing whitespace 자동 추가**
   - 사용자 입력 "단원구" 가 라운드트립 후 "단원구 " 로 보여 verify 실패
   - 수정: post-fill verify 비교를 `trim_end()` 적용

### 추가: BinData 보존 우회 (`preserve_images=True` 기본 활성)

rhwp v0.7.x 의 BinData (BMP 이미지 등) 라운드트립 충실도 한계 — 일부 양식에서 한컴이 "손상" 으로 판정 또는 그림 일부 누락. **rhwp 자체의 `LenientCfbReader` + `mini_cfb::build_cfb` 를 활용한 stream-level 머지 우회법**을 도입.

- rhwp output 베이스 + BinData/Preview 만 입력 양식에서 byte-for-byte 보존
- 결과: BinData 54/54 동일 크기, 출력 크기 ≈ 입력 크기 (이전 ~5MB 손실 → 현재 ~80KB 차이)
- HWP 표준 layout (`leaf_to_hwp_path`) 으로 LenientCfbReader 의 leaf-only 경로를 storage path 로 재구성

### 상세 Acknowledgement 추가 (3 개 README)

`README.md`, `hwp-automate-poc/README.md`, `hwp-automate-py/README.md` 모두에 다음 추가:

- **rhwp** ([@edwardkim](https://github.com/edwardkim/rhwp)): 활용 모듈 매트릭스, 의존 방식, 라이선스 (MIT)
- **hop** ([@golbin](https://github.com/golbin/hop)): 흡수한 4 개 패턴, 의존 방식 (코드 의존 없음)
- 한글/한컴 상표 안내, 외부 재배포 의무 (저자 표기, MIT 라이선스 동봉, 상표 안내 유지)

### 변경 통계

| 영역 | 변경 |
|---|---|
| `hwp-automate-py/src/lib.rs` | +238 / -23 (cell_idx 위치 검색, trim_end verify, merge_cfb_preserving_input, leaf_to_hwp_path) |
| `hwp-automate-py/Cargo.toml` | +2 (cfb crate 의존 추가/제거 사이클 후 rhwp internal 모듈 사용) |
| `README.md` | +83 (Acknowledgement) |
| `hwp-automate-poc/README.md` | +67 (Acknowledgement) |
| `hwp-automate-py/README.md` | +74 (Acknowledgement) |
| **합계** | **+441 / -23** |

---

## [2026-04-29] Rust + rhwp 자동화 경로 추가 (cross-platform, COM 불필요)

**커밋**: `39e4070` — Add Rust+rhwp automation path (cross-platform HWP filling)
**PR**: [#1](https://github.com/pathcosmos/hwpx-generator/pull/1)

기존 **경로 A (Python + lxml + COM)** 와 나란히 동작하는 **경로 B (Rust + rhwp)** 를 추가. macOS / Linux / Windows 어디서나 한컴오피스 설치 없이 .hwp(HWP 5.0 binary) 양식을 자동 채우기 가능.

> **원칙**: from-scratch 로 새 문서를 만드는 패턴은 노출하지 않는다. 사용자 양식을 베이스로 빈 셀에만 값 삽입하는 패턴만 지원.

---

### 신규 sub-project

| 디렉토리 | 역할 |
|---------|------|
| `hwp-automate-poc/` | **Rust binary** — 기존 양식의 표 자동 채우기 데모 (헤더 매칭 → 컬럼 자동 탐색 → 라운드트립 검증) |
| `hwp-automate-py/` | **PyO3 abi3-py39 Python 바인딩** — Python 3.9~3.14 어디서나 단일 wheel 사용 |
| `hwp-automate-py/hwp_automate_cli/` | Python 보조 도구 — field_map.json 어댑터 + argparse CLI (analyze / fill / cell) |

### 외부 의존 (별도 git clone, 본 repo 비포함)

| 위치 | 출처 | 라이선스 |
|-----|------|---------|
| `../codebase/rhwp/` | [edwardkim/rhwp](https://github.com/edwardkim/rhwp) | MIT |
| `../codebase/hop/` | [golbin/hop](https://github.com/golbin/hop) (`DocumentCore::from_bytes` 패턴 출처) | MIT |

### Python API (PyO3 노출, `hwp_automate.*`)

| 함수 | 용도 |
|------|------|
| `analyze_template(path)` | 양식 표·스타일·번호 인벤토리 (read-only) |
| `fill_template(template, out, operations, dry_run=False, verify=True)` ★ | 다중 표·다중 컬럼·다중 셀 일괄 채우기. **Pre-flight + post-fill verify + dry_run** |
| `fill_template_table(template, out, mapping, ...)` | 단일 표·단일 컬럼 편의 wrapper |

### CLI

```bash
python -m hwp_automate_cli analyze --template 양식.hwp
python -m hwp_automate_cli fill    --template 양식.hwp --field-map ... --data ... --output ...
python -m hwp_automate_cli cell    --template 양식.hwp --output ... --header-match 성명 --cell 1,5,값
```

### 안전 메커니즘

- **Pre-flight 검증** — 모든 operation 의 표·컬럼·범위가 유효한지 적용 전에 확인. 잘못된 op 1개라도 있으면 양식 무수정 (silent corruption 방지).
- **Post-fill 라운드트립 검증** — 저장 후 재파싱하여 모든 셀 값이 정확히 보존됐는지 자동 확인. 불일치 시 `RuntimeError`.
- **`dry_run` 모드** — 실제 적용·저장 없이 plan 만 검증·반환.

### CLAUDE.md 업데이트

- 두 경로 비교표 (프로젝트 개요)
- 경로 B 전용 섹션 (위치, 사용법, 검증된 능력 5가지, 한계, 함정, 진입점)
- 기존 절들에 적용 범위 표기 (`HWPX 파일 수정 시 주의사항`, `COM 자동화 주의사항`)

---

### 수치 요약

| 항목 | 값 |
|------|---|
| 변경 파일 | 15 개 |
| 추가 행 | +4,189 |
| 삭제 행 | −1 |
| 신규 sub-project | 2 개 (Rust binary + Python 바인딩) |

### 검증

- 7 단계 자동 회귀 통과: analyze, fill_template 다중 op, dry_run, pre-flight 보호, CLI 호출, field_map.json 어댑터, legacy 호환
- 사용자 시각 검증 통과 (한컴에서 poc_v3.hwp 열어 확인)
- biz_plan.hwp 8 개 표 자동 발견, 5×6 인력표 자격증 4 셀 100% 라운드트립

### 알려진 한계 / 향후

- **Mac arm64 wheel 만 현재 빌드** → GitHub Actions matrix 로 macOS+Linux+Windows 자동 빌드 예정
- **rhwp v0.7.x SVG 렌더러는 outline 자동 번호 미렌더** → 한컴/모바일에서는 정상 표시
- 실 양식 검증은 사용자 양식 1 개와 함께 진행 예정

---

## [2026-03-08] Two-Pass Hybrid Form Filler Pipeline

**커밋**: `222eb7d` — HWPX form filler pipeline: two-pass hybrid system

기존 단순 XML 셀 채우기 + COM 찾아바꾸기 방식에서, **마크다운 문서를 파싱하여 HWPX 양식에 자동으로 채워넣는 2-Pass 하이브리드 파이프라인**으로 대폭 확장.

---

### 신규 모듈 (핵심)

| 파일 | 역할 |
|------|------|
| `src/form_filler.py` | **파이프라인 오케스트레이터** — Pass 1(XML 직접 편집) + Pass 2(COM 서식 삽입)를 순차 실행하는 메인 엔트리포인트 |
| `src/md_parser.py` | **마크다운 파서** — 사업계획서 `.md` 파일을 구조화된 블록(헤딩, 문단, 표, 리스트)으로 파싱 |
| `src/md_to_ops.py` | **마크다운→COM 변환기** — 파싱된 블록을 COM 자동화 명령(InsertText, 서식 적용 등) 시퀀스로 변환. 계층적 들여쓰기, 표→텍스트 변환 포함 |
| `src/section_mapper.py` | **섹션 매퍼** — 마크다운 섹션 번호를 HWPX 양식의 마커(##SEC1_CONTENT## 등)에 매핑 |

### 신규 템플릿

| 파일 | 역할 |
|------|------|
| `templates/gyeongnam_rbd/field_map.json` | **경남 R&BD 사업계획서 전용 필드맵** — 36개 표, 184개 셀의 좌표 매핑 (표지, 요약, 시장현황, 수요처, 목표, KPI, 연구원, 생산계획, 예산, 기관현황, 부속서류) |
| `data/form_content_map.json` | **양식 콘텐츠 매핑** — 마크다운 섹션 → HWPX 마커 대응표 |

### 신규 도구 (감사/디버깅/유틸리티)

| 파일 | 역할 |
|------|------|
| `audit_crossrefs.py` | HWPX 내 교차참조(charPrIDRef, paraPrIDRef 등) 유효성 검사 |
| `audit_hwpx_content.py` | 생성된 HWPX 파일의 콘텐츠 무결성 감사 (빈 셀, 누락 마커 탐지) |
| `audit_section0.py` | section0.xml 상세 감사 — 표/셀 구조, 텍스트 내용 점검 |
| `compare_section0.py` | 원본 vs 수정된 section0.xml 비교 (구조적 diff) |
| `compare_section0_v2.py` | 개선된 section0 비교 — 셀 단위 세밀 비교 |
| `debug_crash_isolate.py` | COM 크래시 원인 격리 — 섹션별로 나눠 실행하며 크래시 지점 특정 |
| `diagnose_xml_serialization.py` | XML 직렬화 문제 진단 — 선언부, 네임스페이스, 인코딩 검증 |
| `tools/make_rawcopy.py` | HWPX 원본 복사 유틸리티 — ZIP 엔트리별 압축방식 보존하며 클린 카피 생성 |
| `_bridge_test.py` | WSL↔Windows 브릿지 연결 테스트 |

### 기존 모듈 업데이트 (25개 파일)

#### 핵심 변경

| 파일 | 주요 변경 내용 |
|------|--------------|
| `src/bridge.py` | **COM 포스트 포맷 패턴** 구현 — `InsertText` → 선택 → `char_shape` 적용 → 해제. 0.1pt 렌더링 버그 해결. 2-tier 선택 최적화 (MoveParaBegin/End vs MoveSelLeft×N) |
| `src/hwpx_editor.py` | 다중 `hp:p`/`hp:run` 클리어링, charPrIDRef 보존 강화, 마커 삽입 기능 추가 |
| `src/hwp_com.py` | 포스트 포맷 기반 텍스트 삽입, 계층적 들여쓰기, 표→텍스트 렌더링 지원 |
| `src/extract_template.py` | 경남 R&BD 양식(38페이지, 36개 표) 분석 지원 확대 |
| `src/generate_hwpx.py` | form_filler 파이프라인 통합, 마커 기반 콘텐츠 삽입 흐름 추가 |
| `src/field_mapper.py` | 경남 R&BD 필드맵 지원, 다중 기관/기업 리스트 매핑 |

#### 분석 문서 업데이트

| 파일 | 변경 |
|------|------|
| `analysis/approach_comparison.md` | 2-Pass 하이브리드 방식 평가 결과 추가 |
| `analysis/com_evaluation.md` | 포스트 포맷 패턴 발견 및 검증 결과 기록 |
| `analysis/direct_xml_evaluation.md` | 마커 삽입 방식의 한계와 해결책 기록 |
| `analysis/hwpx_structure_analysis.md` | 경남 R&BD 양식 36개 표 구조 분석 추가 |
| `analysis/pyhwpx_evaluation.md` | 최종 평가 업데이트 |

#### 기타

| 파일 | 변경 |
|------|------|
| `.gitignore` | 출력/임시 파일 패턴 추가 |
| `CLAUDE.md` | 2-Pass 파이프라인 아키텍처, COM 포스트 포맷 주의사항 반영 |
| `README.md` | 전체 리팩토링 (아래 별도 기술) |
| `data/sample_input.json` | 경남 R&BD 양식에 맞는 샘플 데이터로 교체 |
| `data/schema.json` | 경남 R&BD 입력 스키마로 업데이트 |
| `templates/cloud_integrated/*` | 기존 템플릿 호환성 유지 보수 |
| `tests/*` | 테스트 코드 업데이트 — 새 모듈 대응 |

---

### 수치 요약

| 항목 | 값 |
|------|---|
| 변경 파일 수 | 40개 |
| 추가 행 | +11,646 |
| 삭제 행 | -5,534 |
| 순증 행 | +6,112 |
| 신규 파일 | 15개 |
| 수정 파일 | 25개 |

### Pass 1 결과 (XML 직접 편집)

- 14개 핵심 표에 184개 셀 채움
- 5개 콘텐츠 마커 삽입 (`##SEC1-5_CONTENT##`)
- 36개 표 구조 무손상, ZIP 무결성 검증 통과

### Pass 2 결과 (COM 자동화)

- 21페이지 → 65페이지로 확장 (콘텐츠 삽입)
- 5개 섹션, 총 6,396개 COM 명령 실행
- 모든 폰트 사이즈 정상 (0.1pt 버그 해결)
- 실행 시간: 약 8분
