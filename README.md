# HWPX Generator

한컴오피스 한글(Hangul) 문서를 프로그래밍 방식으로 생성하는 도구입니다. 정부 사업계획서 등 복잡한 양식 문서의 자동 생성을 목표로 합니다.

**현재 상태**: Milestone 2 — 다중 템플릿 지원
- 커버 페이지 93개 셀 XML 직접 채우기
- COM 기반 텍스트 찾아바꾸기
- PDF 자동 변환
- SSIM 검증 통과 (0.9660, 기준 0.90)
- **다중 템플릿 지원** — 임의의 HWPX 파일을 템플릿으로 등록하여 사용 가능

## 주요 기능

- **다중 템플릿 지원** — 임의의 HWPX 파일을 템플릿으로 등록, 설정 기반 문서 생성
- **템플릿 기반 HWPX 생성** — XML 셀 직접 수정 + COM 텍스트 교체의 하이브리드 방식
- **PDF 자동 변환** — 한컴오피스 COM API를 통한 정확한 PDF 출력
- **PDF 비교 검증** — SSIM(구조적 유사도) + 텍스트 일치도 자동 비교
- **템플릿 구조 분석** — HWPX 파일의 표/셀/스타일 구조 추출 + 설정 초안 자동 생성

## 아키텍처

### 전체 파이프라인

```
┌─────────────────────────────────────────────────────────────┐
│                      WSL (Ubuntu)                           │
│                                                             │
│  ┌──────────┐    ┌──────────────┐    ┌───────────────┐      │
│  │  JSON    │───>│ field_mapper │───>│  hwpx_editor  │      │
│  │  Input   │    │  (매핑)      │    │  (XML 수정)   │      │
│  └──────────┘    └──────────────┘    └──────┬────────┘      │
│                                             │               │
│                                    ┌────────▼────────┐      │
│                                    │    bridge.py    │      │
│                                    │ (WSL↔Win 브릿지) │      │
│                                    └────────┬────────┘      │
│                                             │               │
├─────────────────────────────────────────────┼───────────────┤
│                   Windows                   │               │
│                                    ┌────────▼────────┐      │
│                                    │   hwp_com.py   │      │
│                                    │  (COM 자동화)   │      │
│                                    └────────┬────────┘      │
│                                        ┌────┴────┐          │
│                                        ▼         ▼          │
│                                     .hwpx      .pdf         │
├─────────────────────────────────────────────────────────────┤
│                      WSL (검증)                              │
│                                    ┌─────────────────┐      │
│                                    │  pdf_compare.py │      │
│                                    │  (SSIM 검증)    │      │
│                                    └─────────────────┘      │
└─────────────────────────────────────────────────────────────┘
```

### 모듈 구성

| 모듈 | 실행 환경 | 역할 |
|------|----------|------|
| `src/generate_hwpx.py` | WSL | 메인 CLI 파이프라인. 전체 흐름 제어 |
| `src/bridge.py` | WSL | WSL↔Windows Python 브릿지. COM 스크립트 실행 |
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

### Windows가 필수인 이유

한컴오피스의 COM API만이 다음을 보장합니다:
- 119페이지, 456개 표, 63개 이미지, 370개 글자 속성의 **100% 서식 재현**
- 내장 렌더링 엔진을 통한 **정확한 PDF 변환** (`SaveAs "PDF"`)
- 기존 문서를 열어서 수정하는 **템플릿 기반 워크플로우**

크로스플랫폼 대안(python-hwpx, 직접 XML 생성)은 서식 재현과 PDF 변환에서 한계가 있습니다. 자세한 비교는 [접근방식 비교](#구현-접근방식-비교)를 참고하세요.

### 필수 구성

**Windows 측**:
- Windows 10/11
- 한컴오피스 2024 (한글)
- Python 3.13+ (Microsoft Store 또는 python.org)
- pywin32 (`pip install pywin32`)

**WSL 측**:
- Ubuntu (WSL2 권장)
- Python 3.12+
- 패키지: `lxml`, `PyMuPDF(fitz)`, `scikit-image`, `Pillow`, `numpy`

### WSL↔Windows 브릿지

WSL Python이 `bridge.py`를 통해 Windows Python(`python.exe`)을 `subprocess`로 호출합니다. Windows Python은 한컴오피스 COM API에 접근하여 문서를 열고, 수정하고, PDF로 저장합니다. 결과 파일은 `/mnt/d/` 등 공유 드라이브를 통해 WSL에서 접근 가능합니다.

## 설치 및 실행

### 1. 전제조건 확인

```bash
# WSL에서 확인
python3 --version        # 3.12+
# Windows Python 경로 확인
/mnt/c/Users/<username>/AppData/Local/Microsoft/WindowsApps/python.exe --version
```

### 2. WSL Python 패키지 설치

```bash
pip3 install --break-system-packages lxml pymupdf Pillow scikit-image numpy
```

### 3. Windows Python 패키지 설치

```powershell
# Windows PowerShell에서
pip install pywin32
```

### 4. 실행

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

# 템플릿을 그대로 PDF로 변환 (검증용)
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --output output/ \
  --pdf-only

# PDF 비교 검증 포함
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --output output/ \
  --pdf-only \
  --compare ref/test_01.pdf

# PDF 생성 없이 HWPX만 생성
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --data data/sample_input.json \
  --output output/ \
  --no-pdf
```

### CLI 옵션

| 옵션 | 설명 |
|------|------|
| `--template, -t` | (필수) 템플릿 HWPX 파일 경로 |
| `--template-dir` | 템플릿 설정 디렉토리 (기본: `templates/cloud_integrated/`) |
| `--data, -d` | JSON 입력 데이터 파일 경로 |
| `--output, -o` | 출력 디렉토리 (기본: `output`) |
| `--pdf-only` | 데이터 없이 템플릿을 그대로 PDF로 변환 |
| `--no-pdf` | PDF 생성 건너뛰기 |
| `--compare, -c` | 비교할 참조 PDF 경로 |

## 다중 템플릿 지원

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

HWPX의 XML은 12개 이상의 네임스페이스를 사용합니다. `ElementTree`는 네임스페이스 프리픽스를 `ns0:`, `ns1:`로 변환하여 한글이 파일을 인식하지 못합니다. **lxml** 사용이 필수입니다 — 원본 프리픽스(`hp:`, `hs:`, `hh:` 등)를 그대로 보존합니다.

### HWPX ZIP 구조

HWPX는 ZIP 아카이브이며, `mimetype` 파일이 **반드시 비압축(STORED)**이어야 합니다. 압축 시 한글이 파일을 열지 못합니다.

## 프로젝트 구조

```
hwpx_generator/
├── README.md
├── CLAUDE.md                           # Claude Code 가이드
├── analysis/                           # 접근방식 평가 보고서
│   ├── approach_comparison.md          #   3가지 방식 종합 비교
│   ├── com_evaluation.md               #   COM API 테스트 결과
│   ├── pyhwpx_evaluation.md            #   python-hwpx 평가
│   ├── direct_xml_evaluation.md        #   직접 XML 평가
│   └── hwpx_structure_analysis.md      #   HWPX 구조 분석
├── data/
│   ├── sample_input.json               # 샘플 입력 데이터
│   └── schema.json                     # 입력 데이터 JSON Schema
├── ref/
│   ├── test_01.hwpx                    # 참조 HWPX (템플릿 원본)
│   ├── test_01.hwp                     # 참조 HWP
│   └── test_01.pdf                     # 참조 PDF (검증 기준)
├── src/
│   ├── __init__.py
│   ├── generate_hwpx.py                # 메인 CLI 파이프라인
│   ├── bridge.py                       # WSL↔Windows 브릿지
│   ├── hwp_com.py                      # 한컴오피스 COM 자동화
│   ├── hwpx_editor.py                  # HWPX XML 편집기 (lxml)
│   ├── field_mapper.py                 # JSON→셀 좌표 매핑
│   ├── pdf_compare.py                  # PDF 비교 검증
│   └── extract_template.py             # 템플릿 구조 분석
├── templates/
│   └── cloud_integrated/
│       ├── template.json               # 템플릿 메타 정보 + 찾아바꾸기 매핑
│       └── field_map.json              # 커버 페이지 필드 매핑 설정
├── tests/
│   ├── test_hwpx_editor.py             # HwpxEditor 단위 테스트 (11개)
│   ├── test_hwp_com_module.py          # COM 모듈 테스트
│   └── test_integration.py             # 통합 테스트
└── output/                             # 생성 결과물 (gitignore 대상)
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

### 단위 테스트

```bash
# HwpxEditor 테스트 (11개)
python3 -m pytest tests/test_hwpx_editor.py -v
```

테스트 항목:
- 표 조회 및 범위 외 인덱스 처리
- 셀 조회 및 텍스트 설정
- 일괄 셀 채우기
- 저장 후 재로드 검증
- 네임스페이스 보존 확인
- charPrIDRef(서식 참조) 보존 확인
- mimetype 비압축(STORED) 확인

### 통합 테스트 (전체 파이프라인)

```bash
python3 src/generate_hwpx.py \
  --template ref/test_01.hwpx \
  --data data/sample_input.json \
  --output output/milestone1 \
  --compare ref/test_01.pdf
```

검증 기준:
- SSIM >= 0.90 (달성: 0.9660)
- 페이지 수 일치: 119/119
- 텍스트 일치도: 0.9959

## HWPX 형식 참고

- HWPX는 `application/hwp+zip` MIME 타입의 ZIP 아카이브
- 단위: HWPUNIT (1/7200 인치). A4 = 59528 x 84188
- 주요 네임스페이스: `hp:`(문단), `hs:`(섹션), `hh:`(헤더), `hc:`(코어), `ha:`(앱)
- 상세 구조: [`CLAUDE.md`](CLAUDE.md) 및 [`analysis/hwpx_structure_analysis.md`](analysis/hwpx_structure_analysis.md) 참고

## 라이선스

이 프로젝트는 내부 업무 자동화 목적으로 개발되었습니다.
