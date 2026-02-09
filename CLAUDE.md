# CLAUDE.md

이 파일은 Claude Code 및 AI 에이전트가 이 저장소에서 작업할 때 반드시 숙지해야 할 사항을 정리한다.

## 프로젝트 개요

한컴오피스 한글(Hangul)의 XML 기반 문서 형식인 HWPX 파일을 프로그래밍 방식으로 생성하는 도구. 정부 사업계획서, 주간/월간 보고서 등 복잡한 양식 문서의 자동 생성을 목표로 한다.

**핵심 전략**: XML 직접 수정 + COM 텍스트 교체/PDF 변환의 **하이브리드 방식**.

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

## HWPX 파일 수정 시 필수 주의사항

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
