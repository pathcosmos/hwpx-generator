# python-hwpx 라이브러리 방식 평가

## 개요

- **라이브러리**: python-hwpx v1.9 (PyPI)
- **설치**: `pip install python-hwpx`
- **의존성**: lxml (내부적으로 lxml 사용하지만, 직렬화는 stdlib ET)
- **접근 방식**: 스켈레톤 HWPX 템플릿 기반 + ElementTree XML 조작

## 아키텍처

```
hwpx/
  document.py      - HwpxDocument: 사용자용 고수준 API
  package.py        - HwpxPackage: ZIP 아카이브 관리
  templates.py      - Skeleton.hwpx 템플릿 로드
  data/Skeleton.hwpx - 빈 문서 템플릿 (11KB, 완전한 HWPX 구조)
  oxml/
    document.py     - XML 요소 조작 (119KB, 핵심 로직)
    header.py       - 헤더(스타일) 파싱
    body.py         - 단락/표 모델
    parser.py       - XML 파싱 유틸리티
```

## 테스트 결과

### 1. 스타일 적용 (test_pyhwpx_styles.py) - 11/11 PASS

| 기능 | 결과 | 방법 |
|------|------|------|
| 기본 단락 추가 | PASS | `doc.add_paragraph("텍스트")` |
| 굵게 (Bold) | PASS | `doc.ensure_run_style(bold=True)` -> charPrIDRef |
| 기울임 (Italic) | PASS | `doc.ensure_run_style(italic=True)` |
| 밑줄 (Underline) | PASS | `doc.ensure_run_style(underline=True)` |
| 복합 서식 | PASS | `ensure_run_style(bold=True, italic=True, underline=True)` |
| 글자 크기 변경 | PASS | charPr height 속성 직접 수정 (1000=10pt, 2000=20pt) |
| 폰트 변경 | PASS | charPr의 fontRef 속성 수정 |
| 문단 정렬 (가운데) | PASS | paraPr의 align horizontal 속성 수정 |
| 줄간격 변경 | PASS | paraPr의 lineSpacing 수정 (type=PERCENT, value=200) |
| 글자 색상 | PASS | charPr textColor 속성 (#FF0000) |

**참고**: `ensure_run_style()`은 bold/italic/underline만 지원. 폰트 크기, 색상 등은 header XML의 charPr를 직접 수정해야 함.

### 2. 표 기능 (test_pyhwpx_table.py) - 9/9 PASS

| 기능 | 결과 | 방법 |
|------|------|------|
| 표 생성 | PASS | `doc.add_table(rows, cols)` |
| 셀 텍스트 설정 | PASS | `tbl.cell(r, c).text = "..."` |
| 행 병합 | PASS | `tbl.merge_cells(0, 0, 0, 3)` |
| 2x2 블록 병합 | PASS | `tbl.merge_cells(0, 0, 1, 1)` |
| 사용자 정의 크기 | PASS | `add_table(rows, cols, width=42520, height=7200)` |
| 셀 크기 조작 | PASS | `cell.set_size(width, height)` |
| 셀 테두리/배경 | PASS | borderFill XML 직접 추가 (fillBrush/windowBrush) |
| 참조문서형 표 | PASS | 병합 + 텍스트 조합 |
| 그리드 검사 | PASS | `tbl.get_cell_map()` |

### 3. 종합 문서 (test_pyhwpx_complex.py) - 8/8 PASS

| 기능 | 결과 | 방법 |
|------|------|------|
| 페이지 설정 | PASS | `props.set_page_size()`, `props.set_page_margins()` |
| 제목 (20pt 볼드 가운데) | PASS | 커스텀 charPr + paraPr |
| 부제목 | PASS | 별도 charPr + paraPr |
| 섹션 제목 | PASS | 별도 charPr + paraPr |
| 본문 텍스트 | PASS | 별도 charPr + paraPr |
| 병합표 | PASS | add_table + merge_cells |
| 추가 콘텐츠 | PASS | 다중 단락 |
| 머리글/바닥글 | PASS | `doc.set_header_text()`, `doc.set_footer_text()` |

## 장점

1. **기본 기능 완비**: 텍스트, 표, 병합, 머리글/바닥글 모두 동작
2. **스켈레톤 기반**: 유효한 HWPX 구조를 자동으로 보장 (20개의 기본 paraPr, 7개 charPr, 2개 borderFill 포함)
3. **API 편의성**: `add_paragraph()`, `add_table()`, `cell().text` 등 직관적 API
4. **셀 병합**: `merge_cells()` API로 복잡한 표 구조 가능
5. **템플릿 조작**: 기존 HWPX 파일을 열어서 수정하는 것도 가능 (`HwpxDocument.open()`)
6. **다양한 도구**: TextExtractor, ObjectFinder 등 분석 도구 내장
7. **한글 호환**: 생성된 파일이 한컴오피스에서 열림 (이전 테스트에서 확인)

## 단점 / 제한사항

1. **네임스페이스 프리픽스 변환**: 직렬화 시 `hp:` -> `ns0:`, `hs:` -> `ns1:` 등으로 변환됨 (stdlib ET 사용). 의미상 동일하나, 한글이 엄격한 검증을 하면 문제 가능성 있음
2. **고급 서식 API 부재**: `ensure_run_style()`은 bold/italic/underline만 지원. 폰트 크기, 색상, 폰트 종류 변경은 header XML 직접 수정 필요
3. **이미지 삽입 API 없음**: 이미지 추가 기능이 내장되어 있지 않음 (package에 바이너리 추가 후 XML에서 참조는 가능할 수 있으나 미검증)
4. **문서 수준 스타일 관리 복잡**: charPr, paraPr를 직접 XML로 조작해야 하는 부분이 많음
5. **lxml 의존성**: 내부적으로 lxml 사용 (일부 기능)

## 서식 재현도 평가

| 항목 | 가능 여부 | 비고 |
|------|-----------|------|
| 폰트 지정 | O | charPr fontRef 수정 |
| 글자 크기 | O | charPr height 수정 (1000 단위) |
| 굵기/기울임/밑줄 | O | ensure_run_style() API |
| 글자 색상 | O | charPr textColor 수정 |
| 문단 정렬 | O | paraPr align horizontal 수정 |
| 줄간격 | O | paraPr lineSpacing 수정 |
| 표 생성 | O | add_table() |
| 셀 병합 | O | merge_cells() |
| 셀 배경색 | O | borderFill에 fillBrush 추가 |
| 셀 테두리 | O | borderFill 커스텀 |
| 머리글/바닥글 | O | set_header_text(), set_footer_text() |
| 이미지 삽입 | ? | 직접 구현 필요 |
| 페이지 번호 | ? | 필드 코드로 구현 필요 |
| 목차 | X | 미지원 |

## PDF 변환

- python-hwpx 자체에는 PDF 변환 기능 없음
- PDF 변환을 위해서는:
  1. 한컴오피스 COM 자동화로 HWPX 열기 -> PDF 저장 (Windows 전용)
  2. LibreOffice + HWP 필터 사용 (리눅스, 제한적)
  3. 별도 변환 서비스 구축

## 개발 복잡도: 중간

- 기본 텍스트/표는 쉬움 (API 직접 사용)
- 고급 서식은 XML 수준 이해 필요 (charPr, paraPr, borderFill 직접 조작)
- 참조 문서 수준의 복잡한 표 재현은 상당한 작업 필요

## 유지보수성: 양호

- PyPI 패키지로 관리 (버전 관리, 업데이트 가능)
- 코드베이스가 잘 구조화되어 있음
- 하위 수준 XML 조작 부분은 문서화가 부족

## 생성 파일 크기

- 빈 문서: ~12KB (스켈레톤)
- 스타일 테스트: ~65KB
- 표 테스트: ~112KB
- 종합 테스트: ~81KB

## 결론

python-hwpx는 프로그래밍 방식으로 HWPX를 생성하는 가장 실용적인 방법. 기본 기능(텍스트, 표, 병합, 머리글)은 API로 잘 제공되며, 고급 서식은 XML 직접 조작으로 구현 가능. 네임스페이스 프리픽스 변환 이슈를 제외하면, 참조 문서 수준의 문서 생성이 가능할 것으로 판단됨.
