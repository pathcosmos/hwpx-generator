# 직접 XML 생성 방식 평가

## 개요

- **방식**: Python의 xml.etree.ElementTree + zipfile 모듈로 HWPX 파일 직접 생성
- **의존성**: 표준 라이브러리만 사용 (외부 패키지 불필요)
- **선택적 의존성**: Jinja2 (템플릿 기반 생성 시)

## 테스트 결과 (test_direct_xml.py) - 9/9 PASS

| 테스트 | 결과 | 설명 |
|--------|------|------|
| mimetype 생성 | PASS | `application/hwp+zip` (19 bytes) |
| version.xml 생성 | PASS | HWPML 1.5 버전 명시 |
| header.xml 생성 | PASS | 폰트, charPr, paraPr, borderFill 정의 (7,386 bytes) |
| section0.xml 생성 | PASS | 단락 3개 + 3x3 표 포함 (9,273 bytes) |
| META-INF 파일 생성 | PASS | container.xml, manifest.xml, container.rdf |
| content.hpf 생성 | PASS | OPF 매니페스트 (1,236 bytes) |
| 완전한 HWPX 조립 | PASS | 10개 파일, 9,423 bytes, mimetype 첫 번째 |
| 참조 파일 비교 | PASS | 공통 10개, 누락 1개 (PrvImage.png만) |
| Jinja2 템플릿 | PASS | 529자 렌더링 성공 |

## HWPX 파일 구조 분석

### 필수 구성 파일

```
mimetype                     # "application/hwp+zip" (ZIP 첫 항목, 비압축)
version.xml                  # <hc:version version="1.5">
Contents/
  content.hpf                # OPF 패키지 매니페스트
  header.xml                 # 문서 헤드: 폰트, charPr, paraPr, borderFill, 스타일
  section0.xml               # 문서 본문: 단락, 표, 레이아웃
META-INF/
  container.xml              # OPF rootfile 참조
  container.rdf              # RDF 메타데이터
  manifest.xml               # ODF 매니페스트
settings.xml                 # 애플리케이션 설정
Preview/
  PrvText.txt                # 텍스트 미리보기
  PrvImage.png               # 썸네일 (선택)
```

### 핵심 XML 요소

#### header.xml (스타일 정의)

```xml
<hh:head>
  <hh:beginNum page="1" ... />
  <hh:refList>
    <hh:fontfaces>         <!-- 폰트 목록 (7개 언어별) -->
    <hh:borderFills>       <!-- 테두리/배경 스타일 -->
    <hh:charProperties>    <!-- 글자 속성 (height, textColor, fontRef, bold, ...) -->
    <hh:paraProperties>    <!-- 문단 속성 (align, lineSpacing, margin, ...) -->
    <hh:styles>            <!-- 스타일 정의 (charPr + paraPr 조합) -->
    <hh:tabProperties>     <!-- 탭 설정 -->
    <hh:numberings>        <!-- 번호 매기기 -->
  </hh:refList>
</hh:head>
```

#### section0.xml (문서 내용)

```xml
<hs:sec>
  <hp:p paraPrIDRef="0" styleIDRef="0">      <!-- 단락 -->
    <hp:run charPrIDRef="0">                   <!-- 실행 (글자 속성 참조) -->
      <hp:secPr>...</hp:secPr>                 <!-- 섹션 속성 (첫 단락에만) -->
      <hp:t>텍스트 내용</hp:t>                  <!-- 텍스트 -->
    </hp:run>
  </hp:p>
  <hp:p>
    <hp:run>
      <hp:tbl rowCnt="3" colCnt="3" borderFillIDRef="2">  <!-- 표 -->
        <hp:sz width="..." height="..." />
        <hp:pos treatAsChar="1" ... />
        <hp:tr>
          <hp:tc borderFillIDRef="2">
            <hp:subList>
              <hp:p><hp:run><hp:t>셀 텍스트</hp:t></hp:run></hp:p>
            </hp:subList>
            <hp:cellAddr colAddr="0" rowAddr="0" />
            <hp:cellSpan colSpan="1" rowSpan="1" />
            <hp:cellSz width="..." height="..." />
          </hp:tc>
        </hp:tr>
      </hp:tbl>
    </hp:run>
  </hp:p>
</hs:sec>
```

## 장점

1. **완전한 제어**: XML 구조를 100% 제어 가능. 한컴 규격과 정확히 일치하는 출력 생성 가능
2. **네임스페이스 보존**: `ET.register_namespace()`로 올바른 프리픽스(hp:, hs:, hh:) 유지
3. **외부 의존성 없음**: Python 표준 라이브러리만 사용 (선택적으로 Jinja2)
4. **경량**: 생성 파일 크기 최소화 (~9KB vs python-hwpx ~65KB)
5. **참조 호환**: 참조 문서와 동일한 구조 재현 가능
6. **Jinja2 통합 가능**: 템플릿 기반으로 다양한 문서 유형을 빠르게 생성

## 단점 / 제한사항

1. **높은 초기 개발 비용**: 모든 XML 요소를 수동으로 작성해야 함
2. **HWPML 스펙 이해 필수**: 누락된 속성이나 잘못된 구조로 한글에서 열리지 않을 위험
3. **셀 병합 로직 직접 구현**: cellAddr, cellSpan, cellSz를 모두 수동 관리
4. **스타일 ID 관리**: charPrIDRef, paraPrIDRef, borderFillIDRef의 정합성을 수동으로 보장
5. **검증 도구 부재**: 생성된 HWPX의 유효성을 검증하는 자동화 수단 없음
6. **엣지 케이스**: 한글이 기대하는 세부 속성(linesegarray 등)을 모두 파악하기 어려움

## Jinja2 템플릿 방식 평가

### 장점
- 문서 구조가 정해진 경우 매우 효율적
- 데이터-템플릿 분리로 유지보수 용이
- 비프로그래머도 템플릿 수정 가능

### 단점
- 동적 표 구조(가변 행/열)에는 복잡한 루프 로직 필요
- XML 이스케이핑 주의 필요 (Jinja2 `|e` 필터 사용)
- 복잡한 셀 병합은 템플릿으로 표현하기 어려움

### 예시

```python
from jinja2 import Template

section_template = Template("""
<hs:sec xmlns:hp="..." xmlns:hs="...">
{% for para in paragraphs %}
<hp:p id="{{ para.id }}" paraPrIDRef="{{ para.para_pr }}" styleIDRef="0"
      pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="{{ para.char_pr }}">
    <hp:t>{{ para.text | e }}</hp:t>
  </hp:run>
</hp:p>
{% endfor %}
</hs:sec>
""")
```

## 참조 문서 비교

| 항목 | 참조 (test_01.hwpx) | 직접 생성 |
|------|---------------------|-----------|
| 파일 수 | 74개 (이미지 63개 포함) | 10개 |
| 메타 파일 | 11개 | 10개 (PrvImage.png 미포함) |
| header.xml | 43KB (128개 borderFill, 74개 charPr, 78개 paraPr) | 7KB (2개 borderFill, 1개 charPr, 1개 paraPr) |
| section0.xml | 대형 (35행x11열 표 포함) | 9KB (3단락 + 3x3 표) |
| 네임스페이스 | hp:, hs:, hh:, hc:, ha: 등 | 동일 (register_namespace 사용) |

## 서식 재현도 평가

| 항목 | 가능 여부 | 난이도 | 비고 |
|------|-----------|--------|------|
| 폰트 지정 | O | 중 | fontface + charPr fontRef 매핑 |
| 글자 크기 | O | 하 | charPr height 속성 |
| 굵기/기울임/밑줄 | O | 하 | charPr 내 bold/italic/underline 요소 |
| 글자 색상 | O | 하 | charPr textColor 속성 |
| 문단 정렬 | O | 하 | paraPr align horizontal |
| 줄간격 | O | 중 | paraPr margin lineSpacingType/lineSpacing |
| 표 생성 | O | 중 | tbl/tr/tc 구조 직접 작성 |
| 셀 병합 | O | 상 | cellAddr/cellSpan/cellSz 정합성 관리 |
| 셀 배경색 | O | 중 | borderFill fillBrush 추가 |
| 셀 테두리 | O | 중 | borderFill leftBorder/rightBorder/... |
| 머리글/바닥글 | O | 중 | secPr 내 header/footer 요소 |
| 이미지 삽입 | O | 상 | BinData/ + content.hpf 매니페스트 + XML 참조 |
| 페이지 번호 | O | 상 | 필드 코드 (fieldBegin/fieldEnd) |

## PDF 변환

- 직접 XML 방식 자체에는 PDF 변환 기능 없음
- python-hwpx와 동일하게 외부 도구 필요

## 개발 복잡도: 높음

- 초기 세팅: HWPML 스펙 분석 + 보일러플레이트 XML 작성 필요
- 표 생성: 셀 좌표, 크기, 병합 로직을 모두 직접 구현
- 스타일 관리: ID 기반 참조 시스템을 직접 관리
- 테스트: 생성된 파일을 한글로 열어서 검증 필요

## 유지보수성: 보통

- 코드 가독성이 높음 (XML 구조가 명시적)
- 한글 버전 업데이트 시 스펙 변경 추적 필요
- Jinja2 사용 시 템플릿과 로직이 분리되어 유지보수 용이

## 하이브리드 접근 제안

**python-hwpx + 직접 XML 조합**이 가장 효율적:

1. **기본 구조**: python-hwpx의 `HwpxDocument.new()` 또는 `open()`으로 시작 (스켈레톤 활용)
2. **고수준 작업**: python-hwpx API 사용 (add_paragraph, add_table, merge_cells)
3. **세부 서식**: header XML 직접 수정 (charPr, paraPr, borderFill 추가)
4. **특수 기능**: package.set_part()로 커스텀 XML 삽입

이 접근은 python-hwpx의 유효한 HWPX 구조 보장 + 직접 XML의 세밀한 제어를 모두 활용.

## 결론

직접 XML 생성은 HWPX 형식에 대한 완전한 제어를 제공하지만, 초기 개발 비용이 높고 엣지 케이스 처리가 어렵다. 순수 직접 XML 방식보다는 **python-hwpx를 기반으로 하되, 부족한 부분을 직접 XML 조작으로 보완하는 하이브리드 방식**을 권장한다.
