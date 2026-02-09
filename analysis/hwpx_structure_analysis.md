# HWPX Reference File Structure Analysis

**File**: `ref/test_01.hwpx`
**Document**: 2026년 클라우드 종합솔루션 지원사업 사업계획서 (Korean Government Cloud Solution Business Proposal)
**Created**: 2022-12-27 by TIPA, Last modified: 2026-02-09 by lanco
**App Version**: Hancom Office Hangul 12.30.0.6163 (MAC64LEDarwin_25.2.0)
**HWPML Version**: 1.5

---

## 1. ZIP Archive Structure

The `.hwpx` file is a ZIP archive with mimetype `application/hwp+zip`.

### File Listing (74 files total)

| Path | Size (bytes) | Compressed | Description |
|------|-------------|-----------|-------------|
| `mimetype` | 19 | No (STORE) | Must be first entry, uncompressed |
| `version.xml` | 312 | No | HWPML version info |
| `META-INF/container.xml` | 475 | DEFLATE | OPF rootfile references |
| `META-INF/container.rdf` | 867 | DEFLATE | RDF metadata about document parts |
| `META-INF/manifest.xml` | 134 | DEFLATE | ODF manifest (empty in this file) |
| `Contents/content.hpf` | 7,413 | DEFLATE | OPF package manifest |
| `Contents/header.xml` | 728,960 | DEFLATE | Document head: fonts, styles, properties |
| `Contents/section0.xml` | 4,343,577 | DEFLATE | Document body (main content) |
| `settings.xml` | 958 | DEFLATE | Application settings |
| `Preview/PrvText.txt` | 1,809 | DEFLATE | Plain-text preview |
| `Preview/PrvImage.png` | 157,161 | No | Thumbnail image |
| `BinData/image1-63` | ~various | Mixed | 63 embedded images (PNG/BMP) |

### Image Distribution
- **PNG**: 15 images (image1, 10, 14, 19, 20, 28, 31-33, 42-50)
- **BMP**: 48 images (rest), all DEFLATE compressed
- Total embedded image data: ~230MB uncompressed

---

## 2. XML Namespaces (Shared Across All XML Files)

```xml
xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"
xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"
xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"
xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"
xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:opf="http://www.idpf.org/2007/opf/"
xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"
xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"
xmlns:epub="http://www.idpf.org/2007/ops"
xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
```

---

## 3. Metadata Files

### 3.1 mimetype
```
application/hwp+zip
```

### 3.2 version.xml
```xml
<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"
  tagetApplication="WORDPROCESSOR"
  major="5" minor="1" micro="1" buildNumber="0" os="10"
  xmlVersion="1.5"
  application="Hancom Office Hangul"
  appVersion="12.30.0.6163 MAC64LEDarwin_25.2.0"/>
```

### 3.3 META-INF/container.xml
```xml
<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container">
  <ocf:rootfiles>
    <ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>
    <ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>
    <ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/>
  </ocf:rootfiles>
</ocf:container>
```

### 3.4 META-INF/container.rdf
Declares document structure relationships:
- `Contents/header.xml` -> type `pkg#HeaderFile`
- `Contents/section0.xml` -> type `pkg#SectionFile`
- Root -> type `pkg#Document`

### 3.5 META-INF/manifest.xml
Empty ODF manifest: `<odf:manifest/>`

### 3.6 settings.xml
```xml
<ha:HWPApplicationSetting>
  <ha:CaretPosition listIDRef="0" paraIDRef="73" pos="0"/>
  <config:config-item-set name="PrintInfo">
    <config:config-item name="PrintMethod" type="short">4</config:config-item>
    <config:config-item name="ZoomX" type="short">100</config:config-item>
    <config:config-item name="ZoomY" type="short">100</config:config-item>
    <!-- ... print settings ... -->
  </config:config-item-set>
</ha:HWPApplicationSetting>
```

---

## 4. Contents/content.hpf (OPF Package)

### Metadata
- **Title**: "솔루션 개요"
- **Language**: ko
- **Creator**: TIPA
- **Last saved by**: lanco
- **Created**: 2022-12-27T05:42:05Z
- **Modified**: 2026-02-09T02:51:56Z

### Manifest Items
- `header` -> `Contents/header.xml`
- `section0` -> `Contents/section0.xml`
- `settings` -> `settings.xml`
- `image1` through `image63` -> `BinData/imageN.{png|bmp}` (all with `isEmbeded="1"`)

### Spine Order
1. `header` (linear=yes)
2. `section0` (linear=yes)

---

## 5. Contents/header.xml Analysis (728KB)

Root: `<hh:head version="1.5" secCnt="1">`

### 5.1 Beginning Numbers
```xml
<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>
```

### 5.2 Font Faces (7 language categories)

| Language | Font Count | Fonts |
|----------|-----------|-------|
| HANGUL | 12 | 굴림, 돋움체, 맑은 고딕, 함초롬돋움, 함초롬바탕, 휴먼명조, HY헤드라인M, 신명 신문명조, 한양중고딕, CIDFont+F4, CIDFont+F6, 폴라리스새바탕-함초롬바탕호환 |
| LATIN | 13 | (same + HCI Poppy) |
| HANJA | 13 | (similar, + 한양신명조) |
| JAPANESE | 13 | (similar) |
| OTHER | 11 | (subset) |
| SYMBOL | 12 | (subset) |
| USER | 12 | (subset, includes 명조) |

Font types: TTF (TrueType), HFT (Hancom Font)
Some fonts have `substFont` (substitute font) definitions pointing to 한컴바탕.

### 5.3 Character Properties (charPr) - 370 definitions

Each `charPr` has:
- **Attributes on element**: `id`, `height` (font size in 1/100 pt), `textColor`, `shadeColor`, `useFontSpace`, `useKerning`, `symMark`, `borderFillIDRef`
- **Child elements**: `fontRef`, `ratio`, `spacing`, `relSz`, `offset`, `underline`, `strikeout`, `outline`, `shadow`
- **Optional children**: `bold` (empty element = bold on), `italic` (empty element = italic on)

**Font Size Distribution** (height in 1/100 pt):
| Size | Point | Count |
|------|-------|-------|
| 1000 | 10pt | 134 |
| 1100 | 11pt | 113 |
| 900 | 9pt | 32 |
| 1200 | 12pt | 27 |
| 1300 | 13pt | 19 |
| 500 | 5pt | 13 |
| 800 | 8pt | 7 |
| 1500 | 15pt | 4 |
| 1600 | 16pt | 3 |

**Text Color Distribution**:
| Color | Count | Usage |
|-------|-------|-------|
| #000000 (black) | 223 | Body text |
| #0000FF (blue) | 97 | Instructions/guidelines |
| #FF0000 (red) | 23 | Emphasis |
| #808080 (gray) | 5 | De-emphasized |

**Formatting**: 98 bold definitions, 64 italic definitions out of 370 total.

**Example charPr Structure**:
```xml
<hh:charPr id="2" height="1400" textColor="#000000" shadeColor="none"
           useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">
  <hh:fontRef hangul="2" latin="2" hanja="2" japanese="2" other="2" symbol="2" user="2"/>
  <hh:ratio hangul="100" latin="100" .../>
  <hh:spacing hangul="0" latin="0" .../>
  <hh:relSz hangul="100" latin="100" .../>
  <hh:offset hangul="0" latin="0" .../>
  <hh:bold/>  <!-- empty element = bold ON -->
  <hh:underline type="NONE" shape="SOLID" color="#000000"/>
  <hh:strikeout shape="NONE" color="#000000"/>
  <hh:outline type="NONE"/>
  <hh:shadow type="NONE" color="#B2B2B2" offsetX="10" offsetY="10"/>
</hh:charPr>
```

### 5.4 Paragraph Properties (paraPr) - 241 definitions

Each `paraPr` has:
- `align` with `horizontal` (JUSTIFY, CENTER, LEFT, RIGHT) and `vertical` (BASELINE)
- `heading` with `type` (NONE, OUTLINE), `idRef`, `level`
- `breakSetting` with `breakLatinWord`, `breakNonLatinWord` (KEEP_WORD/BREAK_WORD), `widowOrphan`, `keepWithNext`, `lineWrap`
- `autoSpacing` with `eAsianEng`, `eAsianNum`
- `switch` with `case`/`default` for HwpUnitChar namespace conditional processing
- `border` with `borderFillIDRef`, offsets, `connect`, `ignoreMargin`

**Note**: Line spacing and margins are defined inside the `switch/case` block for HwpUnitChar namespace, making them conditionally processed.

### 5.5 Border Fills - 178 definitions

Each `borderFill` has:
- `id`, `threeD`, `shadow`, `centerLine`, `breakCellSeparateLine`
- `slash`, `backSlash` elements
- `leftBorder`, `rightBorder`, `topBorder`, `bottomBorder` with `type` (NONE/SOLID/DOUBLE_SLIM), `width` (e.g., "0.12 mm"), `color`
- `diagonal` element
- Optional `fillBrush` containing `winBrush` with `faceColor`, `hatchColor`, `alpha`

**Common Border Patterns**:
- id=1: All borders NONE (invisible border)
- id=3: All borders SOLID 0.12mm black (standard table border)
- id=5: Mixed borders with gray fill (#D6D6D6) - table header style
- id=8: All borders SOLID 0.12mm black (another standard pattern)

### 5.6 Styles - 23 definitions

| ID | Type | Korean Name | English Name | paraPrIDRef | charPrIDRef |
|----|------|------------|-------------|-------------|-------------|
| 0 | PARA | 바탕글 | Normal | 1 | 18 |
| 1 | PARA | 본문 | Body | 22 | 18 |
| 2 | PARA | 개요 1 | Outline 1 | 23 | 18 |
| 3 | PARA | 개요 2 | Outline 2 | 24 | 18 |
| 4 | PARA | 개요 3 | Outline 3 | 25 | 18 |
| 5 | PARA | 개요 4 | Outline 4 | 26 | 18 |
| 6 | PARA | 개요 5 | Outline 5 | 27 | 18 |
| 7 | PARA | 개요 6 | Outline 6 | 28 | 18 |
| 8 | PARA | 개요 7 | Outline 7 | 29 | 18 |
| 9 | CHAR | 쪽 번호 | Page Number | 0 | 17 |
| 10 | PARA | 머리말 | Header | 30 | 19 |
| 11 | PARA | 각주 | Footnote | 31 | 20 |
| 12 | PARA | 미주 | Endnote | 31 | 20 |
| 13 | PARA | 메모 | Memo | 32 | 21 |
| 14 | PARA | 차례 제목 | TOC Heading | 33 | 22 |
| 15-17 | PARA | 차례 1-3 | TOC 1-3 | 34-36 | 23 |
| 18 | PARA | 바탕글 사본3 | Normal Copy3 | 1 | 18 |
| 19 | PARA | 바탕글 사본1 | Normal Copy1 | 4 | 24 |
| 21 | PARA | (custom) | | | |

### 5.7 Numbering Definitions - 6 definitions

Each numbering definition has 10 levels with `paraHead` sub-elements:
- `numFormat` values: DIGIT, HANGUL_SYLLABLE, CIRCLED_DIGIT, CIRCLED_HANGUL_SYLLABLE
- `textOffsetType`: PERCENT, `textOffset`: 50
- `autoIndent`: 1 (auto indent enabled)

### 5.8 Bullet Definitions - 2 definitions

Both bullets use the dash character (`-`), `useImage="0"`.

### 5.9 Tab Properties - 9 definitions

---

## 6. Contents/section0.xml Analysis (4.3MB)

Root: `<hs:sec>` with 540 direct child `<hp:p>` (paragraph) elements.

### 6.1 Page Properties

```xml
<hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY">
  <hp:margin header="2834" footer="2834" gutter="0"
             left="5669" right="5669" top="1417" bottom="1417"/>
</hp:pagePr>
```

- **Page size**: A4 (59528 x 84188 HWPUNIT = 210mm x 297mm)
- **Landscape**: WIDELY (portrait)
- **Margins**: left/right=5669 (~20mm), top/bottom=1417 (~5mm), header/footer=2834 (~10mm)

### 6.2 Element Counts

| Element Type | Count |
|-------------|-------|
| `<hp:p>` paragraphs | 5,828 (540 top-level + nested in tables) |
| `<hp:tbl>` tables | 456 |
| `<hp:pic>` pictures | 66 |
| `<hp:t>` text nodes | 3,893 (184 non-empty) |
| `<hp:run>` runs | varies per paragraph |

### 6.3 Document Structure (Top-level paragraphs)

The document follows this structure:
1. **Cover page table** (35x11, first cell = "2026년 클라우드 종합솔루션 지원사업 사업계획서")
2. **Instructions table** (2x2, "작성 요령")
3. **Table of Contents** paragraphs
4. **Section 1**: 솔루션 구축 개요 (Solution Construction Overview)
   - 1.1 개발배경 및 필요성 (Development Background and Necessity)
   - 1.2 개발기관 및 업종 현황 (Development Institutions and Industry Status)
   - 1.3 요구사항 분석 (Requirements Analysis)
5. **Section 2**: 솔루션 개발내용 (Solution Development Contents)
   - 2.1 개발개요
   - 2.2 주요 개발내용
   - 2.3 클라우드서비스 적용계획
6. **Section 3**: 솔루션 도입효과 (Solution Introduction Effects)
   - 3.1 핵심성과지표(KPI)
   - 3.2 KPI 측정방법
   - 3.3 추가성과지표
7. **Section 4**: 사업화 지원계획 (Commercialization Support Plan)
8. **Section 5**: 표준기반 솔루션 개발 (Standard-Based Development)
9. **Section 6**: 산출물 예정목록 (Expected Deliverables)
10. **Section 7**: 추진일정 (Schedule)
11. **Section 8**: 사업비 구성 (Budget Composition)
12. **Section 9**: 참여인력 (Participating Personnel)

### 6.4 Table Analysis (456 tables)

**Major tables in this document**:

| Table | Rows x Cols | Width | First Cell Content | Cell Merges | Purpose |
|-------|------------|-------|-------------------|-------------|---------|
| 0 | 35x11 | 48140x75106 | "2026년 클라우드 종합솔루션..." | 118 | Cover page form |
| 1 | 2x2 | 46202x4596 | "작성 요령" | 1 | Instructions box |
| 2 | 1x1 | 47100x67480 | "[ 사업 참여 배경 ]" | 0 | Content frame |
| 3 | 1x1 | 47100x74968 | "(디지털 전환 시대...)" | 0 | Content frame |
| 5 | 5x3 | 43104x31455 | "구분" | 0 | Data comparison table |
| 13 | 6x3 | 43483x32370 | "핵심 목표" | 0 | Objectives table |
| 15 | 6x4 | 45263x41776 | "영역" | 0 | Area details table |

**Table patterns observed**:
- **1x1 tables**: Used as content frames/boxes (like div containers in HTML)
- **NxM tables**: Used for actual data tabulation with headers
- **Cover table (35x11)**: Complex form layout with extensive cell merging

**Table structure**:
```xml
<hp:tbl id="..." rowCnt="35" colCnt="11" cellSpacing="0" borderFillIDRef="3"
        noAdjust="1" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" pageBreak="TABLE">
  <hp:sz width="48140" widthRelTo="ABSOLUTE" height="75106" heightRelTo="ABSOLUTE"/>
  <hp:pos treatAsChar="1" vertRelTo="PARA" horzRelTo="COLUMN"/>
  <hp:outMargin left="141" right="141" top="141" bottom="141"/>
  <hp:inMargin left="140" right="140" top="140" bottom="140"/>
  <hp:tr>
    <hp:tc name="" header="0" borderFillIDRef="127">
      <hp:subList textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER">
        <hp:p paraPrIDRef="74" styleIDRef="0">
          <hp:run charPrIDRef="2"><hp:t>Cell text</hp:t></hp:run>
          <hp:linesegarray>...</hp:linesegarray>
        </hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="11" rowSpan="1"/>
      <hp:cellSz width="47898" height="1963"/>
      <hp:cellMargin left="141" right="141" top="141" bottom="141"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

### 6.5 Picture/Image Analysis (66 pictures)

Images are referenced via `<hp:img binaryItemIDRef="imageN">` which maps to `BinData/imageN.{png|bmp}` as declared in `content.hpf`.

**Picture element structure**:
```xml
<hp:pic id="..." zOrder="..." numberingType="PICTURE" textWrap="TOP_AND_BOTTOM">
  <hp:offset x="..." y="..."/>
  <hp:orgSz width="..." height="..."/>       <!-- original size -->
  <hp:curSz width="..." height="..."/>       <!-- current display size -->
  <hp:flip horizontal="0" vertical="0"/>
  <hp:rotationInfo angle="0" .../>
  <hp:renderingInfo>
    <hp:transMatrix .../> <hp:scaMatrix .../> <hp:rotMatrix .../>
  </hp:renderingInfo>
  <hp:img binaryItemIDRef="image1" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>
  <hp:imgRect>
    <hp:pt0 x="0" y="0"/><hp:pt1 .../><hp:pt2 .../><hp:pt3 .../>
  </hp:imgRect>
  <hp:imgClip left="0" right="..." top="0" bottom="..."/>
  <hp:imgDim dimwidth="..." dimheight="..."/>
  <hp:sz width="..." widthRelTo="ABSOLUTE" height="..." heightRelTo="ABSOLUTE"/>
  <hp:pos treatAsChar="1" vertRelTo="PARA" horzRelTo="PARA" .../>
  <hp:shapeComment>그림입니다.\n원본 그림의 이름: image1.png\n...</hp:shapeComment>
  <hp:caption side="BOTTOM" ...>  <!-- optional -->
    <hp:subList>...</hp:subList>
  </hp:caption>
</hp:pic>
```

Key observations:
- Most images use `treatAsChar="1"` (inline with text)
- Images have both original (`orgSz`) and current (`curSz`) size info
- `shapeComment` contains Korean description with original filename and pixel dimensions
- Some images have `caption` elements with `subList` containing paragraphs

### 6.6 Style Reference Usage

**Top paraPrIDRef values** (most frequently used paragraph properties):
| paraPrIDRef | Usage Count | Likely Purpose |
|------------|-------------|---------------|
| 13 | 1,452 | Default table cell (center align) |
| 1 | 570 | Normal body text (justify) |
| 39 | 447 | Table content |
| 64 | 344 | Table content variant |
| 3 | 291 | Center-aligned |
| 190 | 219 | Specific formatting |
| 2 | 208 | Center-aligned |
| 63 | 201 | Content |
| 74 | 161 | Table header |

**styleIDRef distribution** (11 unique values):
| styleIDRef | Count | Style Name |
|-----------|-------|-----------|
| 0 | 5,444 | 바탕글 (Normal) |
| 21 | 186 | (custom) |
| 18 | 65 | 바탕글 사본3 (Normal Copy3) |
| 4 | 46 | 개요 3 (Outline 3) |
| 3 | 27 | 개요 2 (Outline 2) |
| 2 | 19 | 개요 1 (Outline 1) |
| 13 | 18 | 메모 (Memo) |

**Top charPrIDRef values** (most frequently used character properties):
| charPrIDRef | Usage Count | Properties |
|------------|-------------|-----------|
| 312 | 707 | (specific formatting) |
| 14 | 425 | 맑은 고딕, 10pt, black |
| 39 | 278 | |
| 329 | 238 | |
| 36 | 216 | |
| 3 | 174 | 맑은 고딕, 10pt, black (no bold) |

### 6.7 Paragraph Structure

Each paragraph follows this pattern:
```xml
<hp:p id="..." paraPrIDRef="1" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <!-- One or more runs -->
  <hp:run charPrIDRef="60">
    <!-- Optional section properties (first paragraph only) -->
    <hp:secPr>...</hp:secPr>
    <!-- Optional control elements -->
    <hp:ctrl>...</hp:ctrl>
    <!-- Optional table -->
    <hp:tbl>...</hp:tbl>
    <!-- Optional picture -->
    <hp:pic>...</hp:pic>
    <!-- Text content -->
    <hp:t>Text content here</hp:t>
  </hp:run>
  <!-- Line segment array (rendering hint) -->
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1400" textheight="1400"
                baseline="1190" spacing="420" horzpos="0" horzsize="47860"
                flags="393216"/>
  </hp:linesegarray>
</hp:p>
```

**Key**: A paragraph can contain mixed runs with different character properties, and runs can contain tables, pictures, or text.

### 6.8 Headers/Footers
- **No headers or footers** found in this document
- No masterPage elements
- No headerFooter elements
- No footnotes or endnotes

---

## 7. Key Insights for HWPX Generation

### 7.1 Minimum Required Files
1. `mimetype` (uncompressed, first entry)
2. `version.xml`
3. `META-INF/container.xml`
4. `META-INF/container.rdf`
5. `META-INF/manifest.xml`
6. `Contents/content.hpf`
7. `Contents/header.xml`
8. `Contents/section0.xml`
9. `settings.xml`

### 7.2 Critical Format Rules
- **mimetype must be the first ZIP entry** and stored uncompressed (compress_type=0)
- All XML files use `UTF-8` encoding with `standalone="yes"` declaration
- HWPUNIT: 1/7200 inch (A4 = 59528 x 84188)
- Font size in charPr `height` attribute: value in 1/100 of a point (1000 = 10pt)
- Border widths specified as strings like "0.12 mm"
- Style IDs are numeric, referenced by various `*IDRef` attributes
- `linesegarray` appears to be a rendering hint (may be recalculated by Hangul on open)

### 7.3 ID Reference Chain
```
paragraph (hp:p)
  ├── paraPrIDRef -> header.xml paraPr[id]
  │     └── borderFillIDRef -> header.xml borderFill[id]
  ├── styleIDRef -> header.xml style[id]
  │     ├── paraPrIDRef -> header.xml paraPr[id]
  │     └── charPrIDRef -> header.xml charPr[id]
  └── run (hp:run)
        └── charPrIDRef -> header.xml charPr[id]
              ├── fontRef[hangul] -> header.xml fontface[lang=HANGUL]/font[id]
              └── borderFillIDRef -> header.xml borderFill[id]
```

### 7.4 Table Cell Structure
```
hp:tbl (borderFillIDRef for table-level border)
  └── hp:tr
       └── hp:tc (borderFillIDRef for cell-level border)
            ├── hp:subList (vertAlign, lineWrap)
            │    └── hp:p (standard paragraph with runs)
            ├── hp:cellAddr (colAddr, rowAddr)
            ├── hp:cellSpan (colSpan, rowSpan)
            ├── hp:cellSz (width, height)
            └── hp:cellMargin (left, right, top, bottom)
```

### 7.5 Complexity Assessment for Reproduction

| Feature | Complexity | Notes |
|---------|-----------|-------|
| Basic text paragraphs | Low | Simple hp:p + hp:run + hp:t structure |
| Bold/italic formatting | Low | Empty child elements in charPr |
| Font selection | Medium | Requires fontface + fontRef mapping per language |
| Tables (simple) | Medium | Row/col structure with cell properties |
| Tables (merged cells) | High | cellSpan + cellAddr coordination |
| Images | Medium | Need binary data + img reference chain |
| Cover page form | Very High | 35x11 table with 118 merged cells |
| Full style chain | High | 370 charPr + 241 paraPr + 178 borderFill definitions |
| linesegarray | Unknown | May be required or may be auto-generated |

### 7.6 Content Summary (Korean)

**문서 제목**: 2026년 클라우드 종합솔루션 지원사업 사업계획서

**솔루션 이름**: 웹 기반 설계자산(도면-문서) 통합 관리와 상태기반(CBM) 설비보전 기술이 결합된 LLM 융합 중소제조기업형 SaaS 플랫폼

**개발 솔루션 5가지**:
1. 상태기반 설비보전시스템 (eCMMS)
2. 공장에너지관리시스템 (FEMS)
3. 웹기반 엔지니어링 설계 자산(CAD)관리시스템
4. 클라우드 기반 문서관리 시스템
5. RAG 기반 LLM 지식검색 서비스

**수행기간**: 솔루션 개발 '26.6.30~'27.6.30 (12개월), 실증 '27.6.30~'27.12.31 (6개월)
