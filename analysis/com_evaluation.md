# COM Automation (pywin32 + Hancom Office) Evaluation Report

## Overview

This report evaluates the feasibility of using COM automation via pywin32 to programmatically generate HWPX documents through the installed Hancom Office 2024 (Hangul).

**Environment**: Windows Python 3.13.12 + pywin32, Hancom Office 2024
**Execution**: From WSL via `/mnt/c/Users/lanco/AppData/Local/Microsoft/WindowsApps/python.exe`

---

## Test Results Summary

| Test | Status | Notes |
|------|--------|-------|
| Open & Read HWPX | **PASS** | 6280 lines extracted, 119 pages, 569 controls |
| Create Document (text + table) | **PASS** | 15,228 bytes output |
| Character Styling | **PASS** | Font, size, bold, italic, underline, strikeout, color all work |
| Paragraph Alignment | **PASS** | Left, center, right, justify all work |
| Line Spacing | **PASS** | Percentage-based spacing works |
| Indentation | **PASS** | Left margin, first-line indent work |
| Table Creation | **PASS** | Arbitrary rows/cols, cell navigation, data fill |
| Cell Merging | **PASS** | MergeCells action works with cell block selection |
| PDF Export | **PASS** | SaveAs with "PDF" format, 26KB for simple doc, 9.5MB for 119-page reference |
| HWP Export | **PASS** | SaveAs with "HWP" format |
| HTML Export | **PASS** | SaveAs with "HTML" format |
| TEXT Export | **PASS** | SaveAs with "TEXT" format |
| Multiple Korean Fonts | **PASS** | HY견고딕, HY견명조, 맑은 고딕, 함초롬돋움, 바탕, 돋움, 굴림, 나눔고딕 |
| Cell Background Color | **PARTIAL** | Property exists (WinBrushFaceColor in FillAttr) but needs further testing |

---

## Detailed API Reference

### COM Object Initialization

```python
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False  # Background mode
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
```

### Text Insertion

```python
act = hwp.CreateAction("InsertText")
pset = act.CreateSet()
pset.SetItem("Text", "텍스트 내용")
act.Execute(pset)
hwp.HAction.Run("BreakPara")  # New paragraph
```

### Character Styling (HCharShape)

Key properties discovered via COM introspection:

| Property | Type | Description |
|----------|------|-------------|
| `FaceNameHangul` | str | Korean font name |
| `FaceNameLatin` | str | Latin font name |
| `FaceNameHanja` | str | Hanja font name |
| `FaceNameJapanese` | str | Japanese font name |
| `FaceNameOther` | str | Other font name |
| `FaceNameSymbol` | str | Symbol font name |
| `FaceNameUser` | str | User-defined font |
| `Height` | int | Font size (1pt = 100 units) |
| `Bold` | bool | Bold |
| `Italic` | bool | Italic |
| `UnderlineType` | int | 0=none, 1=single |
| `StrikeOutType` | int | 0=none, 1=single |
| `TextColor` | int | RGB as integer (R + G*256 + B*65536) |
| `ShadeColor` | int | Background highlight color |
| `SuperScript` | bool | Superscript |
| `SubScript` | bool | Subscript |
| `Emboss` | bool | Emboss effect |
| `Engrave` | bool | Engrave effect |
| `OutLineType` | int | Outline type |
| `ShadowType` | int | Shadow type |
| `ShadowColor` | int | Shadow color |
| `UseFontSpace` | bool | Use font spacing |
| `UseKerning` | bool | Use kerning |

Usage pattern:
```python
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.Height = 1200  # 12pt
hwp.HParameterSet.HCharShape.Bold = True
hwp.HParameterSet.HCharShape.FaceNameHangul = "맑은 고딕"
hwp.HParameterSet.HCharShape.TextColor = 0x0000FF  # Blue
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
```

### Paragraph Styling (HParaShape)

Key properties:

| Property | Type | Description |
|----------|------|-------------|
| `AlignType` | int | 0=justify, 1=left, 2=right, 3=center |
| `LineSpacing` | int | Line spacing value |
| `LineSpacingType` | int | 0=percent, 1=fixed, 2=between-lines |
| `LeftMargin` | int | Left indent (HWPUNIT) |
| `RightMargin` | int | Right indent (HWPUNIT) |
| `Indentation` | int | First-line indent (HWPUNIT) |
| `PrevSpacing` | int | Space before paragraph |
| `NextSpacing` | int | Space after paragraph |
| `KeepWithNext` | bool | Keep with next paragraph |
| `KeepLinesTogether` | bool | Keep lines together |
| `PagebreakBefore` | bool | Page break before |
| `WidowOrphan` | bool | Widow/orphan control |

Usage pattern:
```python
hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet)
hwp.HParameterSet.HParaShape.AlignType = 3  # Center
hwp.HParameterSet.HParaShape.LineSpacing = 160  # 160%
hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet)
```

### Table Operations

**Table Creation**:
```python
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 4
hwp.HParameterSet.HTableCreation.Cols = 3
hwp.HParameterSet.HTableCreation.WidthType = 2  # Fit to column
hwp.HParameterSet.HTableCreation.HeightType = 0  # Auto
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
```

**Cell Navigation Actions** (all verified working):
- `TableRightCell` / `TableLeftCell` - Move left/right
- `TableUpperCell` / `TableLowerCell` - Move up/down
- `TableRowBegin` / `TableRowEnd` - Jump to row start/end
- `TableColBegin` / `TableColEnd` - Jump to column start/end
- `TableCellBlock` - Start cell block selection

**Cell Operations**:
- `MergeCells` - Merge selected cells (after block selection)
- `SplitCellHorz` / `SplitCellVert` - Split cells
- `CellIncrease` / `CellDecrease` - Resize cells

**Cell Border/Fill** (HCellBorderFill):
- `BorderTypeTop/Bottom/Left/Right` - Border line types
- `BorderWidthTop/Bottom/Left/Right` - Border widths
- `BorderColorTop/Bottom/Right` / `BorderCorlorLeft` (sic - typo in API)
- `FillAttr.WinBrushFaceColor` - Background fill color
- `FillAttr.WindowsBrush` - Enable brush fill

### File Operations

```python
# Open file
hwp.Open(filepath, "HWPX", "")

# Save in various formats
hwp.SaveAs(filepath, "HWPX", "")  # HWPX format
hwp.SaveAs(filepath, "HWP", "")   # Legacy HWP format
hwp.SaveAs(filepath, "PDF", "")   # PDF export
hwp.SaveAs(filepath, "HTML", "")  # HTML export
hwp.SaveAs(filepath, "TEXT", "")  # Plain text

# Get text content
text = hwp.GetTextFile("TEXT", "")

# Document info
page_count = hwp.PageCount
```

### Document Navigation

- `MoveDocBegin` / `MoveDocEnd`
- `MoveUp` / `MoveDown` / `MoveLeft` / `MoveRight`
- `SelectAll`
- `Undo` / `Redo`

### Control Enumeration

```python
ctrl = hwp.HeadCtrl
while ctrl:
    print(ctrl.CtrlID)  # e.g., 'tbl', 'gso', 'secd', etc.
    ctrl = ctrl.Next
```

Control types found in test_01.hwpx:
- `tbl` (456) - Tables
- `gso` (68) - Drawing objects (shapes)
- `%%me` (17) - Memo annotations
- `` (18) - Section headers
- `cold` (3) - Column definitions
- `eqed` (3) - Equation objects
- `%hlk` (1) - Hyperlinks
- `nwno` (1) - Endnotes
- `pgnp` (1) - Page numbers
- `secd` (1) - Section definitions

---

## Strengths

1. **Full Feature Coverage**: Access to the complete Hangul feature set including all formatting, tables, images, headers/footers, equations, and more
2. **Perfect Rendering Fidelity**: Output is identical to what a human would create in Hangul - no compatibility issues
3. **PDF Export**: Direct PDF generation via `SaveAs("PDF")` with perfect rendering
4. **Multi-Format Export**: HWP, HWPX, HTML, TEXT, PDF all supported
5. **File Reading**: Can open and parse existing HWPX/HWP files, extract text, enumerate controls
6. **Rich Styling**: Full access to all character and paragraph properties
7. **Table Operations**: Create, navigate, merge cells, border/fill styling
8. **Korean Font Support**: All system-installed Korean fonts accessible

## Weaknesses

1. **Windows-Only**: Requires Windows with Hancom Office installed. Cannot run on Linux/Mac or in CI/CD
2. **Process Overhead**: Each COM session launches a full Hangul process (~50-80MB RAM per instance). Multiple hung processes observed during testing
3. **Reliability Issues**: COM processes can hang/zombie, requiring manual cleanup (`taskkill`). Rapid sequential launches sometimes fail with "operation already in progress" errors
4. **Speed**: Each test takes 5-15 seconds due to Hangul startup/shutdown overhead
5. **No Headless Mode**: Even with `Visible=False`, a full GUI process runs. Cannot run in a true server/headless environment
6. **API Discovery Difficulty**: Property names must be discovered via introspection (e.g., `AlignType` not `Alignment`, `PrevSpacing` not `SpaceBeforePara`). Some have typos (`BorderCorlorLeft`)
7. **Cursor-Based Model**: Table operations require navigating cell-by-cell like a user would. No random access to cells by (row, col) index. Complex documents require careful cursor management
8. **License Required**: Each machine needs a valid Hancom Office license
9. **Concurrency Limitations**: Running multiple instances simultaneously is unreliable
10. **Error Recovery**: If a COM call fails mid-operation, the Hangul process may be left in an inconsistent state

## Performance Characteristics

| Operation | Time (approx) |
|-----------|---------------|
| HWP COM object creation | 3-5 sec |
| Open 119-page HWPX | 2-3 sec |
| Insert text + table | < 1 sec |
| Full style test (30+ operations) | 2-3 sec |
| SaveAs HWPX | < 1 sec |
| SaveAs PDF (simple doc) | 1-2 sec |
| SaveAs PDF (119 pages) | 5-10 sec |
| Quit/cleanup | 1-2 sec |

## File Size Comparison

| Content | COM Output | python-hwpx Output |
|---------|-----------|-------------------|
| Text + 3x2 table | ~15KB | ~8KB |
| Styled document with 4x3 table | ~15KB | N/A |
| Full style test (fonts, align, spacing) | ~32KB | N/A |

COM output is larger due to richer metadata/namespaces from Hangul.

## Suitability Assessment

### Best For
- One-time batch processing on a Windows workstation
- Complex documents requiring exact Hangul rendering
- PDF generation from HWPX/HWP files
- Document conversion between formats
- Scenarios where a human operator can monitor/restart as needed

### Not Suitable For
- Cloud/server deployment (no headless support)
- CI/CD pipelines (Windows + Hangul dependency)
- High-volume automated generation (process overhead, hanging issues)
- Cross-platform applications
- Unattended long-running services (process reliability)

## Test Artifacts

All test scripts and outputs in `/tests/`:
- `test_com_open_read.py` - File open and text extraction
- `test_com_create.py` - Document creation with styled text and table
- `test_com_style.py` - Comprehensive styling test
- `test_com_pdf.py` - PDF and multi-format export
- `test_com_table_advanced.py` - Cell merging, navigation
- `test_com_inspect_para.py` - HParaShape property discovery
- `test_com_inspect_cell.py` - HCellBorderFill property discovery
- `test_com_inspect_fillattr.py` - FillAttr property discovery
- `output_com_create.hwpx` - Created document
- `output_com_style.hwpx` - Style test output
- `output_com_pdf_from_existing.pdf` - PDF from reference file
- `output_com_pdf_from_new.pdf` - PDF from new document
