# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This project generates HWPX files — the XML-based document format used by Hancom Office Hangul (한컴오피스 한글), Korea's dominant word processor. The goal is to programmatically create `.hwpx` documents.

## HWPX Format Reference

The `ref/` directory contains sample files (`test_01.hwpx`, `test_01.hwp`, `test_01.pdf`) for reverse-engineering the format. The HWPX file in the reference is a Korean government business proposal document (2026년 클라우드 종합솔루션 지원사업 사업계획서).

A `.hwpx` file is a ZIP archive (`application/hwp+zip` mimetype) with this structure:

```
mimetype                    # "application/hwp+zip" (must be first, uncompressed)
version.xml                 # HWPML version info (currently 1.5)
META-INF/
  container.xml             # OPF rootfile references
  container.rdf             # RDF metadata
  manifest.xml              # ODF manifest
Contents/
  content.hpf               # OPF package manifest (lists all items, spine order)
  header.xml                # Document head: fonts, styles (charPr, paraPr), border fills
  section0.xml              # Document body: paragraphs, tables, images, layout
BinData/                    # Embedded images (PNG, BMP)
Preview/
  PrvText.txt               # Plain-text preview
  PrvImage.png              # Thumbnail preview
settings.xml                # Application settings (print, zoom, caret position)
```

### Key XML Namespaces

All HWPML XML uses namespaces under `http://www.hancom.co.kr/hwpml/2011/`:
- `hh:` (head) — fonts, char/para properties, styles, border fills
- `hp:` (paragraph) — paragraphs (`<hp:p>`), runs (`<hp:run>`), text (`<hp:t>`), tables (`<hp:tbl>`), cells (`<hp:tc>`)
- `hs:` (section) — section root (`<hs:sec>`), page properties
- `hc:` (core) — core types
- `ha:` (app) — application settings
- `opf:` — OPF package format (content.hpf)

### Important Format Details

- Units are in HWPUNIT (1/7200 inch). Standard A4: width=59528, height=84188.
- Styles use numeric IDs referenced via attributes like `charPrIDRef`, `paraPrIDRef`, `borderFillIDRef`.
- Tables use `<hp:tbl>` with `<hp:tr>` rows and `<hp:tc>` cells; cells contain `<hp:subList>` with paragraphs.
- Cell spanning via `<hp:cellSpan colSpan="N" rowSpan="N"/>`.
- Images are referenced by `binaryItemIDRef` matching IDs in `content.hpf`.

## Development Notes

- To inspect the reference HWPX file: `python3 -c "import zipfile; z = zipfile.ZipFile('ref/test_01.hwpx'); print(z.namelist())"`
- The project is in early/pre-development stage. No build system, tests, or source code exist yet.
