#!/usr/bin/env python3
"""
Cross-reference audit for filled HWPX file.
Checks section0.xml references against header.xml definitions.
"""

import zipfile
import sys
from lxml import etree
from collections import defaultdict

HWPX_PATH = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx"

NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hpf': 'http://www.hancom.co.kr/schema/2011/hpf',
    'opf': 'http://www.idpf.org/2007/opf/',
}

def parse_xml_from_zip(zf, member_name):
    """Parse an XML file from a ZIP archive."""
    with zf.open(member_name) as f:
        data = f.read()
    return etree.fromstring(data)

def collect_ids_from_header(header_root, tag_local, id_attr='id'):
    """Collect all IDs for a given tag in header.xml."""
    ids = set()
    # Search with namespace prefix hh
    hh_ns = NAMESPACES['hh']
    for elem in header_root.iter(f'{{{hh_ns}}}{tag_local}'):
        val = elem.get(id_attr)
        if val is not None:
            ids.add(val)
    return ids

def collect_ids_from_section(section_root, tag_ns_key, tag_local, attr_name):
    """Collect all unique attribute values for a given tag in section0.xml."""
    values = set()
    ns = NAMESPACES[tag_ns_key]
    for elem in section_root.iter(f'{{{ns}}}{tag_local}'):
        val = elem.get(attr_name)
        if val is not None:
            values.add(val)
    return values

def audit_references(zf):
    print("=" * 70)
    print("HWPX CROSS-REFERENCE AUDIT")
    print("=" * 70)
    print(f"File: {HWPX_PATH}\n")

    # -----------------------------------------------------------------------
    # 1. List all ZIP members
    # -----------------------------------------------------------------------
    all_members = zf.namelist()
    print(f"[ZIP] Total members: {len(all_members)}")
    for m in sorted(all_members):
        info = zf.getinfo(m)
        print(f"  {m}  ({info.file_size:,} bytes)")
    print()

    # -----------------------------------------------------------------------
    # 2. Load section0.xml and header.xml
    # -----------------------------------------------------------------------
    section_name = None
    header_name = None
    content_hpf_name = None

    for m in all_members:
        if 'section0.xml' in m or 'Section0.xml' in m:
            section_name = m
        if 'header.xml' in m or 'Header.xml' in m:
            header_name = m
        if 'content.hpf' in m or 'Content.hpf' in m:
            content_hpf_name = m

    print(f"[Files] section0.xml -> {section_name}")
    print(f"[Files] header.xml   -> {header_name}")
    print(f"[Files] content.hpf  -> {content_hpf_name}")
    print()

    if not section_name:
        print("ERROR: section0.xml not found in ZIP!")
        return
    if not header_name:
        print("ERROR: header.xml not found in ZIP!")
        return

    section_root = parse_xml_from_zip(zf, section_name)
    header_root  = parse_xml_from_zip(zf, header_name)

    # -----------------------------------------------------------------------
    # 3. Extract references from section0.xml
    # -----------------------------------------------------------------------
    # charPrIDRef from hp:run
    char_pr_refs = collect_ids_from_section(section_root, 'hp', 'run', 'charPrIDRef')
    # paraPrIDRef from hp:p
    para_pr_refs = collect_ids_from_section(section_root, 'hp', 'p', 'paraPrIDRef')
    # styleIDRef from hp:p
    style_id_refs = collect_ids_from_section(section_root, 'hp', 'p', 'styleIDRef')
    # borderFillIDRef from hp:tc (table cells)
    border_fill_refs = collect_ids_from_section(section_root, 'hp', 'tc', 'borderFillIDRef')

    print(f"[Section0] Unique charPrIDRef values:     {len(char_pr_refs)}")
    print(f"[Section0] Unique paraPrIDRef values:     {len(para_pr_refs)}")
    print(f"[Section0] Unique styleIDRef values:      {len(style_id_refs)}")
    print(f"[Section0] Unique borderFillIDRef values: {len(border_fill_refs)}")
    print()

    # -----------------------------------------------------------------------
    # 4. Extract definitions from header.xml
    # -----------------------------------------------------------------------
    defined_char_pr    = collect_ids_from_header(header_root, 'charPr')
    defined_para_pr    = collect_ids_from_header(header_root, 'paraPr')
    defined_style      = collect_ids_from_header(header_root, 'style')
    defined_border_fill = collect_ids_from_header(header_root, 'borderFill')

    print(f"[Header]   Defined hh:charPr IDs:      {len(defined_char_pr)}")
    print(f"[Header]   Defined hh:paraPr IDs:      {len(defined_para_pr)}")
    print(f"[Header]   Defined hh:style IDs:       {len(defined_style)}")
    print(f"[Header]   Defined hh:borderFill IDs:  {len(defined_border_fill)}")
    print()

    # -----------------------------------------------------------------------
    # 5. Cross-check: find missing refs
    # -----------------------------------------------------------------------
    issues = []

    def check_refs(label, refs, defined):
        missing = refs - defined
        ok = refs - missing
        print(f"[CHECK] {label}")
        print(f"        Used: {sorted(refs)[:20]}{'...' if len(refs)>20 else ''}")
        print(f"        Defined (sample): {sorted(defined)[:20]}{'...' if len(defined)>20 else ''}")
        if missing:
            print(f"        *** MISSING ({len(missing)} refs): {sorted(missing)}")
            issues.append((label, sorted(missing)))
        else:
            print(f"        OK - all {len(refs)} refs resolved")
        print()
        return missing

    missing_char   = check_refs("charPrIDRef   (hp:run  -> hh:charPr)",   char_pr_refs,    defined_char_pr)
    missing_para   = check_refs("paraPrIDRef   (hp:p    -> hh:paraPr)",   para_pr_refs,    defined_para_pr)
    missing_style  = check_refs("styleIDRef    (hp:p    -> hh:style)",    style_id_refs,   defined_style)
    missing_border = check_refs("borderFillIDRef(hp:tc  -> hh:borderFill)", border_fill_refs, defined_border_fill)

    # -----------------------------------------------------------------------
    # 6. Summary of ALL IDs in header for completeness
    # -----------------------------------------------------------------------
    print("[Header Summary]")
    print(f"  charPr IDs range:     {sorted(defined_char_pr)[:5]} ... {sorted(defined_char_pr)[-5:] if len(defined_char_pr)>5 else ''}")
    print(f"  paraPr IDs range:     {sorted(defined_para_pr)[:5]} ... {sorted(defined_para_pr)[-5:] if len(defined_para_pr)>5 else ''}")
    print(f"  style IDs range:      {sorted(defined_style)[:5]} ... {sorted(defined_style)[-5:] if len(defined_style)>5 else ''}")
    print(f"  borderFill IDs range: {sorted(defined_border_fill)[:5]} ... {sorted(defined_border_fill)[-5:] if len(defined_border_fill)>5 else ''}")
    print()

    # -----------------------------------------------------------------------
    # 7. content.hpf manifest check
    # -----------------------------------------------------------------------
    print("[content.hpf Manifest Check]")
    if content_hpf_name:
        hpf_root = parse_xml_from_zip(zf, content_hpf_name)
        # Find all manifest items
        hpf_ns = NAMESPACES['hpf']
        opf_ns = NAMESPACES['opf']

        manifest_items = []
        # Try hpf namespace
        for item in hpf_root.iter(f'{{{hpf_ns}}}item'):
            href = item.get('href')
            if href:
                manifest_items.append(href)
        # Try opf namespace
        for item in hpf_root.iter(f'{{{opf_ns}}}item'):
            href = item.get('href')
            if href:
                manifest_items.append(href)
        # Try without namespace
        for item in hpf_root.iter('item'):
            href = item.get('href')
            if href and href not in manifest_items:
                manifest_items.append(href)

        print(f"  Manifest items listed: {len(manifest_items)}")
        for item in sorted(manifest_items):
            print(f"    manifest: {item}")

        # Check each manifest item exists in zip
        manifest_missing = []
        for href in manifest_items:
            # Normalize path: content.hpf is at root or in a subdir
            # Try as-is and with common prefixes
            found = False
            for candidate in [href, href.lstrip('/'), f"OEBPS/{href}", f"Contents/{href}"]:
                if candidate in all_members:
                    found = True
                    break
            if not found:
                manifest_missing.append(href)
                print(f"    *** MISSING in ZIP: {href}")
            else:
                print(f"    OK: {href}")

        # Check ZIP members not in manifest (excluding content.hpf and mimetype)
        manifest_set = set(manifest_items)
        # Normalize manifest paths
        manifest_normalized = set()
        for m in manifest_items:
            manifest_normalized.add(m.lstrip('/'))

        unlisted = []
        for member in all_members:
            # Skip container/meta files
            if any(skip in member for skip in ['mimetype', 'content.hpf', 'Content.hpf',
                                                 'META-INF', 'settings.xml', 'Settings.xml',
                                                 'docInfo.xml', 'DocInfo.xml']):
                continue
            base = member.lstrip('/')
            if base not in manifest_normalized and member not in manifest_normalized:
                unlisted.append(member)

        if unlisted:
            print(f"\n  *** ZIP members NOT listed in manifest ({len(unlisted)}):")
            for u in unlisted:
                print(f"    {u}")
        else:
            print(f"\n  All significant ZIP members are in manifest.")

        if manifest_missing:
            issues.append(("content.hpf manifest items missing from ZIP", manifest_missing))
    else:
        print("  WARNING: content.hpf not found in ZIP")
        issues.append(("content.hpf", ["file not found in ZIP"]))
    print()

    # -----------------------------------------------------------------------
    # 8. Final summary
    # -----------------------------------------------------------------------
    print("=" * 70)
    print("FINAL AUDIT SUMMARY")
    print("=" * 70)
    if not issues:
        print("ALL CROSS-REFERENCES VALID - No issues found.")
    else:
        print(f"ISSUES FOUND: {len(issues)} categories with problems\n")
        for label, items in issues:
            print(f"  [{label}]")
            for item in items:
                print(f"    - {item}")
    print()

    # -----------------------------------------------------------------------
    # 9. Raw value dumps for debugging
    # -----------------------------------------------------------------------
    print("[Raw Values - charPrIDRef used in section0]")
    for v in sorted(char_pr_refs):
        in_header = "OK" if v in defined_char_pr else "MISSING"
        print(f"  {v:>6}  {in_header}")
    print()

    print("[Raw Values - paraPrIDRef used in section0]")
    for v in sorted(para_pr_refs):
        in_header = "OK" if v in defined_para_pr else "MISSING"
        print(f"  {v:>6}  {in_header}")
    print()

    print("[Raw Values - styleIDRef used in section0]")
    for v in sorted(style_id_refs):
        in_header = "OK" if v in defined_style else "MISSING"
        print(f"  {v:>6}  {in_header}")
    print()

    print("[Raw Values - borderFillIDRef used in section0]")
    for v in sorted(border_fill_refs):
        in_header = "OK" if v in defined_border_fill else "MISSING"
        print(f"  {v:>6}  {in_header}")
    print()

def main():
    try:
        with zipfile.ZipFile(HWPX_PATH, 'r') as zf:
            audit_references(zf)
    except FileNotFoundError:
        print(f"ERROR: File not found: {HWPX_PATH}")
        sys.exit(1)
    except zipfile.BadZipFile as e:
        print(f"ERROR: Bad ZIP file: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
