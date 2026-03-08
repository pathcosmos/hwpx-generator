#!/usr/bin/env python3
"""
Binary-level comparison of section0.xml between original template and filled output HWPX files.
"""

import zipfile
import tempfile
import os
import sys
import hashlib
from lxml import etree
from collections import defaultdict

ORIGINAL_HWPX = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx"
FILLED_HWPX   = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx"
SECTION_PATH  = "Contents/section0.xml"

def hex_dump(data: bytes, width: int = 16) -> str:
    lines = []
    for i in range(0, len(data), width):
        chunk = data[i:i+width]
        hex_part = " ".join(f"{b:02x}" for b in chunk)
        ascii_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"  {i:04x}: {hex_part:<{width*3}} |{ascii_part}|")
    return "\n".join(lines)

def get_all_text(elem) -> str:
    """Get all text content recursively."""
    parts = []
    if elem.text:
        parts.append(elem.text)
    for child in elem:
        parts.append(get_all_text(child))
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)

def element_path(elem) -> str:
    """Build an XPath-like path for an element."""
    path_parts = []
    node = elem
    while node is not None and node.tag != "ROOT":
        tag = node.tag
        # Strip namespace
        if "{" in tag:
            tag = tag.split("}", 1)[1]
        parent = node.getparent()
        if parent is not None:
            siblings = [c for c in parent if c.tag == node.tag]
            if len(siblings) > 1:
                idx = siblings.index(node)
                tag = f"{tag}[{idx}]"
        path_parts.append(tag)
        node = parent
    return "/" + "/".join(reversed(path_parts))

def collect_elements(root):
    """Collect all elements with their text content."""
    elements = []
    for elem in root.iter():
        tag = elem.tag
        if "{" in tag:
            tag = tag.split("}", 1)[1]
        text = (elem.text or "").strip()
        elements.append((tag, text, elem))
    return elements

def list_zip_entries(zf: zipfile.ZipFile):
    """Return dict of {name: (size, crc, compress_type)}."""
    return {
        info.filename: (info.file_size, info.CRC, info.compress_type, info.compress_size)
        for info in zf.infolist()
    }

def main():
    print("=" * 80)
    print("HWPX section0.xml Binary + Structural Comparison")
    print("=" * 80)
    print(f"Original : {ORIGINAL_HWPX}")
    print(f"Filled   : {FILLED_HWPX}")
    print()

    # ─── Step 1: Extract section0.xml raw bytes ───────────────────────────────
    print("─" * 80)
    print("STEP 1: Extract section0.xml raw bytes")
    print("─" * 80)

    with zipfile.ZipFile(ORIGINAL_HWPX, "r") as zf_orig:
        orig_bytes = zf_orig.read(SECTION_PATH)

    with zipfile.ZipFile(FILLED_HWPX, "r") as zf_fill:
        fill_bytes = zf_fill.read(SECTION_PATH)

    print(f"Original  size: {len(orig_bytes):,} bytes")
    print(f"Filled    size: {len(fill_bytes):,} bytes")
    print(f"Size delta    : {len(fill_bytes) - len(orig_bytes):+,} bytes")
    print(f"Original  MD5 : {hashlib.md5(orig_bytes).hexdigest()}")
    print(f"Filled    MD5 : {hashlib.md5(fill_bytes).hexdigest()}")
    print()

    # ─── Step 2: XML Declaration / first 100 bytes hex dump ──────────────────
    print("─" * 80)
    print("STEP 2: First 100 bytes — XML declaration & BOM check")
    print("─" * 80)

    INSPECT = 100
    print(f"\n[ORIGINAL] first {INSPECT} bytes:")
    print(hex_dump(orig_bytes[:INSPECT]))
    print(f"\n[FILLED]   first {INSPECT} bytes:")
    print(hex_dump(fill_bytes[:INSPECT]))

    # BOM check
    for label, data in [("Original", orig_bytes), ("Filled", fill_bytes)]:
        if data[:3] == b"\xef\xbb\xbf":
            print(f"  !! {label}: UTF-8 BOM DETECTED")
        elif data[:2] in (b"\xff\xfe", b"\xfe\xff"):
            print(f"  !! {label}: UTF-16 BOM DETECTED")
        else:
            print(f"  OK {label}: No BOM")

    # XML declaration comparison
    orig_decl_end = orig_bytes.find(b"?>") + 2
    fill_decl_end = fill_bytes.find(b"?>") + 2
    orig_decl = orig_bytes[:orig_decl_end]
    fill_decl = fill_bytes[:fill_decl_end]

    print(f"\nOriginal XML decl : {orig_decl!r}")
    print(f"Filled   XML decl : {fill_decl!r}")
    if orig_decl == fill_decl:
        print("  => XML declarations are IDENTICAL")
    else:
        print("  => XML declarations DIFFER!")
        # Show byte-by-byte diff
        for i, (a, b) in enumerate(zip(orig_decl, fill_decl)):
            if a != b:
                print(f"     byte[{i}]: orig={a:02x} ({chr(a) if 32<=a<127 else '?'})  fill={b:02x} ({chr(b) if 32<=b<127 else '?'})")

    # Check byte immediately after declaration
    print(f"\nByte after decl (orig): {orig_bytes[orig_decl_end:orig_decl_end+4]!r}")
    print(f"Byte after decl (fill): {fill_bytes[fill_decl_end:fill_decl_end+4]!r}")
    print()

    # ─── Step 3: Parse with lxml and structural diff ─────────────────────────
    print("─" * 80)
    print("STEP 3: lxml structural diff")
    print("─" * 80)

    orig_root = etree.fromstring(orig_bytes)
    fill_root = etree.fromstring(fill_bytes)

    orig_all = list(orig_root.iter())
    fill_all = list(fill_root.iter())

    print(f"\nTotal elements (original) : {len(orig_all):,}")
    print(f"Total elements (filled)   : {len(fill_all):,}")
    print(f"Element delta             : {len(fill_all) - len(orig_all):+,}")
    print()

    # ─── Count elements by tag ────────────────────────────────────────────────
    def count_by_tag(root):
        counts = defaultdict(int)
        for e in root.iter():
            tag = e.tag
            if "{" in tag:
                tag = tag.split("}", 1)[1]
            counts[tag] += 1
        return counts

    orig_counts = count_by_tag(orig_root)
    fill_counts = count_by_tag(fill_root)

    all_tags = sorted(set(list(orig_counts.keys()) + list(fill_counts.keys())))
    changed_tags = [(t, orig_counts[t], fill_counts[t]) for t in all_tags
                    if orig_counts[t] != fill_counts[t]]

    if changed_tags:
        print("Tags with changed element counts:")
        print(f"  {'Tag':<30} {'Original':>10} {'Filled':>10} {'Delta':>10}")
        print(f"  {'-'*30} {'-'*10} {'-'*10} {'-'*10}")
        for tag, oc, fc in changed_tags:
            print(f"  {tag:<30} {oc:>10} {fc:>10} {fc-oc:>+10}")
    else:
        print("  No tag-count differences found.")
    print()

    # ─── Step 3b: Find modified text in hp:t elements ────────────────────────
    # Compare via positional walk of hp:run/hp:t elements
    print("─" * 80)
    print("STEP 3b: Text content differences (hp:t elements)")
    print("─" * 80)

    HP_NS = "http://www.hancom.co.kr/hwpml/2012/paragraph"

    def collect_runs(root):
        """Collect all (path, text) for hp:t elements."""
        results = []
        for elem in root.iter():
            if elem.tag in (f"{{{HP_NS}}}t", "t"):
                local = elem.tag.split("}", 1)[1] if "{" in elem.tag else elem.tag
                if local == "t":
                    text = elem.text or ""
                    results.append((element_path(elem), text, elem))
        return results

    orig_t = collect_runs(orig_root)
    fill_t = collect_runs(fill_root)

    print(f"\nTotal hp:t elements (original) : {len(orig_t)}")
    print(f"Total hp:t elements (filled)   : {len(fill_t)}")
    print()

    # Build maps by index for positional comparison
    min_len = min(len(orig_t), len(fill_t))
    max_len = max(len(orig_t), len(fill_t))

    text_diffs = []
    for i in range(min_len):
        o_path, o_text, _ = orig_t[i]
        f_path, f_text, _ = fill_t[i]
        if o_text != f_text:
            text_diffs.append((i, o_path, o_text, f_text))

    if text_diffs:
        print(f"Modified hp:t elements (positional): {len(text_diffs)}")
        for idx, path, old_text, new_text in text_diffs:
            print(f"\n  [#{idx}] path snippet: ...{path[-80:]}")
            print(f"    OLD: {old_text!r:.200}")
            print(f"    NEW: {new_text!r:.200}")
    else:
        print("  No positional text differences in common range.")

    if len(orig_t) != len(fill_t):
        print(f"\n  Count mismatch: {len(orig_t)} orig vs {len(fill_t)} filled")
        if len(fill_t) > len(orig_t):
            print(f"  Extra hp:t elements in FILLED (indices {min_len}..{max_len-1}):")
            for i in range(min_len, len(fill_t)):
                _, f_text, _ = fill_t[i]
                print(f"    [#{i}] text={f_text!r:.200}")
        else:
            print(f"  Missing hp:t elements from FILLED (indices {min_len}..{max_len-1}):")
            for i in range(min_len, len(orig_t)):
                _, o_text, _ = orig_t[i]
                print(f"    [#{i}] text={o_text!r:.200}")

    # ─── Step 3c: Search for injected markers ────────────────────────────────
    print()
    print("─" * 80)
    print("STEP 3c: Injected content markers search")
    print("─" * 80)

    MARKERS = ["##SEC1_CONTENT##", "##SEC2_CONTENT##", "##SEC3_CONTENT##",
               "##SEC4_CONTENT##", "##SEC5_CONTENT##", "MARKER", "##"]

    for marker in MARKERS:
        in_orig = marker.encode() in orig_bytes
        in_fill = marker.encode() in fill_bytes
        if in_orig or in_fill:
            print(f"  '{marker}': orig={in_orig}, filled={in_fill}")

    # Also scan filled XML text for any ## patterns
    fill_text_content = fill_bytes.decode("utf-8", errors="replace")
    import re
    found_markers = set(re.findall(r"##[A-Z0-9_]+##", fill_text_content))
    if found_markers:
        print(f"\n  Markers found in filled output: {sorted(found_markers)}")
    else:
        print("\n  No ##MARKER## patterns found in filled output.")

    # ─── Step 4: Compare ALL other ZIP entries ────────────────────────────────
    print()
    print("─" * 80)
    print("STEP 4: Compare ALL ZIP entries (other than section0.xml)")
    print("─" * 80)

    with zipfile.ZipFile(ORIGINAL_HWPX, "r") as zf_orig:
        orig_entries = list_zip_entries(zf_orig)
        orig_names   = set(zf_orig.namelist())

    with zipfile.ZipFile(FILLED_HWPX, "r") as zf_fill:
        fill_entries = list_zip_entries(zf_fill)
        fill_names   = set(zf_fill.namelist())

    only_in_orig = orig_names - fill_names
    only_in_fill = fill_names - orig_names
    common       = orig_names & fill_names

    if only_in_orig:
        print(f"\n  Files ONLY in original ({len(only_in_orig)}):")
        for n in sorted(only_in_orig):
            print(f"    - {n}")

    if only_in_fill:
        print(f"\n  Files ONLY in filled ({len(only_in_fill)}):")
        for n in sorted(only_in_fill):
            print(f"    + {n}")

    print(f"\n  Comparing {len(common)} common files:")
    print(f"  {'File':<50} {'Status':<12} {'OrigSize':>10} {'FillSize':>10} {'Delta':>10}")
    print(f"  {'-'*50} {'-'*12} {'-'*10} {'-'*10} {'-'*10}")

    changed_files = []
    identical_files = []
    for name in sorted(common):
        o_size, o_crc, o_ctype, o_csize = orig_entries[name]
        f_size, f_crc, f_ctype, f_csize = fill_entries[name]
        if o_size != f_size or o_crc != f_crc:
            changed_files.append(name)
            status = "CHANGED"
        else:
            identical_files.append(name)
            status = "identical"
        delta = f_size - o_size
        print(f"  {name:<50} {status:<12} {o_size:>10,} {f_size:>10,} {delta:>+10,}")

    print(f"\n  Summary: {len(identical_files)} identical, {len(changed_files)} changed")
    if changed_files:
        print(f"  Changed files: {changed_files}")

    # Deep byte comparison for changed files (other than section0.xml)
    other_changed = [f for f in changed_files if f != SECTION_PATH]
    if other_changed:
        print(f"\n  UNEXPECTED changes in non-section files!")
        with zipfile.ZipFile(ORIGINAL_HWPX, "r") as zf_orig, \
             zipfile.ZipFile(FILLED_HWPX,   "r") as zf_fill:
            for fname in other_changed:
                o_data = zf_orig.read(fname)
                f_data = zf_fill.read(fname)
                print(f"\n  [{fname}]")
                print(f"    Original  ({len(o_data)} bytes): {o_data[:200]!r}")
                print(f"    Filled    ({len(f_data)} bytes): {f_data[:200]!r}")
    else:
        print("\n  OK: Only section0.xml was changed in the filled HWPX.")

    # ─── Step 5: ZIP structure / mimetype / ordering check ───────────────────
    print()
    print("─" * 80)
    print("STEP 5: ZIP structure — mimetype, ordering, compression")
    print("─" * 80)

    with zipfile.ZipFile(ORIGINAL_HWPX, "r") as zf_orig:
        orig_order = [i.filename for i in zf_orig.infolist()]
        orig_mime  = zf_orig.infolist()[0]

    with zipfile.ZipFile(FILLED_HWPX, "r") as zf_fill:
        fill_order = [i.filename for i in zf_fill.infolist()]
        fill_mime  = zf_fill.infolist()[0]

    print(f"\n  Original  first entry: '{orig_mime.filename}' compress_type={orig_mime.compress_type}")
    print(f"  Filled    first entry: '{fill_mime.filename}' compress_type={fill_mime.compress_type}")
    if orig_mime.filename == fill_mime.filename and orig_mime.compress_type == fill_mime.compress_type:
        print("  OK: mimetype is first entry and STORED (compress_type=0)")
    else:
        print("  WARNING: mimetype entry mismatch!")

    # Check entry ordering
    if orig_order == fill_order:
        print(f"\n  ZIP entry order: IDENTICAL ({len(orig_order)} entries)")
    else:
        print(f"\n  ZIP entry order DIFFERS!")
        print(f"  Original order ({len(orig_order)}): {orig_order[:10]}...")
        print(f"  Filled   order ({len(fill_order)}): {fill_order[:10]}...")
        # Find first mismatch
        for i, (a, b) in enumerate(zip(orig_order, fill_order)):
            if a != b:
                print(f"  First mismatch at index {i}: orig='{a}' fill='{b}'")
                break

    # ─── Step 6: Deep byte diff — find first divergence in section0.xml ──────
    print()
    print("─" * 80)
    print("STEP 6: First byte divergence in section0.xml")
    print("─" * 80)

    first_diff = None
    for i, (a, b) in enumerate(zip(orig_bytes, fill_bytes)):
        if a != b:
            first_diff = i
            break

    if first_diff is not None:
        print(f"\n  First byte difference at offset {first_diff} (0x{first_diff:x})")
        ctx_start = max(0, first_diff - 32)
        ctx_end   = min(len(orig_bytes), first_diff + 64)
        print(f"\n  Context around offset {first_diff}:")
        print(f"  [ORIGINAL bytes {ctx_start}..{ctx_end}]")
        print(hex_dump(orig_bytes[ctx_start:ctx_end]))
        print(f"  [FILLED   bytes {ctx_start}..{ctx_end}]")
        print(hex_dump(fill_bytes[ctx_start:ctx_end]))
        # Decode as UTF-8 for context
        try:
            orig_ctx_str = orig_bytes[ctx_start:ctx_end].decode("utf-8", errors="replace")
            fill_ctx_str = fill_bytes[ctx_start:ctx_end].decode("utf-8", errors="replace")
            print(f"\n  Original text context: ...{orig_ctx_str!r}...")
            print(f"  Filled   text context: ...{fill_ctx_str!r}...")
        except Exception:
            pass
    else:
        if len(orig_bytes) == len(fill_bytes):
            print("  Files are byte-for-byte IDENTICAL!")
        else:
            shorter = "original" if len(orig_bytes) < len(fill_bytes) else "filled"
            print(f"  No diff in shared prefix; {shorter} is shorter by "
                  f"{abs(len(orig_bytes)-len(fill_bytes))} bytes")

    # ─── Step 7: Attribute changes on modified cells ─────────────────────────
    print()
    print("─" * 80)
    print("STEP 7: Attribute changes on hp:run elements")
    print("─" * 80)

    def collect_runs_attrs(root):
        HP_NS2 = "http://www.hancom.co.kr/hwpml/2012/paragraph"
        results = []
        for elem in root.iter(f"{{{HP_NS2}}}run"):
            attrs = dict(elem.attrib)
            t_elem = elem.find(f"{{{HP_NS2}}}t")
            text = (t_elem.text or "") if t_elem is not None else ""
            results.append((attrs, text))
        return results

    orig_runs = collect_runs_attrs(orig_root)
    fill_runs = collect_runs_attrs(fill_root)

    print(f"\n  Total hp:run elements: orig={len(orig_runs)}, filled={len(fill_runs)}")
    attr_diffs = 0
    for i, (o, f) in enumerate(zip(orig_runs, fill_runs)):
        o_attrs, o_text = o
        f_attrs, f_text = f
        if o_attrs != f_attrs:
            attr_diffs += 1
            if attr_diffs <= 10:
                print(f"\n  Run #{i} attr diff:")
                print(f"    OLD attrs: {o_attrs}")
                print(f"    NEW attrs: {f_attrs}")
                print(f"    text: {o_text!r} -> {f_text!r}")
    if attr_diffs == 0:
        print("  OK: No attribute differences on hp:run elements in common range.")
    else:
        print(f"\n  Total hp:run attribute differences: {attr_diffs}")

    print()
    print("=" * 80)
    print("Comparison complete.")
    print("=" * 80)

if __name__ == "__main__":
    main()
