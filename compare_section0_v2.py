#!/usr/bin/env python3
"""
Thorough binary-level comparison of section0.xml between original template and filled output.
v2: Uses namespace-agnostic search via localname, covers hp:t, hp:run, cell content, markers.
"""

import zipfile
import hashlib
import re
from lxml import etree
from collections import defaultdict

ORIGINAL_HWPX = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx"
FILLED_HWPX   = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx"
SECTION_PATH  = "Contents/section0.xml"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def local(tag: str) -> str:
    """Strip namespace."""
    return tag.split("}", 1)[1] if "{" in tag else tag

def hex_dump(data: bytes, width: int = 16, offset: int = 0) -> str:
    lines = []
    for i in range(0, len(data), width):
        chunk = data[i:i+width]
        hex_part = " ".join(f"{b:02x}" for b in chunk)
        asc_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"  {offset+i:06x}: {hex_part:<{width*3}} |{asc_part}|")
    return "\n".join(lines)

def iter_local(root, localname):
    """Iterate elements matching a local tag name, regardless of namespace."""
    for elem in root.iter():
        if local(elem.tag) == localname:
            yield elem

def get_cell_text(cell_elem) -> str:
    """Collect all hp:t text inside a cell."""
    parts = []
    for t in iter_local(cell_elem, "t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)

def cell_id(cell_elem) -> str:
    """Make a readable id from cell attributes."""
    a = cell_elem.attrib
    row = a.get("rowAddr", a.get("row", "?"))
    col = a.get("colAddr", a.get("col", "?"))
    return f"row={row},col={col}"

def collect_cells(root):
    """Return list of (cell_id_str, text, elem) for each hp:tc element."""
    results = []
    for elem in iter_local(root, "tc"):
        results.append((cell_id(elem), get_cell_text(elem), elem))
    return results

def collect_t_elements(root):
    """Return list of (index, text) for all hp:t elements."""
    return [(i, (e.text or ""), e)
            for i, e in enumerate(iter_local(root, "t"))]

def collect_run_elements(root):
    """Return list of (index, attrs_dict, text) for all hp:run elements."""
    results = []
    for i, e in enumerate(iter_local(root, "run")):
        t_child = next(iter_local(e, "t"), None)
        text = (t_child.text or "") if t_child is not None else ""
        results.append((i, dict(e.attrib), text))
    return results

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    DIVIDER = "─" * 80

    print("=" * 80)
    print("HWPX section0.xml Thorough Binary + Structural Comparison  (v2)")
    print("=" * 80)
    print(f"Original : {ORIGINAL_HWPX}")
    print(f"Filled   : {FILLED_HWPX}")
    print()

    # ── 1. Extract raw bytes ──────────────────────────────────────────────────
    print(DIVIDER)
    print("STEP 1: Raw bytes")
    print(DIVIDER)

    with zipfile.ZipFile(ORIGINAL_HWPX, "r") as z:
        orig_bytes = z.read(SECTION_PATH)
    with zipfile.ZipFile(FILLED_HWPX, "r") as z:
        fill_bytes = z.read(SECTION_PATH)

    print(f"  Original  : {len(orig_bytes):,} bytes  MD5={hashlib.md5(orig_bytes).hexdigest()}")
    print(f"  Filled    : {len(fill_bytes):,} bytes  MD5={hashlib.md5(fill_bytes).hexdigest()}")
    print(f"  Size delta: {len(fill_bytes) - len(orig_bytes):+,} bytes")
    print()

    # ── 2. XML declaration + BOM ──────────────────────────────────────────────
    print(DIVIDER)
    print("STEP 2: XML declaration & BOM (first 100 bytes)")
    print(DIVIDER)

    for label, data in [("Original", orig_bytes), ("Filled", fill_bytes)]:
        bom = "UTF-8 BOM" if data[:3] == b"\xef\xbb\xbf" else \
              "UTF-16-LE BOM" if data[:2] == b"\xff\xfe" else \
              "UTF-16-BE BOM" if data[:2] == b"\xfe\xff" else "None"
        print(f"\n  [{label}]  BOM={bom}")
        print(hex_dump(data[:100]))

    orig_decl = orig_bytes[:orig_bytes.find(b"?>") + 2]
    fill_decl = fill_bytes[:fill_bytes.find(b"?>") + 2]
    print(f"\n  Original decl: {orig_decl!r}")
    print(f"  Filled   decl: {fill_decl!r}")
    status = "IDENTICAL" if orig_decl == fill_decl else "DIFFER"
    print(f"  => {status}")
    if status == "DIFFER":
        for i, (a, b) in enumerate(zip(orig_decl, fill_decl)):
            if a != b:
                print(f"     byte[{i}]: orig=0x{a:02x} fill=0x{b:02x}")

    after_orig = orig_bytes[len(orig_decl):len(orig_decl)+8]
    after_fill = fill_bytes[len(fill_decl):len(fill_decl)+8]
    print(f"\n  Bytes immediately after decl — orig: {after_orig!r}  fill: {after_fill!r}")
    print()

    # ── 3. lxml parse — element counts ───────────────────────────────────────
    print(DIVIDER)
    print("STEP 3: Element counts (tag-level)")
    print(DIVIDER)

    orig_root = etree.fromstring(orig_bytes)
    fill_root = etree.fromstring(fill_bytes)

    orig_all = list(orig_root.iter())
    fill_all = list(fill_root.iter())
    print(f"\n  Total elements: orig={len(orig_all):,}  fill={len(fill_all):,}  delta={len(fill_all)-len(orig_all):+,}")

    orig_counts = defaultdict(int)
    fill_counts = defaultdict(int)
    for e in orig_all:
        orig_counts[local(e.tag)] += 1
    for e in fill_all:
        fill_counts[local(e.tag)] += 1

    all_tags = sorted(set(list(orig_counts) + list(fill_counts)))
    diff_tags = [(t, orig_counts[t], fill_counts[t]) for t in all_tags
                 if orig_counts[t] != fill_counts[t]]

    if diff_tags:
        print(f"\n  Tags with changed counts ({len(diff_tags)}):")
        print(f"  {'Tag':<28} {'Orig':>8} {'Fill':>8} {'Delta':>8}")
        print(f"  {'-'*28} {'-'*8} {'-'*8} {'-'*8}")
        for t, o, f in diff_tags:
            print(f"  {t:<28} {o:>8} {f:>8} {f-o:>+8}")
    else:
        print("  No tag-count differences.")
    print()

    # ── 4. hp:t text differences (positional) ────────────────────────────────
    print(DIVIDER)
    print("STEP 4: hp:t text content — positional comparison")
    print(DIVIDER)

    orig_t_list = collect_t_elements(orig_root)
    fill_t_list = collect_t_elements(fill_root)
    print(f"\n  hp:t count: orig={len(orig_t_list)}  fill={len(fill_t_list)}")

    min_t = min(len(orig_t_list), len(fill_t_list))
    t_diffs = [(i, ot, ft) for (i, ot, _), (_, ft, _2) in
               [(orig_t_list[i], fill_t_list[i]) for i in range(min_t)]
               if ot != ft]

    if t_diffs:
        print(f"\n  hp:t elements with different text in common range: {len(t_diffs)}")
        for idx, old_text, new_text in t_diffs:
            print(f"\n    [#{idx}]")
            print(f"      OLD: {old_text!r}")
            print(f"      NEW: {new_text!r}")
    else:
        print("  No text differences in common positional range.")

    if len(orig_t_list) != len(fill_t_list):
        print(f"\n  Count mismatch: {abs(len(orig_t_list)-len(fill_t_list))} extra elements in "
              f"{'filled' if len(fill_t_list)>len(orig_t_list) else 'original'}")
        extra_list = fill_t_list[min_t:] if len(fill_t_list) > min_t else orig_t_list[min_t:]
        label = "FILLED (extra)" if len(fill_t_list) > len(orig_t_list) else "ORIGINAL (missing from filled)"
        print(f"  Extra elements ({label}):")
        for idx, text, _ in extra_list[:50]:
            print(f"    [#{idx}] {text!r}")
        if len(extra_list) > 50:
            print(f"    ... and {len(extra_list)-50} more")
    print()

    # ── 5. Cell-level diff (hp:tc) ────────────────────────────────────────────
    print(DIVIDER)
    print("STEP 5: Cell-level diff (hp:tc) — all modified cells")
    print(DIVIDER)

    orig_cells = collect_cells(orig_root)
    fill_cells = collect_cells(fill_root)
    print(f"\n  hp:tc count: orig={len(orig_cells)}  fill={len(fill_cells)}")

    min_cells = min(len(orig_cells), len(fill_cells))
    cell_diffs = []
    for i in range(min_cells):
        o_id, o_text, _ = orig_cells[i]
        f_id, f_text, _ = fill_cells[i]
        if o_text != f_text:
            cell_diffs.append((i, o_id, o_text, f_text))

    if cell_diffs:
        print(f"\n  Modified cells: {len(cell_diffs)}")
        for idx, cid, old_t, new_t in cell_diffs:
            print(f"\n    Cell #{idx} ({cid})")
            # Shorten long text for display
            old_display = old_t[:300] + "..." if len(old_t) > 300 else old_t
            new_display = new_t[:300] + "..." if len(new_t) > 300 else new_t
            print(f"      OLD ({len(old_t)} chars): {old_display!r}")
            print(f"      NEW ({len(new_t)} chars): {new_display!r}")
    else:
        print("  No cell text differences in common range.")

    if len(orig_cells) != len(fill_cells):
        print(f"\n  Cell count mismatch: orig={len(orig_cells)} fill={len(fill_cells)}")
    print()

    # ── 6. hp:run attribute changes ───────────────────────────────────────────
    print(DIVIDER)
    print("STEP 6: hp:run attribute changes (positional)")
    print(DIVIDER)

    orig_runs = collect_run_elements(orig_root)
    fill_runs = collect_run_elements(fill_root)
    print(f"\n  hp:run count: orig={len(orig_runs)}  fill={len(fill_runs)}")

    min_runs = min(len(orig_runs), len(fill_runs))
    run_attr_diffs = []
    run_text_diffs = []
    for i in range(min_runs):
        o_idx, o_attrs, o_text = orig_runs[i]
        f_idx, f_attrs, f_text = fill_runs[i]
        if o_attrs != f_attrs:
            run_attr_diffs.append((i, o_attrs, f_attrs, o_text, f_text))
        if o_text != f_text:
            run_text_diffs.append((i, o_text, f_text))

    if run_attr_diffs:
        print(f"\n  hp:run elements with changed attributes: {len(run_attr_diffs)}")
        for i, oa, fa, ot, ft in run_attr_diffs[:20]:
            print(f"\n    Run #{i}  text: {ot!r} -> {ft!r}")
            # Show only changed attrs
            all_attr_keys = set(list(oa.keys()) + list(fa.keys()))
            for k in sorted(all_attr_keys):
                ov = oa.get(k, "<missing>")
                fv = fa.get(k, "<missing>")
                if ov != fv:
                    print(f"      attr '{k}': {ov!r} -> {fv!r}")
    else:
        print("  No hp:run attribute differences in common range.")

    if run_text_diffs:
        print(f"\n  hp:run text differences (positional): {len(run_text_diffs)}")
        for i, ot, ft in run_text_diffs[:20]:
            print(f"    Run #{i}: {ot!r} -> {ft!r}")
    else:
        print("  No hp:run text differences in common positional range.")

    if len(orig_runs) != len(fill_runs):
        delta = len(fill_runs) - len(orig_runs)
        print(f"\n  Run count delta: {delta:+d}")
        if delta > 0:
            print(f"  Extra runs in filled (last {min(delta, 20)}):")
            for i_idx, attrs, text in fill_runs[min_runs:min_runs+20]:
                print(f"    #{i_idx}: attrs={attrs}  text={text!r}")
        else:
            print(f"  Missing runs in filled (first {min(-delta, 20)}):")
            for i_idx, attrs, text in orig_runs[min_runs:min_runs+20]:
                print(f"    #{i_idx}: attrs={attrs}  text={text!r}")
    print()

    # ── 7. Marker injection check ─────────────────────────────────────────────
    print(DIVIDER)
    print("STEP 7: Injected content markers")
    print(DIVIDER)

    fill_str = fill_bytes.decode("utf-8", errors="replace")
    orig_str = orig_bytes.decode("utf-8", errors="replace")

    all_markers = re.findall(r"##[A-Z0-9_]+##", fill_str)
    unique_markers = sorted(set(all_markers))

    print(f"\n  Markers found in FILLED: {unique_markers}")
    print(f"  Marker occurrences: {len(all_markers)}")
    for m in unique_markers:
        in_orig = m in orig_str
        count_in_fill = fill_str.count(m)
        # Find byte offset of first occurrence in filled
        m_bytes = m.encode("utf-8")
        offset = fill_bytes.find(m_bytes)
        print(f"\n  '{m}':")
        print(f"    In original : {in_orig}")
        print(f"    Count in fill: {count_in_fill}")
        print(f"    First byte offset: {offset} (0x{offset:x})")
        # Show 80 bytes of context around first occurrence
        ctx_start = max(0, offset - 40)
        ctx_end   = min(len(fill_bytes), offset + len(m_bytes) + 40)
        context_str = fill_bytes[ctx_start:ctx_end].decode("utf-8", errors="replace")
        print(f"    Context: ...{context_str!r}...")
    print()

    # ── 8. First byte divergence (raw) ────────────────────────────────────────
    print(DIVIDER)
    print("STEP 8: First raw byte divergence in section0.xml")
    print(DIVIDER)

    first_diff = None
    for i, (a, b) in enumerate(zip(orig_bytes, fill_bytes)):
        if a != b:
            first_diff = i
            break

    if first_diff is not None:
        print(f"\n  First diff at byte offset {first_diff} (0x{first_diff:x})")
        ctx_s = max(0, first_diff - 40)
        ctx_e = min(min(len(orig_bytes), len(fill_bytes)), first_diff + 80)
        print(f"\n  [ORIGINAL  bytes {ctx_s:#x}..{ctx_e:#x}]")
        print(hex_dump(orig_bytes[ctx_s:ctx_e], offset=ctx_s))
        print(f"\n  [FILLED    bytes {ctx_s:#x}..{ctx_e:#x}]")
        print(hex_dump(fill_bytes[ctx_s:ctx_e], offset=ctx_s))
        try:
            print(f"\n  Original text: {orig_bytes[ctx_s:ctx_e].decode('utf-8','replace')!r}")
            print(f"  Filled   text: {fill_bytes[ctx_s:ctx_e].decode('utf-8','replace')!r}")
        except Exception:
            pass
    else:
        if len(orig_bytes) == len(fill_bytes):
            print("  Byte-for-byte IDENTICAL!")
        else:
            shorter = "original" if len(orig_bytes) < len(fill_bytes) else "filled"
            print(f"  Shared prefix identical; {shorter} is shorter")
    print()

    # ── 9. All other ZIP entries ──────────────────────────────────────────────
    print(DIVIDER)
    print("STEP 9: All ZIP entries comparison")
    print(DIVIDER)

    def zip_manifest(path):
        with zipfile.ZipFile(path, "r") as z:
            return {i.filename: (i.file_size, i.CRC, i.compress_type) for i in z.infolist()}, \
                   [i.filename for i in z.infolist()]

    orig_man, orig_order = zip_manifest(ORIGINAL_HWPX)
    fill_man, fill_order = zip_manifest(FILLED_HWPX)

    print(f"\n  Entry order identical: {orig_order == fill_order}")
    print(f"  First entry: orig='{orig_order[0]}'  fill='{fill_order[0]}'")

    print(f"\n  {'File':<48} {'OrigSz':>10} {'FillSz':>10} {'Delta':>10} Status")
    print(f"  {'-'*48} {'-'*10} {'-'*10} {'-'*10} ------")
    only_orig = set(orig_man) - set(fill_man)
    only_fill = set(fill_man) - set(orig_man)
    for n in sorted(set(list(orig_man.keys()) + list(fill_man.keys()))):
        if n in only_orig:
            oz, _, _ = orig_man[n]
            print(f"  {n:<48} {oz:>10}          --          -- ONLY_IN_ORIG")
        elif n in only_fill:
            fz, _, _ = fill_man[n]
            print(f"  {n:<48}         -- {fz:>10}          -- ONLY_IN_FILL")
        else:
            oz, ocrc, _ = orig_man[n]
            fz, fcrc, _ = fill_man[n]
            status = "CHANGED" if (oz != fz or ocrc != fcrc) else "identical"
            print(f"  {n:<48} {oz:>10} {fz:>10} {fz-oz:>+10} {status}")
    print()

    # ── 10. Byte diff count across whole file ─────────────────────────────────
    print(DIVIDER)
    print("STEP 10: Changed byte ranges in section0.xml (diff regions)")
    print(DIVIDER)

    # Build list of changed regions
    regions = []
    i = 0
    max_i = min(len(orig_bytes), len(fill_bytes))
    while i < max_i:
        if orig_bytes[i] != fill_bytes[i]:
            j = i
            while j < max_i and orig_bytes[j] != fill_bytes[j]:
                j += 1
            regions.append((i, j))
            i = j
        else:
            i += 1

    trailing_bytes = abs(len(orig_bytes) - len(fill_bytes)) if len(orig_bytes) != len(fill_bytes) else 0

    print(f"\n  Changed byte regions: {len(regions)}")
    print(f"  Trailing bytes (length difference): {trailing_bytes}")

    print(f"\n  {'#':<5} {'Start':>10} {'End':>10} {'OrigLen':>8} {'FillLen':>8}")
    print(f"  {'-'*5} {'-'*10} {'-'*10} {'-'*8} {'-'*8}")
    for ri, (rs, re_) in enumerate(regions[:60]):
        print(f"  {ri:<5} {rs:>10} {re_:>10} {re_-rs:>8}        ?")
    if len(regions) > 60:
        print(f"  ... and {len(regions)-60} more regions")
    print()

    # ── Summary ───────────────────────────────────────────────────────────────
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"  section0.xml size change: {len(orig_bytes):,} -> {len(fill_bytes):,} ({len(fill_bytes)-len(orig_bytes):+,} bytes)")
    print(f"  Total elements: {len(orig_all):,} -> {len(fill_all):,} ({len(fill_all)-len(orig_all):+,})")
    print(f"  hp:t elements:  orig={len(orig_t_list)} fill={len(fill_t_list)} delta={len(fill_t_list)-len(orig_t_list):+}")
    print(f"  hp:run elements: orig={len(orig_runs)} fill={len(fill_runs)} delta={len(fill_runs)-len(orig_runs):+}")
    print(f"  hp:tc cells: orig={len(orig_cells)} fill={len(fill_cells)}")
    print(f"  Modified cells: {len(cell_diffs)}")
    print(f"  Injected markers: {unique_markers}")
    print(f"  Changed byte regions: {len(regions)}")
    print(f"  Other modified ZIP entries: {[n for n in set(list(orig_man))|set(list(fill_man)) if orig_man.get(n,(0,0,0)) != fill_man.get(n,(0,0,0)) and n != SECTION_PATH]}")
    print()

if __name__ == "__main__":
    main()
