#!/usr/bin/env python3
"""
Deep audit of HWPX text content for crash-causing issues.
Checks both original and filled HWPX files.

HWPX structure:  hp:tbl -> hp:tr -> hp:tc -> hp:subList -> hp:p -> hp:run -> hp:t
Namespace:       hp = http://www.hancom.co.kr/hwpml/2011/paragraph
"""

import zipfile
import re
from lxml import etree

ORIGINAL_PATH = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx"
FILLED_PATH   = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx"

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"

TAG_T       = f"{{{HP}}}t"
TAG_TBL     = f"{{{HP}}}tbl"
TAG_TR      = f"{{{HP}}}tr"
TAG_TC      = f"{{{HP}}}tc"
TAG_P       = f"{{{HP}}}p"
TAG_RUN     = f"{{{HP}}}run"
TAG_SUBLIST = f"{{{HP}}}subList"

ILLEGAL_CTRL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

# -----------------------------------------------------------------------
# helpers
# -----------------------------------------------------------------------

def load_section0(hwpx_path: str) -> bytes:
    with zipfile.ZipFile(hwpx_path, "r") as z:
        for n in z.namelist():
            if "section0" in n.lower() and n.endswith(".xml"):
                return z.read(n)
    raise FileNotFoundError(f"section0 not found in {hwpx_path}")


def build_parent_map(root: etree._Element) -> dict:
    return {c: p for p in root.iter() for c in p}


def get_context(el: etree._Element, parent_map: dict, root: etree._Element) -> str:
    """Walk ancestors to get TBL_IDX / ROW_IDX / COL_IDX context string."""
    all_tbls = None  # lazy
    cur = el
    tbl_idx = row_idx = col_idx = None

    while cur in parent_map:
        cur = parent_map[cur]
        if cur.tag == TAG_TC and col_idx is None:
            par = parent_map.get(cur)
            if par is not None:
                col_idx = list(par).index(cur)
        elif cur.tag == TAG_TR and row_idx is None:
            par = parent_map.get(cur)
            if par is not None:
                row_idx = list(par).index(cur)
        elif cur.tag == TAG_TBL and tbl_idx is None:
            if all_tbls is None:
                all_tbls = list(root.iter(TAG_TBL))
            try:
                tbl_idx = all_tbls.index(cur)
            except ValueError:
                tbl_idx = "?"
    return f"T{tbl_idx} R{row_idx} C{col_idx}"


# -----------------------------------------------------------------------
# main audit
# -----------------------------------------------------------------------

def check_all(label: str, data: bytes) -> dict:
    print(f"\n{'='*70}")
    print(f"  AUDITING: {label}")
    print(f"{'='*70}")

    root = etree.fromstring(data)

    # Confirm we can see hp:t elements
    all_t_els = list(root.iter(TAG_T))
    all_tbls   = list(root.iter(TAG_TBL))
    all_p      = list(root.iter(TAG_P))
    all_run    = list(root.iter(TAG_RUN))
    print(f"  hp:tbl elements : {len(all_tbls)}")
    print(f"  hp:p  elements  : {len(all_p)}")
    print(f"  hp:run elements : {len(all_run)}")
    print(f"  hp:t  elements  : {len(all_t_els)}")

    parent_map = build_parent_map(root)

    issues = {
        "control_chars":   [],
        "long_text":       [],
        "bad_utf8":        [],
        "empty_t":         [],
        "whitespace_only": [],
        "mixed_content":   [],
    }

    for el in all_t_els:
        text = el.text  # lxml: None or str
        loc  = get_context(el, parent_map, root)

        # 1. Control characters (illegal in XML 1.0)
        if text and ILLEGAL_CTRL_RE.search(text):
            matches = [(m.start(), hex(ord(m.group()))) for m in ILLEGAL_CTRL_RE.finditer(text)]
            issues["control_chars"].append((loc, repr(text[:80]), matches))

        # 2. Extremely long text (potential buffer overflow)
        if text and len(text) > 1000:
            issues["long_text"].append((loc, len(text), text[:80] + "..."))

        # 4. Korean / UTF-8 encoding validity
        if text:
            try:
                text.encode("utf-8")
            except (UnicodeEncodeError, UnicodeDecodeError) as e:
                issues["bad_utf8"].append((loc, str(e), repr(text[:40])))

        # 5. Empty hp:t
        if text is None or text == "":
            issues["empty_t"].append(loc)

        # 6. Whitespace-only (not empty, but no visible chars)
        if text and text.strip() == "":
            issues["whitespace_only"].append((loc, repr(text)))

        # 7. Mixed content: hp:t with children (e.g., lineBreak), check for:
        #    a) text AND children simultaneously
        #    b) tail text on child elements (would be outside the hp:t text node)
        children = list(el)
        if children:
            if text and text.strip():
                issues["mixed_content"].append(
                    (loc, "hp:t has both .text AND child elements", repr(text[:60]))
                )
            for child in children:
                if child.tail and child.tail.strip():
                    issues["mixed_content"].append(
                        (loc, f"child <{child.tag}> has non-empty .tail", repr(child.tail[:60]))
                    )

    # --- Check 3: serialization round-trip ---
    print("\n  [CHECK 3] Serialization check for unescaped &/< in raw bytes...")
    raw_bytes = etree.tostring(root, encoding="utf-8", xml_declaration=True)
    # Any & NOT followed by amp;/lt;/gt;/apos;/quot;/# is unescaped
    bad_amp = re.findall(rb'&(?!amp;|lt;|gt;|apos;|quot;|#[0-9a-fA-F])', raw_bytes)
    if bad_amp:
        print(f"    WARNING: {len(bad_amp)} unescaped '&' found in serialized bytes!")
        # Show up to 5 examples with context
        for m in re.finditer(rb'&(?!amp;|lt;|gt;|apos;|quot;|#[0-9a-fA-F])', raw_bytes):
            ctx_bytes = raw_bytes[max(0, m.start()-30):m.start()+40]
            print(f"      ... {ctx_bytes} ...")
            if bad_amp.index(m.group()) >= 4:
                break
    else:
        print(f"    OK: No unescaped & in serialized output.")

    # --- Check 8: round-trip parse ---
    print("\n  [CHECK 8] Round-trip parse test...")
    try:
        root2    = etree.fromstring(raw_bytes)
        cnt_orig = len(list(root.iter()))
        cnt_rt   = len(list(root2.iter()))
        t_orig   = len(list(root.iter(TAG_T)))
        t_rt     = len(list(root2.iter(TAG_T)))
        p_orig   = len(list(root.iter(TAG_P)))
        p_rt     = len(list(root2.iter(TAG_P)))
        print(f"    Total elements — before: {cnt_orig}, after: {cnt_rt} — {'OK' if cnt_orig==cnt_rt else 'MISMATCH!'}")
        print(f"    hp:p  elements — before: {p_orig},   after: {p_rt}   — {'OK' if p_orig==p_rt else 'MISMATCH!'}")
        print(f"    hp:t  elements — before: {t_orig},   after: {t_rt}   — {'OK' if t_orig==t_rt else 'MISMATCH!'}")
    except Exception as e:
        print(f"    PARSE ERROR on round-trip: {e}")

    # --- Report all findings ---
    print(f"\n  --- ISSUE SUMMARY ---")

    print(f"\n  [1] Control chars (illegal XML 1.0, 0x00-0x08,0x0B,0x0C,0x0E-0x1F): {len(issues['control_chars'])}")
    for loc, snippet, matches in issues["control_chars"][:30]:
        print(f"       {loc}: chars={matches}")
        print(f"              text={snippet}")

    print(f"\n  [2] Long hp:t (>1000 chars): {len(issues['long_text'])}")
    for loc, length, snippet in issues["long_text"][:20]:
        print(f"       {loc}: {length} chars — '{snippet}'")

    print(f"\n  [4] Bad UTF-8: {len(issues['bad_utf8'])}")
    for loc, err, snippet in issues["bad_utf8"][:20]:
        print(f"       {loc}: {err}  text={snippet}")

    print(f"\n  [5] Empty hp:t (text is None or ''): {len(issues['empty_t'])}")
    for loc in issues["empty_t"][:40]:
        print(f"       {loc}")
    if len(issues["empty_t"]) > 40:
        print(f"       ... and {len(issues['empty_t'])-40} more")

    print(f"\n  [6] Whitespace-only hp:t: {len(issues['whitespace_only'])}")
    for loc, ws in issues["whitespace_only"][:20]:
        print(f"       {loc}: {ws}")

    print(f"\n  [7] Mixed content in hp:t: {len(issues['mixed_content'])}")
    for loc, reason, snippet in issues["mixed_content"][:20]:
        print(f"       {loc}: {reason}")
        print(f"              '{snippet}'")

    return issues


# -----------------------------------------------------------------------
# comparison
# -----------------------------------------------------------------------

def compare_empty(orig_issues: dict, fill_issues: dict):
    print(f"\n{'='*70}")
    print(f"  COMPARISON: Empty hp:t counts  (original vs filled)")
    print(f"{'='*70}")
    o = len(orig_issues["empty_t"])
    f = len(fill_issues["empty_t"])
    d = f - o
    print(f"  Original  empty hp:t: {o}")
    print(f"  Filled    empty hp:t: {f}")
    print(f"  Delta:               {d:+d}")
    if d > 100:
        print(f"  WARNING: Very large increase — the structure-preserving fix is creating many blank hp:t elements.")
    elif d > 20:
        print(f"  NOTE: Moderate increase — may be from cleared extra paragraphs. Review if expected.")
    else:
        print(f"  OK: Minimal change.")


# -----------------------------------------------------------------------
# extra: longest texts in filled
# -----------------------------------------------------------------------

def show_longest(fill_data: bytes, n: int = 10):
    print(f"\n{'='*70}")
    print(f"  TOP {n} LONGEST hp:t texts in FILLED")
    print(f"{'='*70}")
    root = etree.fromstring(fill_data)
    parent_map = build_parent_map(root)
    lengths = []
    for el in root.iter(TAG_T):
        if el.text:
            loc = get_context(el, parent_map, root)
            lengths.append((len(el.text), loc, el.text))
    lengths.sort(reverse=True)
    for length, loc, txt in lengths[:n]:
        print(f"  {length:5d} chars  {loc}: {repr(txt[:120])}")


# -----------------------------------------------------------------------
# also check markers injected into section for pass-2
# -----------------------------------------------------------------------

def check_markers(fill_data: bytes):
    print(f"\n{'='*70}")
    print(f"  MARKER CHECK (Pass-2 injection points)")
    print(f"{'='*70}")
    root = etree.fromstring(fill_data)
    parent_map = build_parent_map(root)
    marker_re = re.compile(r"##\w+##")
    markers_found = []
    for el in root.iter(TAG_T):
        if el.text and marker_re.search(el.text):
            loc = get_context(el, parent_map, root)
            markers_found.append((loc, el.text))
    print(f"  Markers found: {len(markers_found)}")
    for loc, txt in markers_found:
        print(f"    {loc}: {repr(txt[:120])}")


# -----------------------------------------------------------------------
# main
# -----------------------------------------------------------------------

def main():
    print("Loading section0.xml from ORIGINAL...")
    orig_data = load_section0(ORIGINAL_PATH)
    print(f"  size: {len(orig_data):,} bytes")

    print("Loading section0.xml from FILLED...")
    fill_data = load_section0(FILLED_PATH)
    print(f"  size: {len(fill_data):,} bytes")

    orig_issues = check_all("ORIGINAL (form_to_fillout.hwpx)", orig_data)
    fill_issues = check_all("FILLED   (form_pass1.hwpx)",      fill_data)

    compare_empty(orig_issues, fill_issues)
    show_longest(fill_data)
    check_markers(fill_data)

    print(f"\n{'='*70}")
    print("  AUDIT COMPLETE")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
