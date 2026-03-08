#!/usr/bin/env python3
"""
Deep audit of section0.xml inside form_pass1.hwpx vs original form_to_fillout.hwpx.
Checks XML declaration, namespaces, structure integrity, encoding, and more.
"""

import zipfile
import io
import sys
from lxml import etree
from collections import defaultdict

# ── paths ──────────────────────────────────────────────────────────────────────
ORIGINAL = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx"
FILLED   = "/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx"
SECTION  = "Contents/section0.xml"

NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
}

issues = []
info   = []

def log_issue(category, msg):
    issues.append(f"[{category}] {msg}")

def log_info(msg):
    info.append(f"  {msg}")

# ── helper: extract raw bytes from zip ────────────────────────────────────────
def read_raw(path, member):
    with zipfile.ZipFile(path) as z:
        return z.read(member)

# ── helper: parse XML keeping raw bytes for declaration check ─────────────────
def parse_xml(raw):
    return etree.fromstring(raw)

# ── 1. XML DECLARATION (byte-by-byte) ─────────────────────────────────────────
def check_xml_declaration(label, raw):
    print(f"\n{'='*60}")
    print(f"[1] XML DECLARATION — {label}")
    print(f"{'='*60}")

    # Show first 200 bytes as hex + text
    head = raw[:200]
    print(f"First 200 bytes (repr): {repr(head[:200])}")

    # Required declaration
    REQUIRED = b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    ALT_REQUIRED = b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"

    line1 = raw.split(b'\n')[0] if b'\n' in raw[:300] else raw[:200]
    print(f"Line 1: {repr(line1)}")

    if raw.startswith(b'<?xml'):
        # Check double vs single quotes
        if b'"' not in raw[:80]:
            log_issue("XML-DECL", f"{label}: Uses single quotes instead of double quotes in XML declaration")
        else:
            log_info(f"{label}: Uses double quotes — OK")

        # Check standalone
        if b'standalone="yes"' in raw[:100]:
            log_info(f"{label}: standalone=\"yes\" present — OK")
        elif b"standalone='yes'" in raw[:100]:
            log_issue("XML-DECL", f"{label}: standalone uses single quotes")
        else:
            log_issue("XML-DECL", f"{label}: standalone=\"yes\" is MISSING")

        # Check newline before root
        decl_end = raw.index(b'?>') + 2
        between = raw[decl_end:decl_end+10]
        print(f"Bytes between declaration end and root: {repr(between)}")
        if between.startswith(b'\n') or between.startswith(b'\r\n'):
            # Check if it's JUST whitespace/newline
            root_start = raw.index(b'<', decl_end)
            gap = raw[decl_end:root_start]
            if gap == b'\n' or gap == b'\r\n' or gap == b'\r':
                log_info(f"{label}: Single newline before root (common, usually OK for HWP)")
                print(f"Gap before root element: {repr(gap)}")
            else:
                log_issue("XML-DECL", f"{label}: Non-trivial whitespace before root: {repr(gap)}")
        elif between == b'':
            log_issue("XML-DECL", f"{label}: No content after declaration (file truncated?)")
        else:
            print(f"No newline before root — gap: {repr(between[:20])}")
    else:
        log_issue("XML-DECL", f"{label}: Does NOT start with XML declaration")

# ── 2. NAMESPACE DECLARATIONS ─────────────────────────────────────────────────
def check_namespaces(label, root):
    print(f"\n{'='*60}")
    print(f"[2] NAMESPACE DECLARATIONS — {label}")
    print(f"{'='*60}")
    ns = dict(root.nsmap)
    print(f"Namespaces on root: {ns}")
    return ns

def compare_namespaces(orig_ns, fill_ns):
    print(f"\n[2b] NAMESPACE COMPARISON")
    for prefix, uri in orig_ns.items():
        if prefix not in fill_ns:
            log_issue("NAMESPACE", f"Missing prefix '{prefix}' -> {uri} in filled file")
        elif fill_ns[prefix] != uri:
            log_issue("NAMESPACE", f"Prefix '{prefix}' URI mismatch: orig={uri}, filled={fill_ns[prefix]}")
        else:
            log_info(f"'{prefix}' -> {uri} — matches")
    for prefix, uri in fill_ns.items():
        if prefix not in orig_ns:
            log_issue("NAMESPACE", f"Extra prefix in filled file: '{prefix}' -> {uri}")

# ── 3. NESTED hp:p ────────────────────────────────────────────────────────────
def check_nested_hp_p(label, root):
    print(f"\n{'='*60}")
    print(f"[3] NESTED hp:p — {label}")
    print(f"{'='*60}")
    hp_p = '{http://www.hancom.co.kr/hwpml/2011/paragraph}p'
    count = 0
    for p in root.iter(hp_p):
        for child in p:
            if child.tag == hp_p:
                count += 1
                parent_attrib = dict(p.attrib)
                log_issue("NESTED-P", f"{label}: hp:p inside hp:p! Parent paraPrIDRef={parent_attrib.get('paraPrIDRef','?')}, child paraPrIDRef={child.get('paraPrIDRef','?')}")
    if count == 0:
        log_info(f"{label}: No nested hp:p found — OK")
    else:
        print(f"  Found {count} nested hp:p instances!")

# ── 4. EMPTY hp:run ───────────────────────────────────────────────────────────
def check_empty_runs(label, root):
    print(f"\n{'='*60}")
    print(f"[4] EMPTY hp:run (no hp:t child) — {label}")
    print(f"{'='*60}")
    hp_run = '{http://www.hancom.co.kr/hwpml/2011/paragraph}run'
    hp_t   = '{http://www.hancom.co.kr/hwpml/2011/paragraph}t'
    count = 0
    for run in root.iter(hp_run):
        has_t = any(child.tag == hp_t for child in run)
        if not has_t:
            count += 1
            if count <= 10:  # limit output
                log_issue("EMPTY-RUN", f"{label}: hp:run without hp:t, charPrIDRef={run.get('charPrIDRef','?')}")
    if count == 0:
        log_info(f"{label}: No empty hp:run found — OK")
    elif count > 10:
        log_issue("EMPTY-RUN", f"{label}: {count} total empty hp:run elements (showing first 10)")
    else:
        print(f"  Found {count} empty hp:run instances!")

# ── 5. MISSING REQUIRED ATTRIBUTES ───────────────────────────────────────────
def check_required_attrs(label, root):
    print(f"\n{'='*60}")
    print(f"[5] MISSING REQUIRED ATTRIBUTES — {label}")
    print(f"{'='*60}")
    hp_p   = '{http://www.hancom.co.kr/hwpml/2011/paragraph}p'
    hp_run = '{http://www.hancom.co.kr/hwpml/2011/paragraph}run'

    missing_para   = 0
    missing_run    = 0
    para_samples   = []
    run_samples    = []

    for p in root.iter(hp_p):
        if 'paraPrIDRef' not in p.attrib:
            missing_para += 1
            if len(para_samples) < 5:
                para_samples.append(dict(p.attrib))

    for run in root.iter(hp_run):
        if 'charPrIDRef' not in run.attrib:
            missing_run += 1
            if len(run_samples) < 5:
                run_samples.append(dict(run.attrib))

    if missing_para == 0:
        log_info(f"{label}: All hp:p have paraPrIDRef — OK")
    else:
        log_issue("ATTR", f"{label}: {missing_para} hp:p elements missing paraPrIDRef")
        for s in para_samples:
            print(f"    Sample: {s}")

    if missing_run == 0:
        log_info(f"{label}: All hp:run have charPrIDRef — OK")
    else:
        log_issue("ATTR", f"{label}: {missing_run} hp:run elements missing charPrIDRef")
        for s in run_samples:
            print(f"    Sample: {s}")

# ── 6. TABLE STRUCTURE INTEGRITY ──────────────────────────────────────────────
def check_tables(label, root):
    print(f"\n{'='*60}")
    print(f"[6] TABLE STRUCTURE INTEGRITY — {label}")
    print(f"{'='*60}")

    hp  = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
    hp_tbl    = f'{{{hp}}}tbl'
    hp_tr     = f'{{{hp}}}tr'
    hp_tc     = f'{{{hp}}}tc'
    hp_sublist= f'{{{hp}}}subList'
    hp_p      = f'{{{hp}}}p'

    tables = list(root.iter(hp_tbl))
    print(f"  Total tables found: {len(tables)}")

    table_issues = []
    table_stats  = []

    for t_idx, tbl in enumerate(tables):
        # get rowCnt / colCnt attributes
        row_cnt_attr = int(tbl.get('rowCnt', -1))
        col_cnt_attr = int(tbl.get('colCnt', -1))

        # count actual tr elements (direct children)
        actual_rows = [c for c in tbl if c.tag == hp_tr]
        actual_row_count = len(actual_rows)

        row_mismatch = (row_cnt_attr != -1 and actual_row_count != row_cnt_attr)

        col_mismatches = []
        chain_errors   = []

        for r_idx, tr in enumerate(actual_rows):
            tcs = [c for c in tr if c.tag == hp_tc]
            actual_cols = len(tcs)
            if col_cnt_attr != -1 and actual_cols != col_cnt_attr:
                col_mismatches.append((r_idx, actual_cols))

            for c_idx, tc in enumerate(tcs):
                # check subList > p chain
                sublists = [c for c in tc if c.tag == hp_sublist]
                if not sublists:
                    chain_errors.append(f"T{t_idx}[r{r_idx}][c{c_idx}] missing subList")
                else:
                    for sl in sublists:
                        paras = [c for c in sl if c.tag == hp_p]
                        if not paras:
                            chain_errors.append(f"T{t_idx}[r{r_idx}][c{c_idx}].subList missing hp:p")

        stat = {
            'idx': t_idx,
            'rowCnt_attr': row_cnt_attr,
            'colCnt_attr': col_cnt_attr,
            'actual_rows': actual_row_count,
            'row_mismatch': row_mismatch,
            'col_mismatches': col_mismatches,
            'chain_errors': chain_errors,
        }
        table_stats.append(stat)

        if row_mismatch:
            log_issue("TABLE", f"{label} T{t_idx}: rowCnt attr={row_cnt_attr} but actual rows={actual_row_count}")
        if col_mismatches:
            for r_idx, ac in col_mismatches:
                log_issue("TABLE", f"{label} T{t_idx}[row{r_idx}]: colCnt attr={col_cnt_attr} but actual cols={ac}")
        for err in chain_errors:
            log_issue("TABLE", f"{label} {err}: broken hp:tbl>hp:tr>hp:tc>hp:subList>hp:p chain")

    # Summary
    ok_tables = [s for s in table_stats if not s['row_mismatch'] and not s['col_mismatches'] and not s['chain_errors']]
    print(f"  Tables with no structural issues: {len(ok_tables)}/{len(tables)}")
    for s in table_stats:
        if s['row_mismatch'] or s['col_mismatches'] or s['chain_errors']:
            print(f"  T{s['idx']}: rowCnt={s['rowCnt_attr']}/actual={s['actual_rows']}, "
                  f"colMismatches={len(s['col_mismatches'])}, chainErrors={len(s['chain_errors'])}")

    return len(tables), table_stats

# ── 7. ORPHANED ELEMENTS ──────────────────────────────────────────────────────
VALID_CHILDREN = {
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}p':       {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}run',
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}ctrl',
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}linesegarray',
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}bullet',
    },
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}run':     {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}t',
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}img',
    },
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}subList': {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}p',
    },
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}tbl':     {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}tr',
    },
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}tr':      {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}tc',
    },
    '{http://www.hancom.co.kr/hwpml/2011/paragraph}tc':      {
        '{http://www.hancom.co.kr/hwpml/2011/paragraph}subList',
    },
}

def check_orphaned_elements(label, root):
    print(f"\n{'='*60}")
    print(f"[7] ORPHANED / UNEXPECTED ELEMENTS — {label}")
    print(f"{'='*60}")
    count = 0
    for parent_tag, allowed_children in VALID_CHILDREN.items():
        for parent in root.iter(parent_tag):
            for child in parent:
                if child.tag not in allowed_children:
                    short_parent = parent_tag.split('}')[-1]
                    short_child  = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    count += 1
                    if count <= 20:
                        log_issue("ORPHAN", f"{label}: Unexpected <{short_child}> inside <{short_parent}>")
    if count == 0:
        log_info(f"{label}: No unexpected child elements found — OK")
    elif count > 20:
        log_issue("ORPHAN", f"{label}: {count} total unexpected children (showing first 20)")

# ── 8. CHARACTER ENCODING ─────────────────────────────────────────────────────
def check_encoding(label, root):
    print(f"\n{'='*60}")
    print(f"[8] CHARACTER ENCODING — {label}")
    print(f"{'='*60}")
    hp_t = '{http://www.hancom.co.kr/hwpml/2011/paragraph}t'
    count = 0
    problem_count = 0
    for t in root.iter(hp_t):
        text = t.text or ''
        count += 1
        for i, ch in enumerate(text):
            code = ord(ch)
            # Null bytes or control chars (except tab, newline, cr) are problematic
            if code == 0:
                problem_count += 1
                log_issue("ENCODING", f"{label}: Null byte (0x00) in hp:t text at char {i}: context={repr(text[:30])}")
            elif code < 0x09 or (0x0B <= code <= 0x0C) or (0x0E <= code <= 0x1F):
                problem_count += 1
                log_issue("ENCODING", f"{label}: Control char U+{code:04X} in hp:t at char {i}: context={repr(text[:30])}")
    if problem_count == 0:
        log_info(f"{label}: No encoding issues in {count} hp:t elements — OK")
    print(f"  Total hp:t elements scanned: {count}, issues: {problem_count}")

# ── 9. hp:linesegarray PRESERVATION ──────────────────────────────────────────
def collect_linesegarray(label, root):
    """Returns set of paraPrIDRef values for paragraphs that HAVE linesegarray."""
    print(f"\n{'='*60}")
    print(f"[9] hp:linesegarray — {label}")
    print(f"{'='*60}")
    hp_p   = '{http://www.hancom.co.kr/hwpml/2011/paragraph}p'
    hp_lsa = '{http://www.hancom.co.kr/hwpml/2011/paragraph}linesegarray'

    total_p    = 0
    p_with_lsa = 0
    ids_with_lsa = set()

    for p in root.iter(hp_p):
        total_p += 1
        has_lsa = any(c.tag == hp_lsa for c in p)
        if has_lsa:
            p_with_lsa += 1
            ids_with_lsa.add(p.get('paraPrIDRef', 'NO-ID'))

    print(f"  Total hp:p: {total_p}, with linesegarray: {p_with_lsa}")
    return ids_with_lsa

def compare_linesegarray(orig_ids, fill_ids, label_o, label_f):
    print(f"\n[9b] linesegarray COMPARISON")
    lost = orig_ids - fill_ids
    gained = fill_ids - orig_ids
    if not lost and not gained:
        log_info("linesegarray: No gain/loss — OK (same set of paraPrIDRef values)")
    else:
        if lost:
            log_issue("LINESEG", f"linesegarray LOST in filled file for paraPrIDRef IDs: {sorted(lost)[:20]}")
        if gained:
            log_info(f"linesegarray GAINED in filled file for paraPrIDRef IDs: {sorted(gained)[:20]} (may be OK)")

# ── EXTRA: count elements summary ─────────────────────────────────────────────
def element_summary(label, root):
    print(f"\n{'='*60}")
    print(f"[SUMMARY] ELEMENT COUNTS — {label}")
    print(f"{'='*60}")
    tags = defaultdict(int)
    for el in root.iter():
        short = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        tags[short] += 1
    for tag, count in sorted(tags.items(), key=lambda x: -x[1])[:30]:
        print(f"  {tag}: {count}")
    return dict(tags)

def compare_element_counts(orig_counts, fill_counts):
    print(f"\n[SUMMARY-DIFF] ELEMENT COUNT DIFFERENCES (orig vs filled)")
    all_tags = set(orig_counts) | set(fill_counts)
    diffs = []
    for tag in sorted(all_tags):
        o = orig_counts.get(tag, 0)
        f = fill_counts.get(tag, 0)
        if o != f:
            diffs.append((tag, o, f, f - o))
    if not diffs:
        print("  No differences in element counts — perfect match")
    else:
        print(f"  {'Tag':<25} {'Original':>10} {'Filled':>10} {'Delta':>8}")
        print(f"  {'-'*55}")
        for tag, o, f, d in sorted(diffs, key=lambda x: abs(x[3]), reverse=True):
            flag = " <-- ISSUE?" if abs(d) > 50 else ""
            print(f"  {tag:<25} {o:>10} {f:>10} {d:>+8}{flag}")

# ── EXTRA: check PrintMethod in settings ──────────────────────────────────────
def check_settings_printmethod(path):
    print(f"\n{'='*60}")
    print(f"[EXTRA] settings.xml PrintMethod check — {path.split('/')[-1]}")
    print(f"{'='*60}")
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            settings_candidates = [n for n in names if 'settings' in n.lower() or 'Settings' in n]
            print(f"  Settings-related members: {settings_candidates}")
            for name in settings_candidates:
                data = z.read(name)
                if b'PrintMethod' in data:
                    # Extract context
                    idx = data.index(b'PrintMethod')
                    ctx = data[max(0,idx-30):idx+60]
                    print(f"  PrintMethod context in {name}: {repr(ctx)}")
                    if b'PrintMethod=\\"4\\"' in data or b'PrintMethod="4"' in data or b'PrintMethod=4' in data:
                        log_issue("SETTINGS", f"{path}: PrintMethod=4 found in {name} — this HALVES PDF pages!")
                    else:
                        log_info(f"PrintMethod context: {repr(ctx)}")
                else:
                    print(f"  No PrintMethod in {name}")
    except Exception as e:
        print(f"  Error checking settings: {e}")

# ── EXTRA: ZIP integrity ──────────────────────────────────────────────────────
def check_zip_integrity(path):
    print(f"\n{'='*60}")
    print(f"[EXTRA] ZIP INTEGRITY — {path.split('/')[-1]}")
    print(f"{'='*60}")
    try:
        with zipfile.ZipFile(path) as z:
            result = z.testzip()
            if result is None:
                log_info(f"{path}: ZIP integrity OK — no bad files")
            else:
                log_issue("ZIP", f"{path}: First bad file in ZIP: {result}")

            # Check mimetype is first and STORED
            names = z.namelist()
            infos = z.infolist()
            print(f"  Total ZIP members: {len(names)}")
            print(f"  First member: {names[0] if names else 'NONE'}")
            if names and names[0] == 'mimetype':
                mime_info = infos[0]
                if mime_info.compress_type == zipfile.ZIP_STORED:
                    log_info(f"mimetype is first and STORED — OK")
                else:
                    log_issue("ZIP", f"mimetype is first but NOT stored (compress_type={mime_info.compress_type})")
                mime_content = z.read('mimetype')
                print(f"  mimetype content: {repr(mime_content)}")
            else:
                log_issue("ZIP", f"mimetype is NOT first member! First is: {names[0] if names else 'NONE'}")
    except zipfile.BadZipFile as e:
        log_issue("ZIP", f"{path}: BadZipFile: {e}")
    except Exception as e:
        log_issue("ZIP", f"{path}: Error: {e}")

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("HWPX SECTION0.XML DEEP AUDIT")
    print(f"Original: {ORIGINAL}")
    print(f"Filled:   {FILLED}")
    print("=" * 70)

    # Load raw bytes
    print("\nLoading files...")
    orig_raw  = read_raw(ORIGINAL, SECTION)
    fill_raw  = read_raw(FILLED,   SECTION)
    print(f"  Original section0.xml size: {len(orig_raw):,} bytes")
    print(f"  Filled   section0.xml size: {len(fill_raw):,} bytes")
    print(f"  Size delta: {len(fill_raw) - len(orig_raw):+,} bytes")

    # Parse XML
    orig_root = parse_xml(orig_raw)
    fill_root = parse_xml(fill_raw)

    # ── Run all checks ────────────────────────────────────────────────────────

    # 1. XML Declaration
    check_xml_declaration("ORIGINAL", orig_raw)
    check_xml_declaration("FILLED",   fill_raw)

    # 2. Namespaces
    orig_ns = check_namespaces("ORIGINAL", orig_root)
    fill_ns = check_namespaces("FILLED",   fill_root)
    compare_namespaces(orig_ns, fill_ns)

    # 3. Nested hp:p
    check_nested_hp_p("ORIGINAL", orig_root)
    check_nested_hp_p("FILLED",   fill_root)

    # 4. Empty hp:run
    check_empty_runs("ORIGINAL", orig_root)
    check_empty_runs("FILLED",   fill_root)

    # 5. Missing required attributes
    check_required_attrs("ORIGINAL", orig_root)
    check_required_attrs("FILLED",   fill_root)

    # 6. Table structure
    orig_tcount, orig_tstats = check_tables("ORIGINAL", orig_root)
    fill_tcount, fill_tstats = check_tables("FILLED",   fill_root)
    if orig_tcount != fill_tcount:
        log_issue("TABLE", f"Table count mismatch: original={orig_tcount}, filled={fill_tcount}")
    else:
        log_info(f"Table count matches: {orig_tcount}")

    # 7. Orphaned elements
    check_orphaned_elements("ORIGINAL", orig_root)
    check_orphaned_elements("FILLED",   fill_root)

    # 8. Character encoding
    check_encoding("ORIGINAL", orig_root)
    check_encoding("FILLED",   fill_root)

    # 9. linesegarray
    orig_lsa = collect_linesegarray("ORIGINAL", orig_root)
    fill_lsa = collect_linesegarray("FILLED",   fill_root)
    compare_linesegarray(orig_lsa, fill_lsa, "ORIGINAL", "FILLED")

    # Extra: element counts
    orig_counts = element_summary("ORIGINAL", orig_root)
    fill_counts = element_summary("FILLED",   fill_root)
    compare_element_counts(orig_counts, fill_counts)

    # Extra: PrintMethod
    check_settings_printmethod(ORIGINAL)
    check_settings_printmethod(FILLED)

    # Extra: ZIP integrity
    check_zip_integrity(ORIGINAL)
    check_zip_integrity(FILLED)

    # ── Final report ──────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("FINAL AUDIT REPORT")
    print("=" * 70)

    print(f"\n{'─'*60}")
    print(f"INFO ({len(info)} items):")
    print(f"{'─'*60}")
    for i in info:
        print(i)

    print(f"\n{'─'*60}")
    print(f"ISSUES ({len(issues)} total):")
    print(f"{'─'*60}")
    if not issues:
        print("  ** NO ISSUES FOUND — files look structurally clean **")
    else:
        for i, iss in enumerate(issues, 1):
            print(f"  {i:3d}. {iss}")

    print(f"\n{'='*70}")
    print(f"AUDIT COMPLETE — {len(issues)} issues, {len(info)} info items")
    print(f"{'='*70}")

    return len(issues)

if __name__ == '__main__':
    sys.exit(0 if main() == 0 else 1)
