#!/usr/bin/env python3
"""
XML Serialization Corruption Diagnostic
========================================
Investigates whether lxml's parse-then-serialize cycle corrupts HWPX XML.

Checks:
  1. Round-trip test on the ORIGINAL (unmodified) file
  2. Namespace attribute order changes
  3. Whitespace changes (inter-element, attribute values, trailing text)
  4. Self-closing vs explicit close tags (e.g., <hp:t/> vs <hp:t></hp:t>)
  5. Entity encoding differences (&#xD; vs CR, &#x9; vs tab, etc.)
  6. Attribute quoting style (single vs double quotes)
  7. serialize_xml() output correctness
"""

import difflib
import io
import re
import sys
import zipfile
from pathlib import Path

ORIGINAL_PATH = Path('/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx')
FILLED_PATH   = Path('/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/output/filled/form_pass1.hwpx')

# ── helpers ───────────────────────────────────────────────────────────────────

SECTION = 'Contents/section0.xml'


def load_section(hwpx_path: Path) -> bytes:
    with zipfile.ZipFile(hwpx_path, 'r') as zf:
        return zf.read(SECTION)


def banner(title: str):
    print()
    print('=' * 70)
    print(f'  {title}')
    print('=' * 70)


def pass_fail(ok: bool, label: str):
    status = 'PASS' if ok else 'FAIL'
    print(f'  [{status}]  {label}')


# ── lxml round-trip serialization (mirrors hwpx_editor.serialize_xml) ─────────

def roundtrip_bytes(raw: bytes) -> bytes:
    """Parse with lxml, re-serialize the same way hwpx_editor.py does."""
    from lxml import etree
    import re as _re

    raw_text = raw.decode('utf-8')
    m = _re.match(r'(<\?xml\s[^?]*\?>)', raw_text)
    HWPX_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    xml_decl = m.group(1) if m else HWPX_XML_DECL

    root = etree.fromstring(raw)
    body = etree.tostring(root, xml_declaration=False, encoding='unicode')
    return (xml_decl + body).encode('utf-8')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 1  —  Round-trip test on the ORIGINAL
# ══════════════════════════════════════════════════════════════════════════════

def check1_roundtrip(orig_bytes: bytes):
    banner('CHECK 1 — Round-trip test on ORIGINAL section0.xml')

    rt_bytes = roundtrip_bytes(orig_bytes)

    if orig_bytes == rt_bytes:
        pass_fail(True, 'Byte-for-byte identical after parse+reserialize')
        return

    pass_fail(False, 'DIFFERENCES FOUND after parse+reserialize')
    print()

    orig_lines = orig_bytes.decode('utf-8', errors='replace').splitlines(keepends=True)
    rt_lines   = rt_bytes.decode('utf-8', errors='replace').splitlines(keepends=True)

    diff = list(difflib.unified_diff(
        orig_lines, rt_lines,
        fromfile='original', tofile='round-tripped',
        n=3, lineterm=''
    ))

    if not diff:
        print('  (unified_diff produced no output — differences are in raw bytes only)')
        # byte-level search
        min_len = min(len(orig_bytes), len(rt_bytes))
        first_diff = next((i for i in range(min_len) if orig_bytes[i] != rt_bytes[i]), None)
        if first_diff is not None:
            print(f'  First differing byte at offset {first_diff}')
            print(f'    original : {orig_bytes[first_diff-20:first_diff+20]!r}')
            print(f'    roundtrip: {rt_bytes[first_diff-20:first_diff+20]!r}')
        if len(orig_bytes) != len(rt_bytes):
            print(f'  Length: original={len(orig_bytes)}, roundtrip={len(rt_bytes)}')
    else:
        # Show first 80 diff lines maximum
        shown = diff[:80]
        for line in shown:
            print('  ' + line.rstrip('\n'))
        if len(diff) > 80:
            print(f'  ... ({len(diff) - 80} more diff lines truncated)')

    print()
    print(f'  Total: original={len(orig_bytes)} bytes, roundtrip={len(rt_bytes)} bytes')
    print(f'  Diff lines: {len(diff)}')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 2  —  Namespace attribute order on the root element
# ══════════════════════════════════════════════════════════════════════════════

def extract_xmlns_attrs(xml_bytes: bytes) -> list[str]:
    """Return xmlns:* attributes from the root element opening tag, in order."""
    text = xml_bytes.decode('utf-8', errors='replace')
    # Match the root opening tag (may span multiple lines, up to first >)
    m = re.search(r'<[a-zA-Z][^>]*>', text[:4000], re.DOTALL)
    if not m:
        return []
    tag_text = m.group(0)
    return re.findall(r'xmlns(?::[a-zA-Z0-9_-]+)?="[^"]*"', tag_text)


def check2_namespace_order(orig_bytes: bytes):
    banner('CHECK 2 — Namespace attribute order on root element')

    from lxml import etree

    rt_bytes = roundtrip_bytes(orig_bytes)

    orig_ns = extract_xmlns_attrs(orig_bytes)
    rt_ns   = extract_xmlns_attrs(rt_bytes)

    print(f'  Original xmlns attrs ({len(orig_ns)}):')
    for a in orig_ns:
        print(f'    {a}')

    print(f'\n  Round-trip xmlns attrs ({len(rt_ns)}):')
    for a in rt_ns:
        print(f'    {a}')

    if orig_ns == rt_ns:
        pass_fail(True, 'Namespace attribute order preserved')
    else:
        pass_fail(False, 'Namespace attribute ORDER CHANGED')
        for i, (o, r) in enumerate(zip(orig_ns, rt_ns)):
            if o != r:
                print(f'    Position {i}: orig={o!r}  rt={r!r}')
        extras_orig = set(orig_ns) - set(rt_ns)
        extras_rt   = set(rt_ns)  - set(orig_ns)
        if extras_orig:
            print(f'    Only in original : {extras_orig}')
        if extras_rt:
            print(f'    Only in roundtrip: {extras_rt}')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 3  —  Whitespace changes
# ══════════════════════════════════════════════════════════════════════════════

def check3_whitespace(orig_bytes: bytes):
    banner('CHECK 3 — Whitespace changes')

    rt_bytes = roundtrip_bytes(orig_bytes)

    orig_str = orig_bytes.decode('utf-8', errors='replace')
    rt_str   = rt_bytes.decode('utf-8', errors='replace')

    # 3a. Inter-element whitespace: text that is purely whitespace between > and <
    ws_pattern = re.compile(r'>(\s+)<')
    orig_ws = ws_pattern.findall(orig_str)
    rt_ws   = ws_pattern.findall(rt_str)

    print('  3a. Inter-element whitespace nodes:')
    if orig_ws == rt_ws:
        pass_fail(True, f'Identical ({len(orig_ws)} whitespace-only text nodes)')
    else:
        pass_fail(False, 'Inter-element whitespace CHANGED')
        diff_count = sum(1 for a, b in zip(orig_ws, rt_ws) if a != b)
        print(f'    Count: original={len(orig_ws)}, roundtrip={len(rt_ws)}, differing={diff_count}')
        # Show first few differences
        for i, (a, b) in enumerate(zip(orig_ws, rt_ws)):
            if a != b:
                print(f'    WS node #{i}: orig={a!r}  rt={b!r}')
                if i > 5:
                    break

    # 3b. Trailing whitespace in text content (<hp:t> elements)
    t_pattern = re.compile(r'<hp:t[^>]*>([^<]*)</hp:t>', re.DOTALL)
    orig_t = t_pattern.findall(orig_str)
    rt_t   = t_pattern.findall(rt_str)

    print('\n  3b. Text content in <hp:t> elements:')
    if orig_t == rt_t:
        pass_fail(True, f'Identical ({len(orig_t)} text nodes)')
    else:
        pass_fail(False, 'Text content in <hp:t> CHANGED')
        for i, (a, b) in enumerate(zip(orig_t, rt_t)):
            if a != b:
                print(f'    Text node #{i}: orig={a!r}  rt={b!r}')
                if i > 5:
                    break

    # 3c. Attribute value whitespace
    attr_pattern = re.compile(r'\w+="([^"]*)"')
    orig_attr_ws = [v for v in attr_pattern.findall(orig_str) if re.search(r'\s', v)]
    rt_attr_ws   = [v for v in attr_pattern.findall(rt_str)   if re.search(r'\s', v)]

    print('\n  3c. Attribute values containing whitespace:')
    if orig_attr_ws == rt_attr_ws:
        pass_fail(True, f'Identical ({len(orig_attr_ws)} whitespace-containing attribute values)')
    else:
        pass_fail(False, 'Attribute value whitespace CHANGED')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 4  —  Self-closing vs explicit close tags
# ══════════════════════════════════════════════════════════════════════════════

def check4_self_closing(orig_bytes: bytes):
    banner('CHECK 4 — Self-closing vs explicit close tags')

    rt_bytes = roundtrip_bytes(orig_bytes)

    orig_str = orig_bytes.decode('utf-8', errors='replace')
    rt_str   = rt_bytes.decode('utf-8', errors='replace')

    # Find self-closing tags in each
    sc_pattern = re.compile(r'<([a-zA-Z0-9:_-]+)[^>]*/>')
    orig_sc = sc_pattern.findall(orig_str)
    rt_sc   = sc_pattern.findall(rt_str)

    # Find explicit-close empty tags  <tag></tag>
    ec_pattern = re.compile(r'<([a-zA-Z0-9:_-]+)[^>]*></\1>')
    orig_ec = ec_pattern.findall(orig_str)
    rt_ec   = ec_pattern.findall(rt_str)

    from collections import Counter
    orig_sc_counts = Counter(orig_sc)
    rt_sc_counts   = Counter(rt_sc)
    orig_ec_counts = Counter(orig_ec)
    rt_ec_counts   = Counter(rt_ec)

    print(f'  Self-closing tags:   original={len(orig_sc)}  roundtrip={len(rt_sc)}')
    print(f'  Explicit-close tags: original={len(orig_ec)}  roundtrip={len(rt_ec)}')

    if orig_sc_counts == rt_sc_counts and orig_ec_counts == rt_ec_counts:
        pass_fail(True, 'No change in self-closing / explicit-close counts')
    else:
        pass_fail(False, 'Self-closing style CHANGED')
        # Show tags that gained/lost self-closing
        all_tags = set(orig_sc_counts) | set(rt_sc_counts) | set(orig_ec_counts) | set(rt_ec_counts)
        for tag in sorted(all_tags):
            osc = orig_sc_counts.get(tag, 0)
            rsc = rt_sc_counts.get(tag, 0)
            oec = orig_ec_counts.get(tag, 0)
            rec = rt_ec_counts.get(tag, 0)
            if osc != rsc or oec != rec:
                print(f'    <{tag}>: self-closing orig={osc} rt={rsc} | explicit orig={oec} rt={rec}')

    # Specifically check hp:t (most likely to be affected)
    print()
    hp_t_sc = orig_str.count('<hp:t/>')
    hp_t_ec = orig_str.count('<hp:t></hp:t>')
    rt_hp_t_sc = rt_str.count('<hp:t/>')
    rt_hp_t_ec = rt_str.count('<hp:t></hp:t>')

    print(f'  hp:t self-closing  (<hp:t/>):        original={hp_t_sc}   roundtrip={rt_hp_t_sc}')
    print(f'  hp:t explicit-close (<hp:t></hp:t>): original={hp_t_ec}   roundtrip={rt_hp_t_ec}')

    if hp_t_sc == rt_hp_t_sc and hp_t_ec == rt_hp_t_ec:
        pass_fail(True, 'hp:t tag style preserved')
    else:
        pass_fail(False, 'hp:t tag style CHANGED — lxml converts between forms')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 5  —  Entity encoding
# ══════════════════════════════════════════════════════════════════════════════

ENTITIES = {
    '&#xD;':  '\r',
    '&#x9;':  '\t',
    '&#xA;':  '\n',
    '&amp;':  '&',
    '&lt;':   '<',
    '&gt;':   '>',
    '&quot;': '"',
    '&apos;': "'",
    '&#10;':  '\n',
    '&#13;':  '\r',
    '&#9;':   '\t',
}


def check5_entity_encoding(orig_bytes: bytes):
    banner('CHECK 5 — Entity encoding')

    rt_bytes = roundtrip_bytes(orig_bytes)

    orig_str = orig_bytes.decode('utf-8', errors='replace')
    rt_str   = rt_bytes.decode('utf-8', errors='replace')

    issues = []
    for entity, char in ENTITIES.items():
        orig_count = orig_str.count(entity)
        rt_count   = rt_str.count(entity)
        # Also count literal occurrences of the character (in attribute values / text)
        # (rough heuristic — just comparing entity counts)
        if orig_count != rt_count:
            issues.append((entity, orig_count, rt_count, char))

    if not issues:
        pass_fail(True, 'Entity encoding unchanged')
    else:
        pass_fail(False, 'Entity encoding CHANGED')
        print(f'  {"Entity":<12}  {"Original":>10}  {"Roundtrip":>10}  Char')
        for entity, oc, rc, ch in issues:
            print(f'  {entity:<12}  {oc:>10}  {rc:>10}  {ch!r}')

    # Look for literal CR / TAB characters in text content (potentially dangerous)
    literal_cr_orig = orig_str.count('\r')
    literal_cr_rt   = rt_str.count('\r')
    literal_tab_orig = orig_str.count('\t')
    literal_tab_rt   = rt_str.count('\t')

    print(f'\n  Literal CR  (\\r):  original={literal_cr_orig}  roundtrip={literal_cr_rt}')
    print(f'  Literal TAB (\\t):  original={literal_tab_orig}  roundtrip={literal_tab_rt}')

    if literal_cr_orig != literal_cr_rt or literal_tab_orig != literal_tab_rt:
        pass_fail(False, 'Literal control characters changed — lxml may be normalizing them')
    else:
        pass_fail(True, 'Literal control characters unchanged')

    # Check numeric character references in attribute values
    ncr_pattern = re.compile(r'&#x?[0-9A-Fa-f]+;')
    orig_ncr = ncr_pattern.findall(orig_str)
    rt_ncr   = ncr_pattern.findall(rt_str)
    if orig_ncr == rt_ncr:
        pass_fail(True, f'Numeric character references unchanged ({len(orig_ncr)} found)')
    else:
        pass_fail(False, f'Numeric character references CHANGED: orig={len(orig_ncr)} rt={len(rt_ncr)}')
        from collections import Counter
        orig_c = Counter(orig_ncr)
        rt_c   = Counter(rt_ncr)
        for ref in sorted(set(orig_c) | set(rt_c)):
            if orig_c[ref] != rt_c[ref]:
                print(f'    {ref}: orig={orig_c[ref]} rt={rt_c[ref]}')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 6  —  Attribute quoting style
# ══════════════════════════════════════════════════════════════════════════════

def check6_attr_quoting(orig_bytes: bytes):
    banner('CHECK 6 — Attribute quoting style')

    rt_bytes = roundtrip_bytes(orig_bytes)

    orig_str = orig_bytes.decode('utf-8', errors='replace')
    rt_str   = rt_bytes.decode('utf-8', errors='replace')

    # Count single-quoted attributes in each
    sq_orig = len(re.findall(r"\w+='[^']*'", orig_str))
    sq_rt   = len(re.findall(r"\w+='[^']*'", rt_str))
    dq_orig = len(re.findall(r'\w+="[^"]*"', orig_str))
    dq_rt   = len(re.findall(r'\w+="[^"]*"', rt_str))

    print(f'  Double-quoted attrs: original={dq_orig}  roundtrip={dq_rt}')
    print(f'  Single-quoted attrs: original={sq_orig}  roundtrip={sq_rt}')

    if sq_orig == sq_rt and dq_orig == dq_rt:
        pass_fail(True, 'Attribute quoting unchanged')
    else:
        pass_fail(False, 'Attribute quoting CHANGED')
        if sq_orig != sq_rt:
            print(f'    Single-quoted: orig={sq_orig}  rt={sq_rt}')

    # Check XML declaration specifically
    decl_pattern = re.compile(r'<\?xml[^?]*\?>')
    orig_decl = decl_pattern.search(orig_str)
    rt_decl   = decl_pattern.search(rt_str)
    if orig_decl and rt_decl:
        print(f'\n  Original XML decl : {orig_decl.group()!r}')
        print(f'  Roundtrip XML decl: {rt_decl.group()!r}')
        if orig_decl.group() == rt_decl.group():
            pass_fail(True, 'XML declaration preserved exactly')
        else:
            pass_fail(False, 'XML declaration CHANGED')


# ══════════════════════════════════════════════════════════════════════════════
# CHECK 7  —  serialize_xml() output correctness
# ══════════════════════════════════════════════════════════════════════════════

def check7_serialize_xml_correctness():
    banner('CHECK 7 — serialize_xml() method correctness analysis')

    sys.path.insert(0, str(Path('/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/src')))
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            'hwpx_editor',
            '/mnt/c/business_forge/2026-digital-gyeongnam-rnd/hwpx-generator/src/hwpx_editor.py'
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        HwpxEditor = mod.HwpxEditor
    except Exception as e:
        print(f'  ERROR: Could not import hwpx_editor: {e}')
        return

    # Instantiate with the ORIGINAL file (no modifications)
    try:
        editor = HwpxEditor(str(ORIGINAL_PATH))
    except Exception as e:
        print(f'  ERROR: Could not open original HWPX: {e}')
        return

    serialized = editor.serialize_xml()

    print(f'  serialize_xml() output length: {len(serialized)} bytes')
    print(f'  Type: {type(serialized)}')

    # Verify it's valid UTF-8
    try:
        decoded = serialized.decode('utf-8')
        pass_fail(True, 'Valid UTF-8')
    except UnicodeDecodeError as e:
        pass_fail(False, f'NOT valid UTF-8: {e}')
        return

    # Check XML declaration
    if decoded.startswith('<?xml'):
        decl_end = decoded.index('?>') + 2
        decl = decoded[:decl_end]
        print(f'\n  XML declaration: {decl!r}')

        checks = {
            'Uses double quotes': '"1.0"' in decl and '"UTF-8"' in decl,
            'Has standalone="yes"': 'standalone="yes"' in decl,
            'No newline before root': decoded[decl_end] != '\n',
            'No newline after ?>': decoded[decl_end:decl_end+1] != '\n',
        }
        for label, ok in checks.items():
            pass_fail(ok, label)
            if label == 'No newline before root' and not ok:
                print(f'    Character after ?>: {decoded[decl_end:decl_end+5]!r}')
    else:
        pass_fail(False, 'Does not start with XML declaration')

    # Verify parseable by lxml
    try:
        from lxml import etree
        root2 = etree.fromstring(serialized)
        pass_fail(True, 'Parseable by lxml after serialization')
    except Exception as e:
        pass_fail(False, f'NOT parseable by lxml: {e}')

    # Compare byte length with original section0.xml
    orig_bytes = load_section(ORIGINAL_PATH)
    print(f'\n  Original section0.xml: {len(orig_bytes)} bytes')
    print(f'  Serialized output:     {len(serialized)} bytes')
    delta = len(serialized) - len(orig_bytes)
    print(f'  Delta: {delta:+d} bytes')

    # Check first 200 bytes
    print(f'\n  First 200 bytes of serialized:')
    print(f'  {serialized[:200]!r}')

    print(f'\n  First 200 bytes of original:')
    print(f'  {orig_bytes[:200]!r}')


# ══════════════════════════════════════════════════════════════════════════════
# BONUS  —  Spot-check key structural differences between original and filled
# ══════════════════════════════════════════════════════════════════════════════

def check_filled_vs_original():
    banner('BONUS — Structural comparison: original vs filled (form_pass1.hwpx)')

    if not FILLED_PATH.exists():
        print(f'  SKIP: filled file not found at {FILLED_PATH}')
        return

    orig_bytes   = load_section(ORIGINAL_PATH)
    filled_bytes = load_section(FILLED_PATH)

    orig_str   = orig_bytes.decode('utf-8', errors='replace')
    filled_str = filled_bytes.decode('utf-8', errors='replace')

    print(f'  Original section0.xml : {len(orig_bytes):>10} bytes')
    print(f'  Filled section0.xml   : {len(filled_bytes):>10} bytes')
    print(f'  Delta                 : {len(filled_bytes) - len(orig_bytes):>+10} bytes')

    # Count key elements
    for tag in ['hp:tbl', 'hp:tr', 'hp:tc', 'hp:p', 'hp:run', 'hp:t']:
        oc = orig_str.count(f'<{tag}')
        fc = filled_str.count(f'<{tag}')
        flag = '  ' if oc == fc else ' !'
        print(f'  {flag} <{tag}>: original={oc}  filled={fc}')

    # Check the XML declaration
    decl_re = re.compile(r'^<\?xml[^?]*\?>')
    orig_decl   = decl_re.match(orig_str)
    filled_decl = decl_re.match(filled_str)

    print()
    if orig_decl and filled_decl:
        od = orig_decl.group()
        fd = filled_decl.group()
        print(f'  Original XML decl: {od!r}')
        print(f'  Filled XML decl  : {fd!r}')
        pass_fail(od == fd, 'XML declarations match')

    # Check first character after XML declaration
    if orig_decl:
        c = orig_str[orig_decl.end()]
        print(f'\n  Original: char after XML decl = {c!r} (ord={ord(c)})')
    if filled_decl:
        c = filled_str[filled_decl.end()]
        print(f'  Filled:   char after XML decl = {c!r} (ord={ord(c)})')


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print('HWPX XML Serialization Corruption Diagnostic')
    print(f'Original : {ORIGINAL_PATH}')
    print(f'Filled   : {FILLED_PATH}')

    if not ORIGINAL_PATH.exists():
        print(f'ERROR: Original file not found: {ORIGINAL_PATH}')
        sys.exit(1)

    orig_bytes = load_section(ORIGINAL_PATH)
    print(f'Loaded section0.xml: {len(orig_bytes)} bytes')

    check1_roundtrip(orig_bytes)
    check2_namespace_order(orig_bytes)
    check3_whitespace(orig_bytes)
    check4_self_closing(orig_bytes)
    check5_entity_encoding(orig_bytes)
    check6_attr_quoting(orig_bytes)
    check7_serialize_xml_correctness()
    check_filled_vs_original()

    print()
    print('=' * 70)
    print('  Diagnostic complete.')
    print('=' * 70)


if __name__ == '__main__':
    main()
