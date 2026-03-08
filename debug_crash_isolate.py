"""Minimal crash isolation test.

Creates 3 HWPX test files from the original template, each adding exactly
one type of change, to identify which operation causes HWP to crash.

Test A: Re-save with NO modifications (pure ZIP round-trip).
Test B: Change exactly ONE cell (T0, row=6, col=3 = "테스트").
Test C: Inject ONE marker paragraph after table index 3.

Run from the hwpx-generator directory:
    python3 debug_crash_isolate.py

Then open each file in Hangul Office and observe which one crashes.
"""

import os
import sys
import zipfile

# Ensure src/ is importable
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'src'))

from hwpx_editor import HwpxEditor  # noqa: E402

TEMPLATE = '/mnt/c/business_forge/2026-digital-gyeongnam-rnd/form_to_fillout.hwpx'
OUT_DIR = os.path.join(SCRIPT_DIR, 'output', 'debug')

OUT_A = os.path.join(OUT_DIR, 'test_a_resave_only.hwpx')
OUT_B = os.path.join(OUT_DIR, 'test_b_one_cell.hwpx')
OUT_C = os.path.join(OUT_DIR, 'test_c_one_marker.hwpx')


def verify(path, label):
    """Open the ZIP and parse section0.xml. Print file size."""
    size = os.path.getsize(path)
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            names = zf.namelist()
            assert 'Contents/section0.xml' in names, "section0.xml missing from ZIP"
            xml_bytes = zf.read('Contents/section0.xml')

        from lxml import etree
        root = etree.fromstring(xml_bytes)
        assert root is not None, "etree.fromstring returned None"

        print(f"  [OK] {label}: {size:,} bytes, ZIP valid, section0.xml parses OK")
        print(f"       ZIP entries: {len(names)}")
    except Exception as exc:
        print(f"  [FAIL] {label}: {exc}")


# ---------------------------------------------------------------------------
# Test A — pure re-save, zero modifications
# ---------------------------------------------------------------------------
def make_test_a():
    print("\n=== Test A: re-save only (no modifications) ===")
    editor = HwpxEditor(TEMPLATE)
    editor.save(OUT_A)
    verify(OUT_A, 'Test A')


# ---------------------------------------------------------------------------
# Test B — change exactly one cell, no markers
# ---------------------------------------------------------------------------
def make_test_b():
    print("\n=== Test B: one cell change (T0 r6 c3 = '테스트') ===")
    editor = HwpxEditor(TEMPLATE)

    table_count = editor.get_table_count()
    print(f"  Template has {table_count} tables")

    t0 = editor.get_table(0)
    if t0 is None:
        print("  [FAIL] Table 0 not found")
        return

    # T0 r=3 c=4 is a blank writable cell ("기관명" row, value column).
    # Verified by inspecting the original template — this cell exists and is empty.
    cell = editor.get_cell(t0, row_addr=3, col_addr=4)
    if cell is None:
        print("  [WARN] T0 r3 c4 not found — trying r0 c0 instead")
        ok = editor.set_cell_text(t0, 0, 0, '테스트')
    else:
        ok = editor.set_cell_text(t0, 3, 4, '테스트')

    print(f"  set_cell_text(T0, r=3, c=4, '테스트') returned: {ok}")
    editor.save(OUT_B)
    verify(OUT_B, 'Test B')


# ---------------------------------------------------------------------------
# Test C — inject one marker after table 3, no cell changes
# ---------------------------------------------------------------------------
def make_test_c():
    print("\n=== Test C: one marker injection after table 3 ===")
    editor = HwpxEditor(TEMPLATE)

    table_count = editor.get_table_count()
    print(f"  Template has {table_count} tables")

    # Use table index 3; fall back to last table if fewer than 4 exist
    target_idx = min(3, table_count - 1)
    ok = editor.inject_marker(target_idx, '##DEBUG_MARKER##')
    print(f"  inject_marker(after_table={target_idx}) returned: {ok}")

    editor.save(OUT_C)
    verify(OUT_C, 'Test C')


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"Template : {TEMPLATE}")
    print(f"Output   : {OUT_DIR}")

    if not os.path.isfile(TEMPLATE):
        print(f"\n[ERROR] Template not found: {TEMPLATE}")
        sys.exit(1)

    make_test_a()
    make_test_b()
    make_test_c()

    print("\n=== Summary ===")
    for label, path in [('A', OUT_A), ('B', OUT_B), ('C', OUT_C)]:
        exists = os.path.isfile(path)
        size = os.path.getsize(path) if exists else 0
        status = f"{size:,} bytes" if exists else "MISSING"
        print(f"  Test {label}: {os.path.basename(path)} — {status}")

    print("\nOpen each file in Hangul Office and note which one crashes.")
    print("  A crashes  → the Python zipfile round-trip itself is broken")
    print("  B crashes  → cell text modification is the problem")
    print("  C crashes  → marker injection (paragraph insertion) is the problem")
    print("  None crash → the crash is in form_filler.py logic, not these primitives")
