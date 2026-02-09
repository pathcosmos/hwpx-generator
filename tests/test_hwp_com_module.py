"""
Test the HwpController module via Windows Python.
Run from WSL: /mnt/c/Users/lanco/AppData/Local/Microsoft/WindowsApps/python.exe tests/test_hwp_com_module.py
Or via bridge: python3 -c "from src.bridge import run_com_script; r = run_com_script('tests/test_hwp_com_module.py'); print(r.stdout); print(r.stderr)"
"""
import os
import sys
import time

if sys.platform != "win32":
    print("ERROR: This script must run under Windows Python.")
    sys.exit(1)

# Add project root to path
project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_dir)

from src.hwp_com import HwpController


def test_create_and_save():
    """Test 1: Create new document with text and table, save as HWPX and PDF."""
    print("=== Test 1: Create document with text + table ===", flush=True)
    output_hwpx = os.path.join(project_dir, "tests", "output_module_create.hwpx")
    output_pdf = os.path.join(project_dir, "tests", "output_module_create.pdf")

    with HwpController(visible=False) as hwp:
        # Title
        hwp.set_char_shape(font="맑은 고딕", size=18, bold=True, color=0x000080)
        hwp.set_para_shape(align="center")
        hwp.insert_text("COM 모듈 테스트 문서")
        hwp.insert_line_break()

        # Body
        hwp.set_char_shape(font="맑은 고딕", size=10, bold=False, color=0x000000)
        hwp.set_para_shape(align="justify", line_spacing=160)
        hwp.insert_text("이 문서는 HwpController 모듈의 기능을 검증하기 위해 자동 생성되었습니다.")
        hwp.insert_line_break()
        hwp.insert_line_break()

        # Table
        hwp.insert_table(3, 3)
        hwp.fill_table([
            ["항목", "내용", "비고"],
            ["모듈명", "HwpController", "src/hwp_com.py"],
            ["테스트", "통과", "자동생성"],
        ])

        # Save
        hwp.save_as(output_hwpx, "HWPX")
        hwp.save_as_pdf(output_pdf)

    hwpx_ok = os.path.exists(output_hwpx) and os.path.getsize(output_hwpx) > 0
    pdf_ok = os.path.exists(output_pdf) and os.path.getsize(output_pdf) > 0
    print(f"  HWPX: {'PASS' if hwpx_ok else 'FAIL'} ({os.path.getsize(output_hwpx):,} bytes)" if hwpx_ok else "  HWPX: FAIL", flush=True)
    print(f"  PDF:  {'PASS' if pdf_ok else 'FAIL'} ({os.path.getsize(output_pdf):,} bytes)" if pdf_ok else "  PDF:  FAIL", flush=True)
    return hwpx_ok and pdf_ok


def test_open_and_read():
    """Test 2: Open existing file, read text and controls."""
    print("\n=== Test 2: Open existing file and read ===", flush=True)
    hwpx_path = os.path.join(project_dir, "ref", "test_01.hwpx")
    if not os.path.exists(hwpx_path):
        print("  SKIP: ref/test_01.hwpx not found", flush=True)
        return True

    with HwpController(visible=False) as hwp:
        hwp.open(hwpx_path)
        text = hwp.get_text()
        page_count = hwp.get_page_count()
        controls = hwp.get_controls()

    text_ok = len(text) > 0
    print(f"  Text length: {len(text)} chars ({'PASS' if text_ok else 'FAIL'})", flush=True)
    print(f"  Page count: {page_count}", flush=True)
    print(f"  Controls: {controls}", flush=True)
    return text_ok


def test_find_and_replace():
    """Test 3: Find and replace in a document."""
    print("\n=== Test 3: Find and replace ===", flush=True)
    output_path = os.path.join(project_dir, "tests", "output_module_replace.hwpx")

    with HwpController(visible=False) as hwp:
        hwp.insert_text("Hello {{NAME}}, welcome to {{COMPANY}}.")
        hwp.insert_line_break()
        hwp.insert_text("Your role is {{ROLE}}.")

        hwp.find_and_replace_all({
            "{{NAME}}": "홍길동",
            "{{COMPANY}}": "한컴오피스",
            "{{ROLE}}": "개발자",
        })

        text = hwp.get_text()
        hwp.save_as(output_path, "HWPX")

    has_name = "홍길동" in text
    no_placeholder = "{{NAME}}" not in text
    ok = has_name and no_placeholder
    print(f"  Replaced text: {text[:100].strip()}", flush=True)
    print(f"  Result: {'PASS' if ok else 'FAIL'}", flush=True)
    return ok


def test_open_save_pdf():
    """Test 4: Open existing HWPX and save as PDF."""
    print("\n=== Test 4: Open existing file -> PDF ===", flush=True)
    hwpx_path = os.path.join(project_dir, "ref", "test_01.hwpx")
    pdf_path = os.path.join(project_dir, "tests", "output_module_existing_to_pdf.pdf")

    if not os.path.exists(hwpx_path):
        print("  SKIP: ref/test_01.hwpx not found", flush=True)
        return True

    with HwpController(visible=False) as hwp:
        hwp.open(hwpx_path)
        hwp.save_as_pdf(pdf_path)

    ok = os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0
    print(f"  PDF: {'PASS' if ok else 'FAIL'} ({os.path.getsize(pdf_path):,} bytes)" if ok else "  PDF: FAIL", flush=True)
    return ok


def main():
    print("=" * 60, flush=True)
    print("HwpController Module Test Suite", flush=True)
    print("=" * 60, flush=True)

    tests = [
        ("Create & Save", test_create_and_save),
        ("Open & Read", test_open_and_read),
        ("Find & Replace", test_find_and_replace),
        ("Open -> PDF", test_open_save_pdf),
    ]

    results = []
    for name, func in tests:
        try:
            ok = func()
        except Exception as e:
            print(f"  EXCEPTION: {e}", flush=True)
            ok = False
        results.append((name, ok))
        # Wait for Hangul process to fully exit before next test
        time.sleep(5)

    print("\n" + "=" * 60, flush=True)
    print("Results:", flush=True)
    all_pass = True
    for name, ok in results:
        status = "PASS" if ok else "FAIL"
        print(f"  {name}: {status}", flush=True)
        if not ok:
            all_pass = False

    print(f"\nOverall: {'ALL PASSED' if all_pass else 'SOME FAILED'}", flush=True)
    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
