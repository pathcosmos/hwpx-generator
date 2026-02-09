"""
통합 테스트: 템플릿 열기 → PDF 저장
Windows Python으로 실행
"""
import sys
import os
import time

if sys.platform != "win32":
    print("ERROR: This script must run under Windows Python.")
    sys.exit(1)

# 프로젝트 루트
project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_dir)

from src.hwp_com import HwpController

def main():
    template = os.path.join(project_dir, "ref", "test_01.hwpx")
    output_dir = os.path.join(project_dir, "output", "integration_test")
    os.makedirs(output_dir, exist_ok=True)

    output_hwpx = os.path.join(output_dir, "generated.hwpx")
    output_pdf = os.path.join(output_dir, "generated.pdf")

    print(f"[1] Opening template: {template}")
    start = time.time()

    with HwpController(visible=False) as hwp:
        hwp.open(template)
        pages = hwp.get_page_count()
        print(f"    Pages: {pages} (loaded in {time.time()-start:.1f}s)")

        # Save as HWPX (copy)
        print(f"[2] Saving HWPX: {output_hwpx}")
        hwp.save_as(output_hwpx, "HWPX")
        print(f"    Size: {os.path.getsize(output_hwpx):,} bytes")

        # Save as PDF
        print(f"[3] Saving PDF: {output_pdf}")
        t = time.time()
        hwp.save_as_pdf(output_pdf)
        print(f"    Size: {os.path.getsize(output_pdf):,} bytes ({time.time()-t:.1f}s)")

    print(f"\n[OK] Total time: {time.time()-start:.1f}s")
    print(f"    HWPX: {output_hwpx}")
    print(f"    PDF:  {output_pdf}")

if __name__ == "__main__":
    main()
