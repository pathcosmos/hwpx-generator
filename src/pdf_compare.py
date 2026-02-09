"""PDF 비교 검증 모듈

두 PDF 파일을 페이지별로 시각적/텍스트 비교하여 유사도를 측정합니다.
- SSIM (구조적 유사도)
- 픽셀 차이 비율
- 텍스트 내용 일치도
"""
import fitz  # PyMuPDF
from PIL import Image
import numpy as np
from skimage.metrics import structural_similarity as ssim
import os
import json
import io
import difflib


class PdfComparator:
    """두 PDF 파일을 페이지별로 비교"""

    def __init__(self, reference_pdf, generated_pdf, dpi=150):
        """
        Args:
            reference_pdf: 참조(원본) PDF 경로
            generated_pdf: 생성된 PDF 경로
            dpi: 비교용 이미지 해상도
        """
        self.reference_pdf = reference_pdf
        self.generated_pdf = generated_pdf
        self.dpi = dpi
        self._ref_doc = None
        self._gen_doc = None

    def _open_docs(self):
        if self._ref_doc is None:
            self._ref_doc = fitz.open(self.reference_pdf)
        if self._gen_doc is None:
            self._gen_doc = fitz.open(self.generated_pdf)

    def _close_docs(self):
        if self._ref_doc is not None:
            self._ref_doc.close()
            self._ref_doc = None
        if self._gen_doc is not None:
            self._gen_doc.close()
            self._gen_doc = None

    def compare(self, output_dir=None, pages=None, threshold=0.90):
        """전체 비교 수행

        Args:
            output_dir: 리포트 출력 디렉토리 (None이면 리포트 생성 안 함)
            pages: 비교할 페이지 범위 (1-based), 예: (1, 5) -> 1~5페이지
            threshold: PASS 기준 SSIM 값

        Returns:
            dict: 비교 결과
        """
        self._open_docs()
        try:
            ref_pages = len(self._ref_doc)
            gen_pages = len(self._gen_doc)

            if pages:
                start, end = pages
                start = max(1, start)
                end = min(end, ref_pages, gen_pages)
                page_range = range(start, end + 1)
            else:
                comparable = min(ref_pages, gen_pages)
                page_range = range(1, comparable + 1)

            page_results = []
            ssim_values = []
            text_matches = []

            for page_num in page_range:
                print(f"  비교 중: 페이지 {page_num}/{page_range[-1]}...", end="\r")
                page_result = self.compare_page(page_num)
                page_results.append(page_result)
                ssim_values.append(page_result["ssim"])
                text_matches.append(1.0 if page_result["text_match"] else page_result.get("text_similarity", 0.0))

            overall_ssim = float(np.mean(ssim_values)) if ssim_values else 0.0
            overall_text_match = float(np.mean(text_matches)) if text_matches else 0.0

            result = {
                "reference_pdf": os.path.abspath(self.reference_pdf),
                "generated_pdf": os.path.abspath(self.generated_pdf),
                "reference_pages": ref_pages,
                "generated_pages": gen_pages,
                "page_count_match": ref_pages == gen_pages,
                "compared_pages": len(page_results),
                "pages": page_results,
                "overall_ssim": round(overall_ssim, 4),
                "overall_text_match": round(overall_text_match, 4),
                "threshold": threshold,
                "pass": overall_ssim >= threshold,
            }

            print()  # newline after \r progress

            if output_dir:
                self.generate_report(result, output_dir, page_range)

            return result
        finally:
            self._close_docs()

    def compare_page(self, page_num):
        """단일 페이지 비교 (1-based page number)

        Returns:
            dict: 페이지 비교 결과
        """
        self._open_docs()

        # 이미지 비교
        img_ref = self.page_to_image(self._ref_doc, page_num)
        img_gen = self.page_to_image(self._gen_doc, page_num)

        # 크기 맞추기 (더 큰 쪽에 맞춤)
        img_ref, img_gen = self._match_sizes(img_ref, img_gen)

        ssim_val = self.compute_ssim(img_ref, img_gen)
        pixel_diff = self.compute_pixel_diff(img_ref, img_gen)

        # 텍스트 비교
        text_result = self.compare_text(page_num)

        return {
            "page": page_num,
            "ssim": round(ssim_val, 4),
            "pixel_diff_percent": round(pixel_diff, 2),
            "text_match": text_result["match"],
            "text_similarity": round(text_result["similarity"], 4),
        }

    def page_to_image(self, doc, page_num, auto_rotate=True):
        """PDF 페이지를 PIL Image (grayscale)로 변환

        Args:
            doc: fitz.Document 객체
            page_num: 1-based 페이지 번호
            auto_rotate: True면 landscape 페이지를 portrait로 자동 회전
        """
        page = doc.load_page(page_num - 1)  # 0-based
        zoom = self.dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        img = Image.frombytes("L", [pix.width, pix.height], pix.samples)

        # landscape 페이지를 portrait로 자동 회전
        if auto_rotate and img.width > img.height:
            img = img.rotate(90, expand=True)

        return img

    def _match_sizes(self, img1, img2):
        """두 이미지 크기를 맞춤 (더 큰 쪽 기준으로 리사이즈)"""
        if img1.size == img2.size:
            return img1, img2

        # 더 큰 쪽에 맞춤 (해상도 보존)
        w = max(img1.width, img2.width)
        h = max(img1.height, img2.height)
        if img1.size != (w, h):
            img1 = img1.resize((w, h), Image.LANCZOS)
        if img2.size != (w, h):
            img2 = img2.resize((w, h), Image.LANCZOS)
        return img1, img2

    def compute_ssim(self, img1, img2):
        """두 이미지의 SSIM (구조적 유사도) 계산

        Returns:
            float: 0.0 ~ 1.0 사이 값 (1.0 = 동일)
        """
        arr1 = np.array(img1)
        arr2 = np.array(img2)
        score, _ = ssim(arr1, arr2, full=True)
        return float(score)

    def compute_pixel_diff(self, img1, img2):
        """픽셀 차이 비율 계산

        Returns:
            float: 차이 비율 (0.0% ~ 100.0%)
        """
        arr1 = np.array(img1, dtype=np.float32)
        arr2 = np.array(img2, dtype=np.float32)
        diff = np.abs(arr1 - arr2)
        # 임계값 10 이상의 차이만 "다른 픽셀"로 카운트
        diff_pixels = np.sum(diff > 10)
        total_pixels = arr1.size
        return float(diff_pixels / total_pixels * 100)

    def compare_text(self, page_num):
        """페이지의 텍스트 내용 비교

        Args:
            page_num: 1-based 페이지 번호

        Returns:
            dict: {"match": bool, "similarity": float, "ref_text": str, "gen_text": str}
        """
        self._open_docs()

        ref_text = self._ref_doc.load_page(page_num - 1).get_text().strip()
        gen_text = self._gen_doc.load_page(page_num - 1).get_text().strip()

        if ref_text == gen_text:
            return {"match": True, "similarity": 1.0, "ref_text": ref_text, "gen_text": gen_text}

        similarity = difflib.SequenceMatcher(None, ref_text, gen_text).ratio()
        return {"match": False, "similarity": similarity, "ref_text": ref_text, "gen_text": gen_text}

    def generate_diff_image(self, page_num, output_path):
        """차이점을 시각화한 이미지 생성

        빨간색으로 차이 영역을 표시한 이미지를 저장합니다.
        """
        self._open_docs()

        img_ref = self.page_to_image(self._ref_doc, page_num)
        img_gen = self.page_to_image(self._gen_doc, page_num)
        img_ref, img_gen = self._match_sizes(img_ref, img_gen)

        arr_ref = np.array(img_ref, dtype=np.float32)
        arr_gen = np.array(img_gen, dtype=np.float32)
        diff = np.abs(arr_ref - arr_gen)

        # 참조 이미지를 RGB로 변환하고 차이 부분을 빨간색으로 표시
        rgb_ref = np.stack([arr_ref] * 3, axis=-1).astype(np.uint8)

        mask = diff > 10
        rgb_ref[mask, 0] = 255  # R
        rgb_ref[mask, 1] = 0    # G
        rgb_ref[mask, 2] = 0    # B

        diff_img = Image.fromarray(rgb_ref, "RGB")
        diff_img.save(output_path)
        return output_path

    def generate_report(self, result, output_dir, page_range=None):
        """비교 결과 리포트 생성

        - report.json: 상세 비교 결과
        - summary.txt: 요약
        - diff_page_N.png: 차이점 이미지 (SSIM < 0.95인 페이지만)
        """
        os.makedirs(output_dir, exist_ok=True)

        # report.json (텍스트 필드 제외)
        report_data = dict(result)
        for p in report_data.get("pages", []):
            p.pop("ref_text", None)
            p.pop("gen_text", None)

        report_path = os.path.join(output_dir, "report.json")
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)

        # summary.txt
        summary_path = os.path.join(output_dir, "summary.txt")
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("PDF 비교 검증 결과 요약\n")
            f.write("=" * 60 + "\n\n")
            f.write(f"참조 PDF: {result['reference_pdf']}\n")
            f.write(f"생성 PDF: {result['generated_pdf']}\n\n")
            f.write(f"참조 페이지 수: {result['reference_pages']}\n")
            f.write(f"생성 페이지 수: {result['generated_pages']}\n")
            f.write(f"페이지 수 일치: {'예' if result['page_count_match'] else '아니오'}\n")
            f.write(f"비교한 페이지 수: {result['compared_pages']}\n\n")
            f.write(f"전체 SSIM 평균: {result['overall_ssim']:.4f}\n")
            f.write(f"전체 텍스트 일치도: {result['overall_text_match']:.4f}\n")
            f.write(f"PASS 기준: {result['threshold']}\n")
            f.write(f"결과: {'PASS' if result['pass'] else 'FAIL'}\n\n")

            f.write("-" * 60 + "\n")
            f.write(f"{'페이지':>6} | {'SSIM':>8} | {'픽셀차이%':>8} | {'텍스트일치':>8} | {'텍스트유사도':>10}\n")
            f.write("-" * 60 + "\n")
            for p in result["pages"]:
                f.write(
                    f"{p['page']:>6} | {p['ssim']:>8.4f} | {p['pixel_diff_percent']:>8.2f} | "
                    f"{'일치' if p['text_match'] else '불일치':>8} | {p['text_similarity']:>10.4f}\n"
                )
            f.write("-" * 60 + "\n")

        # 차이점 이미지 생성 (SSIM < 0.95인 페이지)
        self._open_docs()
        diff_count = 0
        for p in result["pages"]:
            if p["ssim"] < 0.95:
                diff_path = os.path.join(output_dir, f"diff_page_{p['page']:03d}.png")
                try:
                    self.generate_diff_image(p["page"], diff_path)
                    diff_count += 1
                except Exception as e:
                    print(f"  경고: 페이지 {p['page']} diff 이미지 생성 실패: {e}")

        print(f"리포트 생성 완료: {output_dir}")
        print(f"  - report.json, summary.txt")
        if diff_count > 0:
            print(f"  - diff 이미지 {diff_count}개")


def parse_pages(pages_str):
    """페이지 범위 문자열 파싱. '1-5' -> (1, 5)"""
    if not pages_str:
        return None
    parts = pages_str.split("-")
    if len(parts) == 1:
        n = int(parts[0])
        return (n, n)
    return (int(parts[0]), int(parts[1]))


def main():
    """CLI 인터페이스"""
    import argparse

    parser = argparse.ArgumentParser(description="PDF 비교 검증")
    parser.add_argument("reference", help="참조 PDF 경로")
    parser.add_argument("generated", help="생성된 PDF 경로")
    parser.add_argument("--output", "-o", default="output/compare_report", help="리포트 출력 디렉토리")
    parser.add_argument("--dpi", type=int, default=150, help="비교 해상도 (기본: 150)")
    parser.add_argument("--threshold", type=float, default=0.90, help="PASS 기준 SSIM (기본: 0.90)")
    parser.add_argument("--pages", help="비교할 페이지 범위 (예: 1-5)")
    args = parser.parse_args()

    pages = parse_pages(args.pages)

    print(f"PDF 비교 시작")
    print(f"  참조: {args.reference}")
    print(f"  생성: {args.generated}")
    if pages:
        print(f"  페이지: {pages[0]}~{pages[1]}")
    print()

    comparator = PdfComparator(args.reference, args.generated, dpi=args.dpi)
    result = comparator.compare(output_dir=args.output, pages=pages, threshold=args.threshold)

    print()
    print(f"전체 SSIM: {result['overall_ssim']:.4f}")
    print(f"전체 텍스트 일치도: {result['overall_text_match']:.4f}")
    print(f"결과: {'PASS' if result['pass'] else 'FAIL'} (기준: {args.threshold})")


if __name__ == "__main__":
    main()
