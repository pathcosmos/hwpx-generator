"""
HWPX 자동 생성 메인 스크립트

전체 파이프라인:
1. JSON 입력 데이터 로드
2. 템플릿 HWPX 파일 복사
3. XML 셀 채우기 (빈 셀에 기업 정보 입력)
4. COM 모듈 호출하여 내용 삽입/수정
5. HWPX 및 PDF로 저장
6. PDF 비교 검증 실행 (선택)
7. 결과 리포트 출력

Usage:
    # 템플릿 기반 생성 (찾아바꾸기)
    python3 src/generate_hwpx.py --template ref/test_01.hwpx --data data/sample_input.json --output output/

    # 템플릿을 그대로 PDF로 변환 (검증용)
    python3 src/generate_hwpx.py --template ref/test_01.hwpx --output output/ --pdf-only

    # PDF 비교 검증 포함
    python3 src/generate_hwpx.py --template ref/test_01.hwpx --output output/ --pdf-only --compare ref/test_01.pdf
"""
import argparse
import json
import os
import shutil
import sys

# 프로젝트 루트를 path에 추가
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_DIR)

from src.bridge import (
    wsl_to_win_path,
    open_and_save_as_pdf,
    open_and_replace,
    fix_hwpx_for_pdf,
    WIN_PYTHON,
)


def load_input_data(data_path):
    """JSON 입력 데이터 로드"""
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_replacements(data):
    """입력 데이터에서 찾아바꾸기 매핑 생성

    커버 페이지의 텍스트를 기반으로 교체할 텍스트 쌍을 구성합니다.
    """
    replacements = {}

    # 사업명
    if "사업명" in data:
        replacements["클라우드 종합솔루션 지원사업(통합형 클라우드화)"] = data["사업명"]

    # 과제명
    if "과제명" in data:
        replacements[
            "웹 기반 설계자산(도면-문서) 통합 관리와 상태기반(CBM) 설비보전 기술이 결합된 LLM 융합 중소제조기업형 SaaS 플랫폼"
        ] = data["과제명"]

    # 사업개요
    if "사업개요" in data:
        replacements["과제내용에 대하여 간단히 요약 기술"] = data["사업개요"]

    # 수행기간
    if "수행기간" in data:
        period = data["수행기간"]
        if "개발시작" in period and "개발종료" in period:
            dev_period = period.get("개발기간", "12개월")
            old_dev = " · 솔루션 개발 : '26.6.30 ~ '27.6.30  (12개월)"
            new_dev = f" · 솔루션 개발 : {period['개발시작']} ~ {period['개발종료']}  ({dev_period})"
            replacements[old_dev] = new_dev
        if "실증시작" in period and "실증종료" in period:
            test_period = period.get("실증기간", "6개월")
            old_test = " · 솔루션 실증 : '27.6.30 ~ '27.12.31 (6개월)"
            new_test = f" · 솔루션 실증 : {period['실증시작']} ~ {period['실증종료']} ({test_period})"
            replacements[old_test] = new_test

    # 대표공급기업 정보
    if "대표공급기업" in data:
        company = data["대표공급기업"]
        # 필드가 비어있는 경우에 해당하는 텍스트는 찾아바꾸기가 어려우므로
        # 기업명 등 실제 텍스트가 있는 경우만 처리

    return replacements


def generate_from_template(template_path, data_path, output_dir,
                           generate_pdf=True, compare_pdf=None):
    """템플릿 기반 HWPX 생성

    Args:
        template_path: 템플릿 HWPX 파일 경로
        data_path: JSON 입력 데이터 경로 (None이면 수정 없이 복사)
        output_dir: 출력 디렉토리
        generate_pdf: PDF도 생성할지 여부
        compare_pdf: 비교할 참조 PDF 경로 (None이면 비교 안함)
    """
    os.makedirs(output_dir, exist_ok=True)

    output_hwpx = os.path.join(output_dir, "generated.hwpx")
    output_pdf = os.path.join(output_dir, "generated.pdf") if generate_pdf else None

    # 데이터 없으면 원본을 그대로 PDF로 변환
    if data_path is None:
        print(f"[1/3] 템플릿을 PDF로 변환 중...")
        print(f"      템플릿: {template_path}")

        # 템플릿 복사 및 인쇄설정 수정
        shutil.copy2(template_path, output_hwpx)
        fix_hwpx_for_pdf(output_hwpx)
        print(f"      HWPX 복사 및 인쇄설정 수정: {output_hwpx}")

        if generate_pdf:
            success = open_and_save_as_pdf(output_hwpx, output_pdf, timeout=300)
            if success:
                size = os.path.getsize(output_pdf)
                print(f"      PDF 생성 성공: {output_pdf} ({size:,} bytes)")
            else:
                print(f"      PDF 생성 실패!")
                return False
    else:
        # 데이터 로드
        print(f"[1/4] 입력 데이터 로드 중...")
        data = load_input_data(data_path)

        # [STEP 2] XML 셀 채우기 (빈 셀에 기업 정보 입력)
        print(f"[2/4] XML 셀 채우기 중...")
        shutil.copy2(template_path, output_hwpx)
        xml_filled = 0
        try:
            from src.field_mapper import load_field_map, build_cell_data
            from src.hwpx_editor import HwpxEditor

            template_dir = os.path.join(PROJECT_DIR, "templates", "cloud_integrated")
            field_map = load_field_map(template_dir)
            cell_data = build_cell_data(data, field_map)
            if cell_data:
                editor = HwpxEditor(output_hwpx)
                table = editor.get_table(0)
                if table is not None:
                    xml_filled = editor.fill_cells(table, cell_data)
                    editor.save()
                    print(f"      XML 셀 채우기: {xml_filled}/{len(cell_data)}개 셀 수정")
                else:
                    print(f"      경고: 커버 테이블을 찾을 수 없습니다")
            else:
                print(f"      채울 셀 데이터 없음")
        except Exception as e:
            print(f"      XML 셀 채우기 실패: {e}")
            # XML 채우기 실패 시 원본 템플릿으로 복원
            shutil.copy2(template_path, output_hwpx)

        # [STEP 3] COM find-and-replace (사업명, 과제명 등 텍스트 교체)
        replacements = build_replacements(data)
        print(f"[3/4] COM 텍스트 교체 중... ({len(replacements)}개 항목)")

        # PrintMethod=0 적용 (COM이 올바른 인쇄설정으로 PDF 생성하도록)
        fix_hwpx_for_pdf(output_hwpx)

        if replacements:
            print(f"      템플릿: {output_hwpx}")
            success = open_and_replace(
                output_hwpx, replacements, output_hwpx, output_pdf, timeout=300
            )
            if success:
                fix_hwpx_for_pdf(output_hwpx)
                hwpx_size = os.path.getsize(output_hwpx)
                print(f"      HWPX 생성: {output_hwpx} ({hwpx_size:,} bytes)")
                if output_pdf and os.path.exists(output_pdf):
                    pdf_size = os.path.getsize(output_pdf)
                    print(f"      PDF 생성: {output_pdf} ({pdf_size:,} bytes)")
            else:
                print(f"      문서 생성 실패!")
                return False
        else:
            print(f"      교체할 내용 없음")
            fix_hwpx_for_pdf(output_hwpx)
            if generate_pdf:
                success = open_and_save_as_pdf(output_hwpx, output_pdf, timeout=300)
                if not success:
                    print(f"      PDF 생성 실패!")
                    return False

    # PDF 비교
    if compare_pdf and output_pdf and os.path.exists(output_pdf):
        print(f"[4/4] PDF 비교 검증 중...")
        compare_dir = os.path.join(output_dir, "compare_report")
        try:
            from src.pdf_compare import PdfComparator
            comparator = PdfComparator(compare_pdf, output_pdf, dpi=150)
            result = comparator.compare(output_dir=compare_dir)

            print(f"      참조 페이지: {result['reference_pages']}")
            print(f"      생성 페이지: {result['generated_pages']}")
            print(f"      전체 SSIM: {result['overall_ssim']:.4f}")
            print(f"      결과: {'PASS' if result['pass'] else 'FAIL'}")

            if not result['pass']:
                print(f"      차이점 리포트: {compare_dir}/")
        except ImportError as e:
            print(f"      PDF 비교 모듈 로드 실패: {e}")
            print(f"      pip3 install --break-system-packages pymupdf Pillow scikit-image")
        except Exception as e:
            print(f"      PDF 비교 중 오류: {e}")
    else:
        print(f"[4/4] PDF 비교 건너뜀")

    print(f"\n완료! 출력 디렉토리: {output_dir}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="HWPX 자동 생성 프로그램",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  # 템플릿을 PDF로 변환 (검증용)
  python3 src/generate_hwpx.py --template ref/test_01.hwpx --output output/ --pdf-only

  # 데이터로 문서 생성
  python3 src/generate_hwpx.py --template ref/test_01.hwpx --data data/sample_input.json --output output/

  # PDF 비교 검증 포함
  python3 src/generate_hwpx.py --template ref/test_01.hwpx --output output/ --pdf-only --compare ref/test_01.pdf
        """
    )
    parser.add_argument("--template", "-t", required=True,
                        help="템플릿 HWPX 파일 경로")
    parser.add_argument("--data", "-d", default=None,
                        help="JSON 입력 데이터 파일 경로")
    parser.add_argument("--output", "-o", default="output",
                        help="출력 디렉토리 (기본: output)")
    parser.add_argument("--pdf-only", action="store_true",
                        help="데이터 없이 템플릿을 그대로 PDF로 변환")
    parser.add_argument("--no-pdf", action="store_true",
                        help="PDF 생성 건너뛰기")
    parser.add_argument("--compare", "-c", default=None,
                        help="비교할 참조 PDF 경로")

    args = parser.parse_args()

    if not os.path.exists(args.template):
        print(f"오류: 템플릿 파일을 찾을 수 없습니다: {args.template}")
        sys.exit(1)

    if args.data and not os.path.exists(args.data):
        print(f"오류: 데이터 파일을 찾을 수 없습니다: {args.data}")
        sys.exit(1)

    data_path = None if args.pdf_only else args.data
    generate_pdf = not args.no_pdf

    success = generate_from_template(
        template_path=args.template,
        data_path=data_path,
        output_dir=args.output,
        generate_pdf=generate_pdf,
        compare_pdf=args.compare,
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
