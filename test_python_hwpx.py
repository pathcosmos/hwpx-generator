"""
테스트 1: python-hwpx 라이브러리를 사용한 HWPX 파일 생성
WSL 환경에서 직접 실행 가능
"""
from hwpx.document import HwpxDocument


def main():
    # 새 빈 문서 생성
    doc = HwpxDocument.new()
    section = doc.sections[0]

    # 텍스트 추가
    para1 = doc.add_paragraph("테스트 문서 - python-hwpx로 생성", section=section)
    para2 = doc.add_paragraph("2026년 클라우드 종합솔루션 지원사업", section=section)
    doc.add_paragraph("", section=section)  # 빈 줄

    # 테이블 추가
    table = doc.add_table(rows=3, cols=2, section=section)

    # 헤더 행
    table.cell(0, 0).text = "항목"
    table.cell(0, 1).text = "내용"

    # 데이터 행
    table.cell(1, 0).text = "사업명"
    table.cell(1, 1).text = "클라우드 솔루션"
    table.cell(2, 0).text = "연도"
    table.cell(2, 1).text = "2026"

    # 파일 저장
    output_path = "output_python_hwpx.hwpx"
    doc.save(output_path)
    print(f"SUCCESS: {output_path} 생성 완료")

    # 생성된 파일 검증 - ZIP 구조 확인
    import zipfile
    with zipfile.ZipFile(output_path, "r") as z:
        print("\n=== 생성된 HWPX 파일 구조 ===")
        for name in z.namelist():
            info = z.getinfo(name)
            print(f"  {name} ({info.file_size} bytes)")

        # section0.xml 내용 확인
        print("\n=== section0.xml 내용 (일부) ===")
        for part_name in z.namelist():
            if "section" in part_name.lower():
                content = z.read(part_name).decode("utf-8")
                # 처음 2000자만 표시
                print(content[:2000])
                break


if __name__ == "__main__":
    main()
