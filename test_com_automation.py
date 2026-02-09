"""
테스트 2: COM 자동화를 사용한 HWPX 파일 생성
윈도우 Python + pywin32 + 한컴오피스 설치 필요
WSL에서 윈도우 Python으로 실행:
  /mnt/c/Users/lanco/AppData/Local/Programs/Python/Python312/python.exe test_com_automation.py
"""
import os
import sys

# 윈도우 환경 확인
if sys.platform != "win32":
    print("ERROR: 이 스크립트는 윈도우 Python에서 실행해야 합니다.")
    print("사용법: /mnt/c/.../python.exe test_com_automation.py")
    sys.exit(1)

import win32com.client as win32

def main():
    hwp = None
    try:
        # 한글 COM 객체 생성
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = True

        # 보안 모듈 등록 (필요 시)
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()

        pset.SetItem("Text", "테스트 문서 - COM 자동화로 생성")
        act.Execute(pset)

        # 줄바꿈
        hwp.HAction.Run("BreakPara")

        pset.SetItem("Text", "2026년 클라우드 종합솔루션 지원사업")
        act.Execute(pset)

        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")

        # 테이블 삽입 (3행 2열)
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        hwp.HParameterSet.HTableCreation.Rows = 3
        hwp.HParameterSet.HTableCreation.Cols = 2
        hwp.HParameterSet.HTableCreation.WidthType = 2  # 단에 맞춤
        hwp.HParameterSet.HTableCreation.HeightType = 0  # 자동
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

        # 테이블 셀에 데이터 입력
        cell_data = [
            ["항목", "내용"],
            ["사업명", "클라우드 솔루션"],
            ["연도", "2026"],
        ]

        for row_idx, row in enumerate(cell_data):
            for col_idx, text in enumerate(row):
                pset.SetItem("Text", text)
                act.Execute(pset)
                if col_idx < len(row) - 1:
                    hwp.HAction.Run("TableRightCell")
            if row_idx < len(cell_data) - 1:
                hwp.HAction.Run("TableRightCell")

        # 파일 저장
        output_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output_com.hwpx"
        )
        hwp.SaveAs(output_path, "HWPX", "")
        print(f"SUCCESS: {output_path} 생성 완료")

    except Exception as e:
        print(f"ERROR: {e}")
        raise
    finally:
        if hwp:
            hwp.Quit()


if __name__ == "__main__":
    main()
