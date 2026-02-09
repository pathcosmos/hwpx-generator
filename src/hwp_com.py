"""
한컴오피스 COM 자동화 모듈
Windows Python에서 실행 (pywin32 필요)

Usage:
    with HwpController(visible=False) as hwp:
        hwp.insert_text("Hello")
        hwp.save_as("output.hwpx")
"""
import sys
import os
import time

if sys.platform != "win32":
    raise RuntimeError("This module requires Windows Python with pywin32.")

import win32com.client as win32
import pythoncom


# Paragraph alignment constants
ALIGN_JUSTIFY = 0
ALIGN_LEFT = 1
ALIGN_RIGHT = 2
ALIGN_CENTER = 3

# Save format strings recognized by Hangul COM
FORMAT_HWP = "HWP"
FORMAT_HWPX = "HWPX"
FORMAT_PDF = "PDF"
FORMAT_HTML = "HTML"
FORMAT_TEXT = "TEXT"


class HwpController:
    """한글 워드프로세서 COM 제어 클래스"""

    def __init__(self, visible=False, retries=3, retry_delay=5):
        """한글 COM 객체 생성 및 초기화

        Args:
            visible: True이면 한글 창을 표시, False이면 백그라운드 실행
            retries: COM 연결 실패 시 재시도 횟수
            retry_delay: 재시도 간 대기 시간 (초)
        """
        last_error = None
        for attempt in range(retries):
            try:
                self._hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
                break
            except pythoncom.com_error as e:
                last_error = e
                if attempt < retries - 1:
                    # Try with plain Dispatch as fallback
                    try:
                        self._hwp = win32.Dispatch("HWPFrame.HwpObject")
                        break
                    except pythoncom.com_error:
                        time.sleep(retry_delay)
                        continue
        else:
            raise RuntimeError(
                f"Failed to create HWP COM object after {retries} attempts: {last_error}"
            )
        self._hwp.XHwpWindows.Item(0).Visible = visible
        self._hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

    @property
    def hwp(self):
        """내부 COM 객체 직접 접근 (고급 사용)"""
        return self._hwp

    # --- File operations ---

    def open(self, filepath):
        """HWP/HWPX 파일 열기

        Args:
            filepath: 열 파일의 절대 경로 (Windows 경로)

        Returns:
            bool: 열기 성공 여부
        """
        filepath = os.path.abspath(filepath)
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

        ext = os.path.splitext(filepath)[1].lower()
        fmt = "HWPX" if ext == ".hwpx" else "HWP"
        return self._hwp.Open(filepath, fmt, "")

    def save(self, filepath=None):
        """현재 문서 저장

        Args:
            filepath: 저장 경로 (None이면 현재 파일에 덮어쓰기)
        """
        if filepath:
            filepath = os.path.abspath(filepath)
            ext = os.path.splitext(filepath)[1].lower()
            fmt = "HWPX" if ext == ".hwpx" else "HWP"
            self._hwp.SaveAs(filepath, fmt, "")
        else:
            self._hwp.Save()

    def save_as(self, filepath, format="HWPX"):
        """다른 형식으로 저장

        Args:
            filepath: 저장 경로 (절대 경로)
            format: 저장 형식 ("HWPX", "HWP", "PDF", "HTML", "TEXT")
        """
        filepath = os.path.abspath(filepath)
        self._hwp.SaveAs(filepath, format, "")

    def save_as_pdf(self, filepath):
        """PDF로 저장 (편의 메서드)

        Args:
            filepath: PDF 파일 저장 경로
        """
        self.save_as(filepath, FORMAT_PDF)

    def close(self):
        """현재 문서 닫기 (저장 안함)"""
        self._hwp.Clear(1)

    def quit(self):
        """한글 프로그램 종료"""
        try:
            self._hwp.Clear(1)
        except Exception:
            pass
        try:
            self._hwp.Quit()
        except Exception:
            pass

    # --- Text reading ---

    def get_text(self):
        """전체 텍스트 추출

        Returns:
            str: 문서의 전체 텍스트
        """
        self._hwp.HAction.Run("MoveDocBegin")
        self._hwp.HAction.Run("SelectAll")
        text = self._hwp.GetTextFile("TEXT", "")
        self._hwp.HAction.Run("MoveDocEnd")
        return text or ""

    def get_page_count(self):
        """페이지 수 반환

        Returns:
            int: 페이지 수
        """
        return self._hwp.PageCount

    # --- Cursor movement ---

    def move_to_start(self):
        """문서 시작으로 이동"""
        self._hwp.HAction.Run("MoveDocBegin")

    def move_to_end(self):
        """문서 끝으로 이동"""
        self._hwp.HAction.Run("MoveDocEnd")

    # --- Text insertion ---

    def insert_text(self, text):
        """현재 위치에 텍스트 삽입

        Args:
            text: 삽입할 텍스트
        """
        act = self._hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

    def insert_line_break(self):
        """줄바꿈 삽입"""
        self._hwp.HAction.Run("BreakPara")

    # --- Character formatting ---

    def set_char_shape(self, font=None, size=None, bold=None, italic=None,
                       underline=None, strikeout=None, color=None):
        """글자 서식 설정

        Args:
            font: 글꼴 이름 (예: "맑은 고딕", "HY견고딕")
            size: 글자 크기 (pt 단위, 예: 10, 12, 16)
            bold: 굵게 (True/False)
            italic: 기울임 (True/False)
            underline: 밑줄 (True/False)
            strikeout: 취소선 (True/False)
            color: 글자 색상 (RGB 정수, 예: 0xFF0000 = 빨강)
        """
        hwp = self._hwp
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        cs = hwp.HParameterSet.HCharShape

        if font is not None:
            cs.FaceNameUser = font
            cs.FaceNameSymbol = font
            cs.FaceNameOther = font
            cs.FaceNameJapanese = font
            cs.FaceNameHanja = font
            cs.FaceNameLatin = font
            cs.FaceNameHangul = font

        if size is not None:
            cs.Height = int(size * 100)

        if bold is not None:
            cs.Bold = bold

        if italic is not None:
            cs.Italic = italic

        if underline is not None:
            cs.UnderlineType = 1 if underline else 0

        if strikeout is not None:
            cs.StrikeOutType = 1 if strikeout else 0

        if color is not None:
            cs.TextColor = color

        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    # --- Paragraph formatting ---

    def set_para_shape(self, align=None, line_spacing=None, line_spacing_type=None,
                       space_before=None, space_after=None,
                       indent_left=None, indent_right=None, first_line_indent=None):
        """문단 서식 설정

        Args:
            align: 정렬 ("left", "center", "right", "justify") 또는 정수 (0-3)
            line_spacing: 줄간격 (%, 예: 160 = 160%)
            line_spacing_type: 줄간격 종류 (0=percent, 1=fixed, 2=betweenlines)
            space_before: 문단 앞 간격 (HWPUNIT)
            space_after: 문단 뒤 간격 (HWPUNIT)
            indent_left: 왼쪽 들여쓰기 (HWPUNIT)
            indent_right: 오른쪽 들여쓰기 (HWPUNIT)
            first_line_indent: 첫줄 들여쓰기 (HWPUNIT)
        """
        hwp = self._hwp
        hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet)
        ps = hwp.HParameterSet.HParaShape

        if align is not None:
            if isinstance(align, str):
                align_map = {
                    "justify": ALIGN_JUSTIFY,
                    "left": ALIGN_LEFT,
                    "right": ALIGN_RIGHT,
                    "center": ALIGN_CENTER,
                }
                align = align_map.get(align.lower(), ALIGN_JUSTIFY)
            ps.AlignType = align

        if line_spacing is not None:
            ps.LineSpacingType = line_spacing_type if line_spacing_type is not None else 0
            ps.LineSpacing = line_spacing

        if space_before is not None:
            ps.PrevSpacing = space_before

        if space_after is not None:
            ps.NextSpacing = space_after

        if indent_left is not None:
            ps.LeftMargin = indent_left

        if indent_right is not None:
            ps.RightMargin = indent_right

        if first_line_indent is not None:
            ps.Indentation = first_line_indent

        hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet)

    # --- Table operations ---

    def insert_table(self, rows, cols, width_type=2, height_type=0):
        """표 삽입

        Args:
            rows: 행 수
            cols: 열 수
            width_type: 너비 유형 (2=컬럼에 맞춤)
            height_type: 높이 유형 (0=자동)
        """
        hwp = self._hwp
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        tc = hwp.HParameterSet.HTableCreation
        tc.Rows = rows
        tc.Cols = cols
        tc.WidthType = width_type
        tc.HeightType = height_type
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

    def fill_table(self, data):
        """현재 표에 데이터 채우기 (표 삽입 직후 호출)

        Args:
            data: 2차원 리스트 [[row0col0, row0col1, ...], [row1col0, ...], ...]
        """
        for row_idx, row in enumerate(data):
            for col_idx, text in enumerate(row):
                self.insert_text(str(text))
                if col_idx < len(row) - 1:
                    self._hwp.HAction.Run("TableRightCell")
            if row_idx < len(data) - 1:
                self._hwp.HAction.Run("TableRightCell")

    def table_next_cell(self):
        """표에서 다음 셀로 이동"""
        self._hwp.HAction.Run("TableRightCell")

    def table_prev_cell(self):
        """표에서 이전 셀로 이동"""
        self._hwp.HAction.Run("TableLeftCell")

    # --- Find and replace ---

    def find_and_replace(self, find_text, replace_text):
        """텍스트 찾아 바꾸기

        Args:
            find_text: 찾을 텍스트
            replace_text: 바꿀 텍스트
        """
        hwp = self._hwp
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        fr = hwp.HParameterSet.HFindReplace
        fr.FindString = find_text
        fr.ReplaceString = replace_text
        fr.IgnoreMessage = 1
        fr.Direction = 0
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    def find_and_replace_all(self, replacements):
        """여러 텍스트를 한번에 찾아 바꾸기

        Args:
            replacements: dict {찾을_텍스트: 바꿀_텍스트, ...}
        """
        for find_text, replace_text in replacements.items():
            self.find_and_replace(find_text, replace_text)

    # --- Field operations ---

    def get_field_list(self):
        """문서의 필드(누름틀) 목록 조회

        Returns:
            list: 필드 이름 목록
        """
        field_list = self._hwp.GetFieldList(0, 0)
        if field_list:
            return field_list.split('\x02')
        return []

    def set_field_text(self, field_name, text):
        """필드(누름틀)에 텍스트 설정

        Args:
            field_name: 필드 이름
            text: 설정할 텍스트
        """
        self._hwp.PutFieldText(field_name, text)

    def get_field_text(self, field_name):
        """필드(누름틀)의 텍스트 조회

        Args:
            field_name: 필드 이름

        Returns:
            str: 필드 텍스트
        """
        return self._hwp.GetFieldText(field_name)

    # --- Control enumeration ---

    def get_controls(self):
        """문서의 컨트롤(표, 이미지 등) 목록 조회

        Returns:
            dict: {컨트롤ID: 개수}
        """
        self._hwp.HAction.Run("MoveDocBegin")
        ctrl = self._hwp.HeadCtrl
        ctrl_types = {}
        while ctrl:
            ctrlid = ctrl.CtrlID
            ctrl_types[ctrlid] = ctrl_types.get(ctrlid, 0) + 1
            ctrl = ctrl.Next
        return ctrl_types

    # --- Context manager ---

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()
        return False
