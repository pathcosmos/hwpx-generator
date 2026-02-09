"""HWPX 파일의 section0.xml을 수정하는 편집기.

lxml을 사용하여 모든 네임스페이스 선언을 보존하면서
표 셀의 텍스트를 추가/변경할 수 있다.

Usage:
    editor = HwpxEditor('ref/test_01.hwpx')
    table = editor.get_table(0)          # 첫 번째 표
    editor.set_cell_text(table, 6, 3, '홍길동')  # row=6, col=3에 텍스트 설정
    editor.save('output.hwpx')
"""

import os
import zipfile

from lxml import etree

NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
    'hpf': 'http://www.hancom.co.kr/schema/2011/hpf',
    'opf': 'http://www.idpf.org/2007/opf/',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'hhs': 'http://www.hancom.co.kr/hwpml/2011/history',
    'hm': 'http://www.hancom.co.kr/hwpml/2011/master-page',
    'ooxmlchart': 'http://www.hancom.co.kr/hwpml/2016/ooxmlchart',
    'hwpunitchar': 'http://www.hancom.co.kr/hwpml/2016/HwpUnitChar',
    'epub': 'http://www.idpf.org/2007/ops',
    'config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
}

HP_NS = NAMESPACES['hp']


class HwpxEditor:
    """HWPX ZIP 내부의 section0.xml을 수정하는 편집기."""

    def __init__(self, hwpx_path):
        """HWPX 파일을 열고 section0.xml을 파싱한다.

        Args:
            hwpx_path: HWPX 파일 경로
        """
        self.hwpx_path = hwpx_path

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            section_data = zf.read('Contents/section0.xml')

        self.root = etree.fromstring(section_data)

    def get_table(self, index=0):
        """N번째 hp:tbl 요소를 반환한다.

        Args:
            index: 표 인덱스 (0부터 시작)

        Returns:
            lxml Element (hp:tbl) 또는 None
        """
        tables = self.root.findall('.//hp:tbl', NAMESPACES)
        if 0 <= index < len(tables):
            return tables[index]
        return None

    def get_cell(self, table, row_addr, col_addr):
        """표에서 지정된 위치의 hp:tc 요소를 찾는다.

        Args:
            table: hp:tbl 요소
            row_addr: 행 주소 (hp:cellAddr의 rowAddr)
            col_addr: 열 주소 (hp:cellAddr의 colAddr)

        Returns:
            lxml Element (hp:tc) 또는 None
        """
        for tc in table.findall('.//hp:tc', NAMESPACES):
            addr = tc.find('hp:cellAddr', NAMESPACES)
            if addr is not None:
                if (int(addr.get('colAddr', '-1')) == col_addr
                        and int(addr.get('rowAddr', '-1')) == row_addr):
                    return tc
        return None

    def set_cell_text(self, table, row_addr, col_addr, text):
        """표 셀의 텍스트를 설정한다.

        기존 hp:run의 charPrIDRef는 변경하지 않고,
        hp:t 자식만 추가/교체한다.

        Args:
            table: hp:tbl 요소
            row_addr: 행 주소
            col_addr: 열 주소
            text: 설정할 텍스트

        Returns:
            True if successful, False otherwise
        """
        tc = self.get_cell(table, row_addr, col_addr)
        if tc is None:
            return False

        # 셀 내부의 첫 번째 hp:p 찾기 (hp:subList 안에 있음)
        p = tc.find('.//hp:p', NAMESPACES)
        if p is None:
            return False

        # hp:run 찾기
        run = p.find('hp:run', NAMESPACES)
        if run is None:
            return False

        # 기존 hp:t 자식들 제거
        for old_t in run.findall('hp:t', NAMESPACES):
            run.remove(old_t)

        # 새 hp:t 생성 및 추가
        t_elem = etree.SubElement(run, f'{{{HP_NS}}}t')
        t_elem.text = text

        return True

    def fill_cells(self, table, cell_data):
        """여러 셀의 텍스트를 일괄 설정한다.

        Args:
            table: hp:tbl 요소
            cell_data: {(row_addr, col_addr): text} 딕셔너리

        Returns:
            성공한 셀 수
        """
        count = 0
        for (row_addr, col_addr), text in cell_data.items():
            if self.set_cell_text(table, row_addr, col_addr, text):
                count += 1
        return count

    def save(self, output_path=None):
        """수정된 section0.xml을 포함하여 HWPX ZIP을 다시 생성한다.

        Args:
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
        """
        if output_path is None:
            output_path = self.hwpx_path

        tmp_path = output_path + '.tmp'

        with zipfile.ZipFile(self.hwpx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w') as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename == 'Contents/section0.xml':
                        data = etree.tostring(
                            self.root,
                            xml_declaration=True,
                            encoding='UTF-8',
                        )
                        item.file_size = len(data)

                    if item.filename == 'mimetype':
                        zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.writestr(item, data, compress_type=item.compress_type)

        os.replace(tmp_path, output_path)
