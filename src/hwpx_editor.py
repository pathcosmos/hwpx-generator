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
import re
import zipfile

from lxml import etree

# 한컴오피스 HWPX 표준 XML 선언부.
# lxml의 xml_declaration=True는 작은따옴표를 사용하고 standalone을 생략하며
# 루트 요소 앞에 줄바꿈을 추가하는데, 한컴오피스는 이를 인식하지 못한다.
HWPX_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'

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
            # 각 엔트리의 원본 압축 방식 보존
            self._compress_types = {
                info.filename: info.compress_type for info in zf.infolist()
            }

        # 원본 XML 선언부 보존 (한컴오피스 호환성)
        raw_text = section_data.decode('utf-8')
        m = re.match(r'(<\?xml\s[^?]*\?>)', raw_text)
        self._xml_decl = m.group(1) if m else HWPX_XML_DECL

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

        구조 보존 원칙: hp:p, hp:run, hp:linesegarray, hp:lineseg 등
        원본 구조 요소는 절대 삭제하지 않고, hp:t(텍스트)만 교체한다.
        다중 단락 셀은 첫 단락에만 텍스트를 넣고 나머지는 빈 상태로 유지한다.

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

        # 셀 내부의 hp:p들 찾기 (hp:subList 안에 있음)
        sub = tc.find('.//hp:subList', NAMESPACES)
        if sub is None:
            return False

        paragraphs = sub.findall('hp:p', NAMESPACES)
        if not paragraphs:
            return False

        # 모든 단락의 모든 run에서 텍스트(hp:t)만 제거 (구조 보존)
        for para in paragraphs:
            for run in para.findall('hp:run', NAMESPACES):
                for old_t in run.findall('hp:t', NAMESPACES):
                    run.remove(old_t)
            # linesegarray를 새 텍스트에 맞게 보정:
            # 텍스트가 바뀌면 기존 lineseg의 textpos가 유효하지 않을 수 있다.
            # 첫 번째 lineseg(textpos=0)만 남기고 나머지는 제거한다.
            self._remove_linesegarray(para)

        # 첫 번째 단락의 첫 번째 run에 텍스트 삽입
        p = paragraphs[0]
        runs = p.findall('hp:run', NAMESPACES)
        if not runs:
            # run이 없으면 새로 생성 (linesegarray 앞에 삽입하여 순서 보존)
            ref_id = self._find_nearby_char_pr_id(table, row_addr, col_addr)
            run = etree.Element(f'{{{HP_NS}}}run')
            run.set('charPrIDRef', ref_id)
            lsa = p.find('hp:linesegarray', NAMESPACES)
            if lsa is not None:
                p.insert(list(p).index(lsa), run)
            else:
                p.append(run)
        else:
            run = runs[0]

        # 새 hp:t 생성 및 추가
        t_elem = etree.SubElement(run, f'{{{HP_NS}}}t')
        t_elem.text = text

        return True

    @staticmethod
    def _remove_linesegarray(p_elem):
        """텍스트 변경 후 단락의 linesegarray를 제거한다.

        linesegarray는 한컴오피스의 레이아웃 캐시로, 텍스트 변경 시
        기존 값이 유효하지 않게 된다. 제거하면 한컴오피스가 문서를
        열 때 자동으로 재계산한다. (원본 문서에도 linesegarray 없는
        단락이 다수 존재하며 정상 처리된다.)
        """
        lsa = p_elem.find('hp:linesegarray', NAMESPACES)
        if lsa is not None:
            p_elem.remove(lsa)

    def _find_nearby_char_pr_id(self, table, row_addr, col_addr):
        """인접 셀의 charPrIDRef를 찾아 반환한다 (fallback: "0")."""
        # 같은 행의 다른 셀 확인
        for tc in table.findall('.//hp:tc', NAMESPACES):
            addr = tc.find('hp:cellAddr', NAMESPACES)
            if addr is None:
                continue
            if int(addr.get('rowAddr', '-1')) == row_addr:
                run = tc.find('.//hp:run', NAMESPACES)
                if run is not None:
                    return run.get('charPrIDRef', '0')
        # 첫 번째 셀의 run에서 가져오기
        first_run = table.find('.//hp:run', NAMESPACES)
        if first_run is not None:
            return first_run.get('charPrIDRef', '0')
        return '0'

    def get_table_count(self):
        """문서 내 전체 테이블 수를 반환한다."""
        return len(self.root.findall('.//hp:tbl', NAMESPACES))

    def remove_memos(self):
        """문서 내 모든 MEMO(메모/주석) 필드를 제거한다.

        양식 템플릿에 포함된 작성 지침 메모(예: "해당없을 시 행 삭제",
        "user 2026/01/16 ..." 등)를 삭제한다.

        MEMO 필드 구조: hp:run 내에 fieldBegin ctrl과 fieldEnd ctrl이
        쌍으로 존재한다. 둘 다 제거해야 문서가 깨지지 않는다.

        Returns:
            int: 제거된 MEMO 수
        """
        # 1단계: MEMO fieldBegin ctrl을 찾아 제거하고,
        #         같은 부모(run) 내의 다음 fieldEnd ctrl도 함께 제거
        count = 0
        memo_ctrls = self.root.findall('.//hp:ctrl', NAMESPACES)
        for ctrl in memo_ctrls:
            fb = ctrl.find('hp:fieldBegin[@type="MEMO"]', NAMESPACES)
            if fb is None:
                continue
            parent = ctrl.getparent()
            if parent is None:
                continue

            # 같은 부모 내에서 이 ctrl 뒤에 오는 fieldEnd ctrl 찾기
            siblings = list(parent)
            ctrl_idx = siblings.index(ctrl) if ctrl in siblings else -1
            end_ctrl = None
            if ctrl_idx >= 0:
                for sibling in siblings[ctrl_idx + 1:]:
                    if (sibling.tag.endswith('}ctrl')
                            and sibling.find('hp:fieldEnd', NAMESPACES) is not None):
                        end_ctrl = sibling
                        break

            # fieldEnd ctrl 먼저 제거 (인덱스 밀림 방지)
            if end_ctrl is not None:
                parent.remove(end_ctrl)
            parent.remove(ctrl)
            count += 1

        # 2단계: 다른 run/paragraph에 남은 orphan fieldEnd ctrl 정리.
        # MEMO fieldBegin이 모두 제거된 후 남은 fieldEnd는 전부 orphan이다.
        remaining_begins = self.root.findall(
            './/hp:ctrl/hp:fieldBegin', NAMESPACES)
        if not remaining_begins:
            # 문서에 non-MEMO fieldBegin이 없으면 모든 fieldEnd는 orphan
            for ctrl in list(self.root.findall('.//hp:ctrl', NAMESPACES)):
                if ctrl.find('hp:fieldEnd', NAMESPACES) is not None:
                    parent = ctrl.getparent()
                    if parent is not None:
                        parent.remove(ctrl)

        return count

    def inject_marker(self, after_table_index, marker_text):
        """지정된 테이블 뒤에 마커 텍스트가 포함된 새 문단을 삽입한다.

        COM Pass에서 이 마커를 찾아 실제 콘텐츠로 교체한다.

        Args:
            after_table_index: 마커를 삽입할 테이블 인덱스 (이 테이블 뒤에 삽입)
            marker_text: 마커 문자열 (예: '##SEC1_CONTENT##')

        Returns:
            True if successful, False otherwise
        """
        tables = self.root.findall('.//hp:tbl', NAMESPACES)
        if after_table_index < 0 or after_table_index >= len(tables):
            return False

        tbl_elem = tables[after_table_index]

        # 부모 체인: hp:tbl → hp:run → hp:p (앵커 문단) → hs:sec
        # hs:sec 레벨에서 새 hp:p를 sibling으로 삽입해야 한다.
        run_parent = tbl_elem.getparent()       # hp:run
        if run_parent is None:
            return False
        wrapper_p = run_parent.getparent()      # hp:p (테이블 앵커 문단)
        if wrapper_p is None:
            return False
        sec_root = wrapper_p.getparent()        # hs:sec
        if sec_root is None:
            return False

        # 부모 체인 검증
        assert run_parent.tag.endswith('}run'), \
            f"Expected hp:run, got {run_parent.tag}"
        assert wrapper_p.tag.endswith('}p'), \
            f"Expected hp:p, got {wrapper_p.tag}"

        # 바탕글(기본) 스타일 사용 — 테이블 앵커 스타일("표(가운데로)" 등)을
        # 상속하면 COM 삽입 텍스트가 PDF에서 0.1pt로 렌더링됨
        para_pr_ref = '0'
        style_ref = '0'

        # wrapper_p 내 run에서 charPrIDRef 참조
        ref_run = wrapper_p.find('.//hp:run', NAMESPACES)
        char_pr_ref = ref_run.get('charPrIDRef', '0') if ref_run is not None else '0'

        # 새 hp:p 문단 생성 (원본 구조를 최대한 복제)
        new_p = etree.Element(f'{{{HP_NS}}}p')
        new_p.set('paraPrIDRef', para_pr_ref)
        new_p.set('styleIDRef', style_ref)

        # hp:run + hp:t 생성 (run이 linesegarray보다 앞에 위치해야 함)
        run = etree.SubElement(new_p, f'{{{HP_NS}}}run')
        run.set('charPrIDRef', char_pr_ref)
        t_elem = etree.SubElement(run, f'{{{HP_NS}}}t')
        t_elem.text = marker_text

        # linesegarray 생략: 한컴오피스가 문서를 열 때 자동 재계산한다.
        # 원본 문서에도 linesegarray 없는 단락이 21개 이상 존재하며 정상 처리.

        # hs:sec 레벨에서 앵커 문단 바로 뒤에 삽입
        p_index = list(sec_root).index(wrapper_p)
        sec_root.insert(p_index + 1, new_p)
        return True

    def remove_outline_placeholders(self, start_table=2, end_table=None):
        """마커 주입 후, 마커와 다음 테이블 사이의 빈 개요 단락을 제거한다.

        양식 템플릿의 작성 가이드/아웃라인 구조(섹션 번호, □, ○, - 등)를
        모두 제거하여 COM이 삽입한 본문만 남기도록 한다.

        hs:sec 직속 hp:p 중 테이블 앵커가 아니고 마커 텍스트도 아닌 것을 제거.
        start_table부터 end_table까지의 범위만 처리한다.
        """
        sec = self.root
        children = list(sec)

        # 테이블 앵커 단락 식별 (hp:tbl을 포함하는 hp:p)
        table_anchors = set()
        all_tables = self.root.findall('.//hp:tbl', NAMESPACES)
        for tbl in all_tables:
            # hp:tbl → hp:run → hp:p
            run = tbl.getparent()
            if run is not None:
                p = run.getparent()
                if p is not None and p.tag.endswith('}p'):
                    table_anchors.add(p)

        # 테이블 순서대로 인덱스 매핑
        tbl_idx_map = {}  # table anchor paragraph → table index
        tbl_counter = 0
        for child in children:
            if child in table_anchors:
                tbl_idx_map[child] = tbl_counter
                tbl_counter += 1

        if end_table is None:
            end_table = tbl_counter - 1

        # start_table ~ end_table 범위 내 비-테이블/비-마커 단락 제거
        in_range = False
        to_remove = []
        for child in children:
            if not child.tag.endswith('}p'):
                continue

            if child in table_anchors:
                idx = tbl_idx_map[child]
                if idx >= start_table:
                    in_range = True
                if idx > end_table:
                    in_range = False
                continue  # 테이블 앵커는 보존

            if not in_range:
                continue

            # 마커 텍스트 확인
            texts = []
            for t in child.findall('.//hp:t', NAMESPACES):
                if t.text:
                    texts.append(t.text)
            full_text = ''.join(texts).strip()

            if full_text.startswith('##') and full_text.endswith('##'):
                continue  # 마커 보존

            to_remove.append(child)

        for p in to_remove:
            sec.remove(p)

        return len(to_remove)

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

    def serialize_xml(self):
        """section0.xml을 한컴오피스 호환 바이트열로 직렬화한다.

        lxml의 기본 xml_declaration은 한컴오피스와 호환되지 않으므로
        원본 XML 선언부를 보존하여 직접 구성한다.

        Returns:
            bytes: UTF-8 인코딩된 XML 바이트열
        """
        body = etree.tostring(self.root, xml_declaration=False, encoding='unicode')
        return (self._xml_decl + body).encode('utf-8')

    def save(self, output_path=None):
        """수정된 section0.xml을 포함하여 HWPX ZIP을 다시 생성한다.

        원본 ZIP의 엔트리 순서와 각 파일의 압축 방식(STORED/DEFLATED)을
        그대로 유지한다. 한컴오피스는 이 형식에 민감하다.

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
                        data = self.serialize_xml()
                        item.file_size = len(data)

                    # 원본 압축 방식 보존
                    compress = self._compress_types.get(
                        item.filename, item.compress_type
                    )
                    zout.writestr(item, data, compress_type=compress)

        os.replace(tmp_path, output_path)
