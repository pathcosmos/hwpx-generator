"""HWPX 문서에서 내용과 구조를 추출하는 도구.

HWPX 파일의 section0.xml과 header.xml을 파싱하여
문서 구조(커버 페이지, 본문 섹션, 표 등)를 JSON으로 추출한다.
임의의 HWPX 파일을 --hwpx 인자로 지정할 수 있다.

Usage:
    python3 src/extract_template.py --hwpx ref/test_01.hwpx             # 전체 구조 추출
    python3 src/extract_template.py --hwpx ref/test_01.hwpx --cover     # 커버 페이지만
    python3 src/extract_template.py --hwpx ref/test_01.hwpx --all-tables  # 모든 표 목록
    python3 src/extract_template.py --hwpx ref/새양식.hwpx --generate-template-config -o templates/새양식/
"""

import zipfile
import xml.etree.ElementTree as ET
import json
import sys
import os
from pathlib import Path

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
}

# Register namespaces so ET doesn't rename prefixes
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

HWPX_PATH = Path(__file__).resolve().parent.parent / 'ref' / 'test_01.hwpx'


def read_xml_from_hwpx(hwpx_path, inner_path):
    """HWPX(ZIP) 파일에서 특정 XML 파일을 읽어 ElementTree root를 반환."""
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        data = z.read(inner_path)
    return ET.fromstring(data)


def get_cell_text(tc):
    """hp:tc 요소에서 모든 텍스트를 추출. 줄바꿈은 \\n으로 구분."""
    lines = []
    for sub in tc.findall('.//hp:subList', NAMESPACES):
        for p in sub.findall('hp:p', NAMESPACES):
            parts = []
            for t in p.findall('.//hp:t', NAMESPACES):
                if t.text:
                    parts.append(t.text)
            line = ''.join(parts).strip()
            if line:
                lines.append(line)
    return '\n'.join(lines)


def get_paragraph_text(p_elem):
    """hp:p 요소에서 직접 포함된 run의 텍스트만 추출 (하위 표/그림 제외)."""
    texts = []
    for run in p_elem.findall('hp:run', NAMESPACES):
        for t in run.findall('hp:t', NAMESPACES):
            if t.text:
                texts.append(t.text)
    return ''.join(texts).strip()


def extract_text_runs(p_elem):
    """hp:p 요소에서 텍스트 run들을 추출. 각 run의 텍스트와 charPrIDRef 반환."""
    runs = []
    for run in p_elem.findall('hp:run', NAMESPACES):
        char_pr_id = run.get('charPrIDRef', '')
        text_parts = []
        for t in run.findall('hp:t', NAMESPACES):
            if t.text:
                text_parts.append(t.text)
        text = ''.join(text_parts)
        if text.strip():
            runs.append({
                'text': text,
                'charPrIDRef': char_pr_id,
            })
    return runs


def extract_table(tbl_elem):
    """hp:tbl 요소에서 표 구조를 추출."""
    result = {
        'rowCnt': int(tbl_elem.get('rowCnt', '0')),
        'colCnt': int(tbl_elem.get('colCnt', '0')),
        'borderFillIDRef': tbl_elem.get('borderFillIDRef', ''),
        'rows': [],
    }

    sz = tbl_elem.find('hp:sz', NAMESPACES)
    if sz is not None:
        result['width'] = int(sz.get('width', '0'))
        result['height'] = int(sz.get('height', '0'))

    for tr in tbl_elem.findall('hp:tr', NAMESPACES):
        row_cells = []
        for tc in tr.findall('hp:tc', NAMESPACES):
            cell = {}

            addr = tc.find('hp:cellAddr', NAMESPACES)
            if addr is not None:
                cell['colAddr'] = int(addr.get('colAddr', '0'))
                cell['rowAddr'] = int(addr.get('rowAddr', '0'))

            span = tc.find('hp:cellSpan', NAMESPACES)
            if span is not None:
                cs = int(span.get('colSpan', '1'))
                rs = int(span.get('rowSpan', '1'))
                if cs > 1:
                    cell['colSpan'] = cs
                if rs > 1:
                    cell['rowSpan'] = rs

            cell['text'] = get_cell_text(tc)
            cell['borderFillIDRef'] = tc.get('borderFillIDRef', '')

            cell_sz = tc.find('hp:cellSz', NAMESPACES)
            if cell_sz is not None:
                cell['width'] = int(cell_sz.get('width', '0'))
                cell['height'] = int(cell_sz.get('height', '0'))

            row_cells.append(cell)
        result['rows'].append(row_cells)

    return result


def extract_cover_table(hwpx_path):
    """커버 페이지 표(35x11)를 추출하여 필드 매핑 반환."""
    root = read_xml_from_hwpx(hwpx_path, 'Contents/section0.xml')
    top_paras = root.findall('hp:p', NAMESPACES)

    # 첫 번째 단락에 커버 표가 있음
    p0 = top_paras[0]
    tbl = p0.find('.//hp:tbl', NAMESPACES)
    if tbl is None:
        return None

    table_data = extract_table(tbl)

    # 셀을 (rowAddr, colAddr) -> cell 맵으로 변환
    cell_map = {}
    for row_cells in table_data['rows']:
        for cell in row_cells:
            key = (cell.get('rowAddr', 0), cell.get('colAddr', 0))
            cell_map[key] = cell

    # 커버 필드 매핑
    cover = {
        '문서제목': cell_map.get((0, 0), {}).get('text', ''),
        '사업명': cell_map.get((1, 3), {}).get('text', ''),
        '사업개요': cell_map.get((2, 3), {}).get('text', ''),
        '과제명': cell_map.get((3, 3), {}).get('text', ''),
        '개발솔루션기능': cell_map.get((4, 3), {}).get('text', ''),
        '수행기간': cell_map.get((5, 3), {}).get('text', ''),
        '대표공급기업': _extract_entity_block(cell_map, start_row=6, label_col=1, value_col=3,
                                          right_label_col=7, right_value_col=9),
        '클라우드사업자': _extract_entity_block(cell_map, start_row=13, label_col=1, value_col=3,
                                           right_label_col=7, right_value_col=9),
        '협력기관': _extract_entity_block(cell_map, start_row=19, label_col=1, value_col=3,
                                      right_label_col=7, right_value_col=9),
        '참여공급기업': _extract_company_list(cell_map, start_row=25, end_row=28,
                                                has_header=True),
        '도입실증기업': _extract_company_list(cell_map, start_row=29, end_row=33,
                                              has_header=False),
        '서명문구': cell_map.get((34, 0), {}).get('text', ''),
    }

    return cover


def _extract_entity_block(cell_map, start_row, label_col, value_col,
                          right_label_col, right_value_col):
    """기업/기관 정보 블록 추출 (대표공급기업, 클라우드사업자, 협력기관)."""
    # 필드 매핑: (relative_row, side) -> field_name
    field_rows = [
        (0, '기업명', '사업자등록번호'),
        (1, '대표자명', '법인등록번호'),
        (2, '본사정보', '지역'),
        (3, '주요솔루션', None),  # 주요 솔루션은 전체 행 병합
        (4, '성명', 'E-mail'),
        (5, '부서', '전화'),
        (6, '직위', '휴대전화'),
    ]

    entity = {}
    for offset, left_field, right_field in field_rows:
        row = start_row + offset
        if left_field:
            entity[left_field] = cell_map.get((row, value_col), {}).get('text', '')
        if right_field:
            entity[right_field] = cell_map.get((row, right_value_col), {}).get('text', '')

    # 담당자 정보를 하위 객체로 정리
    contact_fields = ['성명', 'E-mail', '부서', '전화', '직위', '휴대전화']
    담당자 = {}
    for f in contact_fields:
        if f in entity:
            담당자[f] = entity.pop(f)
    if any(담당자.values()):
        entity['담당자'] = 담당자

    return entity


def _extract_company_list(cell_map, start_row, end_row, has_header=True):
    """참여공급기업/도입실증기업 목록 추출."""
    # has_header=True: 헤더 행이 start_row에 있고 데이터는 start_row+1부터
    # has_header=False: 데이터가 start_row부터 시작 (도입실증기업)
    data_start = start_row + 1 if has_header else start_row
    companies = []
    for row in range(data_start, end_row + 1):
        company = {
            '번호': cell_map.get((row, 1), {}).get('text', ''),
            '기업명': cell_map.get((row, 2), {}).get('text', ''),
            '대표자명': cell_map.get((row, 4), {}).get('text', ''),
            '전화': cell_map.get((row, 5), {}).get('text', ''),
            '휴대전화': cell_map.get((row, 6), {}).get('text', ''),
            'E-mail': cell_map.get((row, 8), {}).get('text', ''),
            '지역': cell_map.get((row, 10), {}).get('text', ''),
        }
        companies.append(company)
    return companies


def extract_body_sections(hwpx_path):
    """본문 섹션 구조를 추출 (목차 이후의 실제 내용)."""
    root = read_xml_from_hwpx(hwpx_path, 'Contents/section0.xml')
    top_paras = root.findall('hp:p', NAMESPACES)

    sections = []
    current_section = None
    current_subsection = None

    # 본문은 P36 ("1. 솔루션 구축 개요") 부터 시작
    # P0-P1: 커버+작성요령, P2: 빈 줄, P3: 목차, P4-P35: TOC 항목
    for i, p in enumerate(top_paras):
        text = get_paragraph_text(p)
        para_pr = p.get('paraPrIDRef', '')
        style = p.get('styleIDRef', '0')

        # 직접 포함된 표/그림 확인
        has_table = False
        has_pic = False
        table_info = None
        for run in p.findall('hp:run', NAMESPACES):
            tbl = run.find('hp:tbl', NAMESPACES)
            if tbl is not None:
                has_table = True
                table_info = {
                    'rowCnt': int(tbl.get('rowCnt', '0')),
                    'colCnt': int(tbl.get('colCnt', '0')),
                }
            if run.find('hp:pic', NAMESPACES) is not None:
                has_pic = True

        # 메인 섹션 헤더 감지 (예: "1. 솔루션 구축 개요")
        if text and _is_main_section_header(text, para_pr):
            current_section = {
                'type': 'section',
                'title': text,
                'paraIndex': i,
                'paraPrIDRef': para_pr,
                'children': [],
            }
            sections.append(current_section)
            current_subsection = None
            continue

        # 서브섹션 헤더 감지 (예: "1.1 솔루션 개발 배경 및 필요성")
        if text and _is_subsection_header(text, para_pr):
            current_subsection = {
                'type': 'subsection',
                'title': text,
                'paraIndex': i,
                'paraPrIDRef': para_pr,
                'children': [],
            }
            if current_section:
                current_section['children'].append(current_subsection)
            continue

        # 일반 콘텐츠
        content_item = None
        if has_table and table_info:
            content_item = {
                'type': 'table',
                'paraIndex': i,
                'rowCnt': table_info['rowCnt'],
                'colCnt': table_info['colCnt'],
            }
            if text:
                content_item['caption'] = text
        elif has_pic:
            content_item = {
                'type': 'picture',
                'paraIndex': i,
            }
            if text:
                content_item['caption'] = text
        elif text:
            content_item = {
                'type': 'paragraph',
                'paraIndex': i,
                'text': text,
                'paraPrIDRef': para_pr,
            }

        if content_item:
            target = current_subsection if current_subsection else current_section
            if target:
                target['children'].append(content_item)

    return sections


def _is_main_section_header(text, para_pr):
    """메인 섹션 헤더인지 판별 (숫자. 으로 시작, 특정 paraPr)."""
    import re
    # "1. ", "2. " 등으로 시작하며, 서브섹션 번호(1.1)가 아닌 것
    if re.match(r'^\d+\.\s+\S', text) and not re.match(r'^\d+\.\d+', text):
        return True
    return False


def _is_subsection_header(text, para_pr):
    """서브섹션 헤더인지 판별."""
    import re
    # "1.1 ", "2.3. " 등으로 시작
    if re.match(r'^\d+\.\d+\.?\s+\S', text):
        return True
    # "□ " 로 시작하는 항목 제목
    if text.startswith('□ '):
        return True
    return False


def extract_styles(hwpx_path):
    """header.xml에서 사용된 스타일 정보를 추출."""
    root = read_xml_from_hwpx(hwpx_path, 'Contents/header.xml')

    result = {
        'fonts': {},
        'charProperties': [],
        'paraProperties': [],
        'borderFills': [],
        'styles': [],
    }

    # 폰트 정보
    for fontface in root.findall('.//hh:fontface', NAMESPACES):
        lang = fontface.get('lang', '')
        fonts = []
        for font in fontface.findall('hh:font', NAMESPACES):
            fonts.append({
                'id': font.get('id', ''),
                'face': font.get('face', ''),
                'type': font.get('type', ''),
            })
        result['fonts'][lang] = fonts

    # 글자 속성 (charPr) - 주요 속성만
    for cp in root.findall('.//hh:charPr', NAMESPACES):
        char_prop = {
            'id': cp.get('id', ''),
            'height': cp.get('height', ''),
            'textColor': cp.get('textColor', ''),
            'borderFillIDRef': cp.get('borderFillIDRef', ''),
        }
        # bold/italic 확인
        if cp.find('hh:bold', NAMESPACES) is not None:
            char_prop['bold'] = True
        if cp.find('hh:italic', NAMESPACES) is not None:
            char_prop['italic'] = True
        # 폰트 참조
        font_ref = cp.find('hh:fontRef', NAMESPACES)
        if font_ref is not None:
            char_prop['fontRef'] = {
                'hangul': font_ref.get('hangul', ''),
                'latin': font_ref.get('latin', ''),
            }
        result['charProperties'].append(char_prop)

    # 문단 속성 (paraPr) - 주요 속성만
    for pp in root.findall('.//hh:paraPr', NAMESPACES):
        para_prop = {
            'id': pp.get('id', ''),
        }
        align = pp.find('hh:align', NAMESPACES)
        if align is not None:
            para_prop['horizontal'] = align.get('horizontal', '')
            para_prop['vertical'] = align.get('vertical', '')
        heading = pp.find('hh:heading', NAMESPACES)
        if heading is not None:
            para_prop['headingType'] = heading.get('type', '')
            para_prop['headingLevel'] = heading.get('level', '')
        result['paraProperties'].append(para_prop)

    # 테두리/채우기 (borderFill) - 주요 속성만
    for bf in root.findall('.//hh:borderFill', NAMESPACES):
        border = {
            'id': bf.get('id', ''),
        }
        for side in ['leftBorder', 'rightBorder', 'topBorder', 'bottomBorder']:
            elem = bf.find(f'hh:{side}', NAMESPACES)
            if elem is not None:
                border[side] = {
                    'type': elem.get('type', ''),
                    'width': elem.get('width', ''),
                    'color': elem.get('color', ''),
                }
        fill = bf.find('.//hh:winBrush', NAMESPACES)
        if fill is not None:
            border['fillColor'] = fill.get('faceColor', '')
        result['borderFills'].append(border)

    # 스타일 정의
    for style in root.findall('.//hh:style', NAMESPACES):
        result['styles'].append({
            'id': style.get('id', ''),
            'type': style.get('type', ''),
            'name': style.get('name', ''),
            'engName': style.get('engName', ''),
            'paraPrIDRef': style.get('paraPrIDRef', ''),
            'charPrIDRef': style.get('charPrIDRef', ''),
            'nextStyleIDRef': style.get('nextStyleIDRef', ''),
        })

    return result


def extract_document_structure(hwpx_path):
    """HWPX 파일에서 전체 문서 구조를 추출."""
    return {
        'cover': extract_cover_table(hwpx_path),
        'sections': extract_body_sections(hwpx_path),
    }


def extract_key_tables(hwpx_path):
    """본문의 주요 데이터 표들을 추출 (스키마 정의에 필요한 표)."""
    root = read_xml_from_hwpx(hwpx_path, 'Contents/section0.xml')
    top_paras = root.findall('hp:p', NAMESPACES)

    key_tables = {}

    # 추진체계 (P163: 6x14)
    _try_extract_table(top_paras, 163, '추진체계', key_tables)

    # 추진체계별 주요 역할 (P166: 7x2)
    _try_extract_table(top_paras, 166, '추진체계별역할', key_tables)

    # 핵심성과지표 KPI (P391: 6x8)
    _try_extract_table(top_paras, 391, 'KPI', key_tables)

    # KPI 측정방법 (P396: 5x2)
    _try_extract_table(top_paras, 396, 'KPI측정방법', key_tables)

    # 추진일정 (P458: 15x17)
    _try_extract_table(top_paras, 458, '추진일정', key_tables)

    # 사업비 총괄 (P464: 5x8)
    _try_extract_table(top_paras, 464, '사업비총괄', key_tables)

    # 참여인력 총괄 (P510: 8x8)
    _try_extract_table(top_paras, 510, '참여인력총괄', key_tables)

    # 국내표준 (P447: 3x2)
    _try_extract_table(top_paras, 447, '국내표준', key_tables)

    # 해외표준 (P450: 3x2)
    _try_extract_table(top_paras, 450, '해외표준', key_tables)

    # 산출물 예정목록 (P455: 13x3)
    _try_extract_table(top_paras, 455, '산출물예정목록', key_tables)

    # 개발개요 (P160: 6x2)
    _try_extract_table(top_paras, 160, '개발개요', key_tables)

    return key_tables


def _try_extract_table(top_paras, para_idx, name, result_dict):
    """특정 인덱스의 단락에서 표를 추출 시도."""
    if para_idx < len(top_paras):
        p = top_paras[para_idx]
        tbl = p.find('.//hp:tbl', NAMESPACES)
        if tbl is not None:
            result_dict[name] = extract_table(tbl)


def generate_sample_data(hwpx_path):
    """참조 문서에서 실제 내용을 추출하여 sample_input.json 형태로 반환."""
    cover = extract_cover_table(hwpx_path)
    key_tables = extract_key_tables(hwpx_path)

    # 개발 솔루션 목록 (줄바꿈으로 구분)
    sol_text = cover.get('개발솔루션기능', '')
    솔루션목록 = [s.strip() for s in sol_text.split('\n') if s.strip()] if sol_text else []
    # 줄바꿈이 없는 경우 알려진 패턴으로 분리
    if len(솔루션목록) <= 1 and sol_text:
        import re
        # 알려진 솔루션 이름 패턴으로 분리
        patterns = [
            '상태기반 설비보전시스템 (eCMMS)',
            '공장에너지관리시스템 (FEMS)',
            '웹기반 엔지니어링 설계 자산(CAD)관리시스템',
            '클라우드 기반 문서관리 시스템',
            'RAG 기반 LLM 지식검색 서비스',
        ]
        found = []
        for pat in patterns:
            if pat in sol_text:
                found.append(pat)
        if found:
            솔루션목록 = found

    # 수행기간 파싱
    period_text = cover.get('수행기간', '')
    수행기간 = {
        '원본텍스트': period_text,
        '개발시작': "'26.6.30",
        '개발종료': "'27.6.30",
        '개발기간': '12개월',
        '실증시작': "'27.6.30",
        '실증종료': "'27.12.31",
        '실증기간': '6개월',
    }

    # 개발개요 표에서 정보 추출
    개발개요 = {}
    if '개발개요' in key_tables:
        tbl = key_tables['개발개요']
        for row_cells in tbl['rows']:
            if len(row_cells) >= 2:
                label = row_cells[0].get('text', '').strip()
                value = row_cells[1].get('text', '').strip()
                if label and value:
                    개발개요[label] = value

    sample = {
        '사업명': cover.get('사업명', ''),
        '과제명': cover.get('과제명', ''),
        '사업개요': cover.get('사업개요', ''),
        '개발솔루션': 솔루션목록,
        '수행기간': 수행기간,
        '대표공급기업': cover.get('대표공급기업', {}),
        '클라우드사업자': cover.get('클라우드사업자', {}),
        '협력기관': cover.get('협력기관', {}),
        '참여공급기업': cover.get('참여공급기업', []),
        '도입실증기업': cover.get('도입실증기업', []),
        '개발개요': 개발개요,
        '서명문구': cover.get('서명문구', ''),
    }

    return sample


def extract_all_tables(hwpx_path):
    """HWPX 파일에서 모든 표를 추출하여 요약 목록 반환."""
    root = read_xml_from_hwpx(hwpx_path, 'Contents/section0.xml')
    tables = root.findall('.//hp:tbl', NAMESPACES)

    result = []
    for i, tbl in enumerate(tables):
        row_cnt = int(tbl.get('rowCnt', '0'))
        col_cnt = int(tbl.get('colCnt', '0'))

        # 첫 번째 셀의 텍스트를 미리보기로 추출
        first_cell_text = ''
        first_tc = tbl.find('.//hp:tc', NAMESPACES)
        if first_tc is not None:
            first_cell_text = get_cell_text(first_tc)[:50]

        sz = tbl.find('hp:sz', NAMESPACES)
        width = int(sz.get('width', '0')) if sz is not None else 0
        height = int(sz.get('height', '0')) if sz is not None else 0

        result.append({
            'index': i,
            'rowCnt': row_cnt,
            'colCnt': col_cnt,
            'width': width,
            'height': height,
            'preview': first_cell_text,
        })

    return result


def generate_template_config(hwpx_path, output_dir):
    """HWPX 파일을 분석하여 template.json과 field_map.json 초안을 생성.

    Args:
        hwpx_path: 분석할 HWPX 파일 경로
        output_dir: 설정 파일을 저장할 디렉토리

    Returns:
        dict: 생성 결과 요약
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    root = read_xml_from_hwpx(hwpx_path, 'Contents/section0.xml')
    all_tables = root.findall('.//hp:tbl', NAMESPACES)

    # 가장 큰 표를 커버 테이블 후보로 선택
    cover_idx = 0
    max_cells = 0
    for i, tbl in enumerate(all_tables):
        rows = int(tbl.get('rowCnt', '0'))
        cols = int(tbl.get('colCnt', '0'))
        cells = rows * cols
        if cells > max_cells:
            max_cells = cells
            cover_idx = i

    cover_tbl = all_tables[cover_idx] if all_tables else None

    # 커버 테이블에서 빈 셀 위치 분석 → field_map 초안 생성
    entity_blocks = []
    if cover_tbl is not None:
        row_cnt = int(cover_tbl.get('rowCnt', '0'))
        col_cnt = int(cover_tbl.get('colCnt', '0'))

        # 셀 맵 구성
        cell_map = {}
        for tr in cover_tbl.findall('hp:tr', NAMESPACES):
            for tc in tr.findall('hp:tc', NAMESPACES):
                addr = tc.find('hp:cellAddr', NAMESPACES)
                if addr is not None:
                    r = int(addr.get('rowAddr', '0'))
                    c = int(addr.get('colAddr', '0'))
                    text = get_cell_text(tc)
                    cell_map[(r, c)] = text

        # 비어있는 셀들을 찾아 field_map 후보 생성
        empty_cells = []
        for (r, c), text in sorted(cell_map.items()):
            if not text.strip():
                # 왼쪽 또는 위쪽 셀의 텍스트를 라벨로 사용
                label = cell_map.get((r, c - 1), '') or cell_map.get((r - 1, c), '')
                label = label.strip().replace('\n', ' ')[:30]
                empty_cells.append({
                    'row': r,
                    'col': c,
                    'label_hint': label if label else f'row{r}_col{c}',
                })

        # 빈 셀이 있으면 단일 entity_block으로 구성
        if empty_cells:
            fields = []
            for ec in empty_cells:
                fields.append({
                    'offset': ec['row'],
                    'left': {'col': ec['col'], 'field': ec['label_hint']},
                })
            entity_blocks.append({
                'name': 'cover_fields',
                'data_path': 'cover_fields',
                'start_row': 0,
                'fields': fields,
            })

    # template.json 생성
    template_name = Path(hwpx_path).stem
    template_config = {
        'name': template_name,
        'description': f'{template_name} 템플릿 (자동 생성 — 검토 필요)',
        'cover_table_index': cover_idx,
        'replacements': [],
    }

    template_path = os.path.join(output_dir, 'template.json')
    with open(template_path, 'w', encoding='utf-8') as f:
        json.dump(template_config, f, ensure_ascii=False, indent=2)

    # field_map.json 생성
    field_map = {
        'entity_blocks': entity_blocks,
        'company_lists': [],
    }

    field_map_path = os.path.join(output_dir, 'field_map.json')
    with open(field_map_path, 'w', encoding='utf-8') as f:
        json.dump(field_map, f, ensure_ascii=False, indent=2)

    return {
        'template_json': template_path,
        'field_map_json': field_map_path,
        'total_tables': len(all_tables),
        'cover_table_index': cover_idx,
        'cover_table_size': f'{cover_tbl.get("rowCnt", "?")}x{cover_tbl.get("colCnt", "?")}' if cover_tbl is not None else 'N/A',
        'empty_cells_found': len(empty_cells) if cover_tbl is not None else 0,
    }


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='HWPX 참조 문서 구조 추출',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  # 문서 구조 분석
  python3 src/extract_template.py --hwpx ref/test_01.hwpx --tables

  # 템플릿 설정 초안 자동 생성
  python3 src/extract_template.py --hwpx ref/새양식.hwpx --generate-template-config -o templates/새양식/
        """,
    )
    parser.add_argument('--hwpx', default=str(HWPX_PATH),
                        help='HWPX 파일 경로 (기본: ref/test_01.hwpx)')
    parser.add_argument('--cover', action='store_true',
                        help='커버 페이지만 추출')
    parser.add_argument('--sections', action='store_true',
                        help='본문 섹션만 추출')
    parser.add_argument('--styles', action='store_true',
                        help='스타일 정보만 추출')
    parser.add_argument('--tables', action='store_true',
                        help='주요 표 추출')
    parser.add_argument('--all-tables', action='store_true',
                        help='모든 표 요약 목록 출력')
    parser.add_argument('--sample-data', action='store_true',
                        help='sample_input.json 생성')
    parser.add_argument('--generate-template-config', action='store_true',
                        help='template.json + field_map.json 초안 자동 생성')
    parser.add_argument('--output', '-o', default=None,
                        help='출력 파일/디렉토리 경로 (기본: stdout)')

    args = parser.parse_args()
    hwpx = args.hwpx

    if args.generate_template_config:
        output_dir = args.output or '.'
        result = generate_template_config(hwpx, output_dir)
        print(f"템플릿 설정 파일 생성 완료:", file=sys.stderr)
        print(f"  template.json: {result['template_json']}", file=sys.stderr)
        print(f"  field_map.json: {result['field_map_json']}", file=sys.stderr)
        print(f"  표 개수: {result['total_tables']}", file=sys.stderr)
        print(f"  커버 테이블: index={result['cover_table_index']}, size={result['cover_table_size']}", file=sys.stderr)
        print(f"  빈 셀 발견: {result['empty_cells_found']}개", file=sys.stderr)
        print(f"\n※ 생성된 파일을 수동으로 검토하고 보정하세요.", file=sys.stderr)
        return

    if args.cover:
        result = extract_cover_table(hwpx)
    elif args.sections:
        result = extract_body_sections(hwpx)
    elif args.styles:
        result = extract_styles(hwpx)
    elif args.tables:
        result = extract_key_tables(hwpx)
    elif args.all_tables:
        result = extract_all_tables(hwpx)
    elif args.sample_data:
        result = generate_sample_data(hwpx)
    else:
        result = extract_document_structure(hwpx)

    output = json.dumps(result, ensure_ascii=False, indent=2)

    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"결과를 {args.output}에 저장했습니다.", file=sys.stderr)
    else:
        print(output)


if __name__ == '__main__':
    main()
