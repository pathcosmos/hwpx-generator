"""마크다운 블록 → COM 오퍼레이션 컴파일러.

md_parser.py의 블록 리스트를 bridge.py가 실행할 수 있는
JSON 오퍼레이션 리스트로 변환한다.

Usage:
    from md_parser import parse_markdown
    from md_to_ops import compile_blocks_to_ops

    blocks = parse_markdown(md_text)
    ops = compile_blocks_to_ops(blocks)
    # ops → bridge.create_document() 또는 bridge.fill_template()에 전달
"""

import re
from src.md_parser import parse_table, parse_inline, strip_markdown

# ── 색상 상수 (BGR 변환은 COM 측에서 처리) ─────────────────
COLOR_BLACK = 0x000000
COLOR_GREEN = 0x003070    # #007030 → BGR: 0x300070
COLOR_WHITE = 0xFFFFFF
COLOR_GRAY = 0xD9D9D9
COLOR_LIGHT_GREEN = 0xE9F5E8  # #E8F5E9 → BGR
COLOR_DARK_TEXT = 0x333333

# ── 폰트 상수 ────────────────────────────────────────────
FONT_BODY = '바탕'
FONT_HEADER = '굴림'
FONT_MONO = 'Consolas'

# ── HWP 단위 ─────────────────────────────────────────────
MM = 283.46
PT_TO_HWPUNIT = 100  # 1pt = 100 HwpUnit (for char size)

# heading level → 왼쪽 들여쓰기 (mm)
# 양식 계층: 가.(H2)=0mm, 1)(H3)=5mm, (1)(H4)=10mm, 가)(H5)=14mm, (가)(H6)=17mm
HEADING_INDENT_MM = {1: 0, 2: 0, 3: 5, 4: 10, 5: 14, 6: 17}


def _char_op(font=None, size=None, bold=None, italic=None, color=None,
             underline=None):
    """글자 서식 오퍼레이션 생성."""
    op = {'op': 'set_char_shape'}
    if font is not None:
        op['font'] = font
    if size is not None:
        op['size'] = size
    if bold is not None:
        op['bold'] = bold
    if italic is not None:
        op['italic'] = italic
    if color is not None:
        op['color'] = color
    if underline is not None:
        op['underline'] = underline
    return op


def _para_op(align=None, line_spacing=None, space_before=None,
             space_after=None, indent_left=None, first_line_indent=None):
    """문단 서식 오퍼레이션 생성."""
    op = {'op': 'set_para_shape'}
    if align is not None:
        op['align'] = align
    if line_spacing is not None:
        op['line_spacing'] = line_spacing
    if space_before is not None:
        op['space_before'] = space_before
    if space_after is not None:
        op['space_after'] = space_after
    if indent_left is not None:
        op['indent_left'] = indent_left
    if first_line_indent is not None:
        op['first_line_indent'] = first_line_indent
    return op


def _text_op(text):
    """텍스트 삽입 오퍼레이션 생성."""
    return {'op': 'insert_text', 'text': text}


def _break_op():
    """줄바꿈(문단 나누기) 오퍼레이션."""
    return {'op': 'line_break'}


def _page_break_op():
    """페이지 나누기 오퍼레이션."""
    return {'op': 'page_break'}


def _table_op(rows, cols):
    """테이블 삽입 오퍼레이션."""
    return {'op': 'insert_table', 'rows': rows, 'cols': cols}


def _fill_table_op(data):
    """테이블 데이터 채우기 오퍼레이션."""
    return {'op': 'fill_table', 'data': data}


def _cell_bg_op(r, g, b):
    """셀 배경색 오퍼레이션."""
    return {'op': 'set_cell_background', 'r': r, 'g': g, 'b': b}


# ── 블록별 컴파일러 ──────────────────────────────────────

def compile_header(block):
    """헤더 블록 → 오퍼레이션 리스트."""
    level = block['level']
    text = block['text']
    ops = []

    size_map = {1: 14, 2: 12, 3: 11, 4: 10, 5: 10, 6: 9}
    font = FONT_HEADER if level <= 3 else FONT_BODY
    size = size_map.get(level, 9)

    space_before_map = {1: int(18 * MM / 10), 2: int(12 * MM / 10),
                        3: int(8 * MM / 10), 4: int(6 * MM / 10),
                        5: int(5 * MM / 10), 6: int(4 * MM / 10)}
    space_after_map = {1: int(12 * MM / 10), 2: int(8 * MM / 10),
                       3: int(6 * MM / 10), 4: int(4 * MM / 10),
                       5: int(3 * MM / 10), 6: int(2 * MM / 10)}

    indent_left = int(HEADING_INDENT_MM.get(level, 0) * MM)

    ops.append(_para_op(
        align='left',
        line_spacing=130 if level <= 2 else 120,
        space_before=space_before_map.get(level, 0),
        space_after=space_after_map.get(level, 0),
        indent_left=indent_left,
        first_line_indent=0,
    ))
    ops.append(_char_op(font=font, size=size, bold=True, italic=False,
                        color=COLOR_BLACK))
    ops.append(_text_op(strip_markdown(text)))
    ops.append(_break_op())

    return ops


def compile_paragraph(block, heading_level=2):
    """본문 단락 블록 → 오퍼레이션 리스트."""
    text = block['text']
    ops = []

    indent_left = int(HEADING_INDENT_MM.get(heading_level, 0) * MM)

    ops.append(_para_op(
        align='left',
        line_spacing=130,
        space_before=int(2 * MM / 10),
        space_after=int(4 * MM / 10),
        indent_left=indent_left,
        first_line_indent=int(5 * MM),
    ))

    # 인라인 서식 처리
    runs = parse_inline(text)
    for run in runs:
        ops.append(_char_op(
            font=FONT_BODY, size=10,
            bold=run['bold'], italic=run['italic'],
            color=COLOR_BLACK,
        ))
        ops.append(_text_op(run['text']))

    ops.append(_break_op())
    return ops


def compile_table(block):
    """테이블 블록 → 오퍼레이션 리스트."""
    rows_data = parse_table(block['lines'])
    if not rows_data or len(rows_data) < 2:
        return []

    n_rows = len(rows_data)
    n_cols = len(rows_data[0])

    ops = []
    # 문단 서식 리셋
    ops.append(_para_op(align='left', line_spacing=100,
                        space_before=0, space_after=0, first_line_indent=0))

    # 테이블 삽입
    ops.append(_table_op(n_rows, n_cols))

    # 데이터 채우기 — 마크다운 **bold** 제거
    clean_data = []
    for row in rows_data:
        clean_row = [strip_markdown(cell) for cell in row]
        clean_data.append(clean_row)

    ops.append(_fill_table_op(clean_data))
    ops.append(_break_op())

    return ops


def compile_list(block, heading_level=2):
    """리스트 블록 → 오퍼레이션 리스트."""
    ops = []
    base_mm = HEADING_INDENT_MM.get(heading_level, 0)
    for item in block['items']:
        indent_level = item['indent'] // 2
        left_indent = int((base_mm + 5 + indent_level * 5) * MM)
        prefix = '• ' if indent_level == 0 else '- '

        ops.append(_para_op(
            align='left',
            line_spacing=120,
            space_before=int(1 * MM / 10),
            space_after=int(1 * MM / 10),
            indent_left=left_indent,
            first_line_indent=int(-3 * MM),
        ))

        runs = parse_inline(item['text'])
        # 접두사
        ops.append(_char_op(font=FONT_BODY, size=10, bold=False, color=COLOR_BLACK))
        ops.append(_text_op(prefix))
        # 내용
        for run in runs:
            ops.append(_char_op(
                font=FONT_BODY, size=10,
                bold=run['bold'], italic=run['italic'],
                color=COLOR_BLACK,
            ))
            ops.append(_text_op(run['text']))
        ops.append(_break_op())

    return ops


def compile_blockquote(block, heading_level=2):
    """인용문 블록 → 오퍼레이션 리스트."""
    ops = []
    base_indent = int(HEADING_INDENT_MM.get(heading_level, 0) * MM)
    ops.append(_para_op(
        align='left',
        line_spacing=120,
        space_before=int(4 * MM / 10),
        space_after=int(4 * MM / 10),
        indent_left=base_indent + int(10 * MM),
        first_line_indent=0,
    ))
    ops.append(_char_op(font=FONT_BODY, size=9, bold=False, italic=False,
                        color=0x444444))
    ops.append(_text_op('▎ ' + block['content']))
    ops.append(_break_op())
    return ops


def compile_code(block, heading_level=2):
    """코드 블록 → 오퍼레이션 리스트."""
    ops = []
    base_indent = int(HEADING_INDENT_MM.get(heading_level, 0) * MM)
    ops.append(_para_op(
        align='left',
        line_spacing=100,
        space_before=int(2 * MM / 10),
        space_after=int(2 * MM / 10),
        indent_left=base_indent + int(5 * MM),
        first_line_indent=0,
    ))
    ops.append(_char_op(font=FONT_MONO, size=7, bold=False, color=COLOR_DARK_TEXT))
    ops.append(_text_op(block['content']))
    ops.append(_break_op())
    return ops


def compile_table_as_text(block, heading_level=2):
    """테이블 → 텍스트 오퍼레이션 (COM insert_table 실패 대비)."""
    rows_data = parse_table(block['lines'])
    if not rows_data or len(rows_data) < 2:
        return []

    ops = []
    indent_left = int(HEADING_INDENT_MM.get(heading_level, 0) * MM)

    # 헤더 행: [Col1 / Col2 / Col3]
    headers = [strip_markdown(h) for h in rows_data[0]]
    ops.append(_para_op(align='left', line_spacing=110,
                        space_before=int(4 * MM / 10), space_after=int(2 * MM / 10),
                        indent_left=indent_left, first_line_indent=0))
    ops.append(_char_op(font=FONT_BODY, size=9, bold=True, color=COLOR_DARK_TEXT))
    ops.append(_text_op('[' + ' / '.join(headers) + ']'))
    ops.append(_break_op())

    # 데이터 행: ■ Label: value -- value
    for row in rows_data[1:]:
        cells = [strip_markdown(c) for c in row]
        if len(cells) >= 2:
            row_text = f'  \u25a0 {cells[0]}: {" -- ".join(cells[1:])}'
        else:
            row_text = f'  \u25a0 {cells[0]}' if cells else ''
        ops.append(_para_op(align='left', line_spacing=110,
                            space_before=int(1 * MM / 10), space_after=int(1 * MM / 10),
                            indent_left=indent_left, first_line_indent=0))
        ops.append(_char_op(font=FONT_BODY, size=9, bold=False, color=COLOR_BLACK))
        ops.append(_text_op(row_text))
        ops.append(_break_op())

    return ops


# ── 커스텀 테이블 감지 및 컴파일 ─────────────────────────

def _is_gantt(table_lines):
    header = table_lines[0] if table_lines else ''
    return '활동 구분' in header and 'M1' in header


def _is_researcher(table_lines):
    header = table_lines[0] if table_lines else ''
    return ('구분' in header and '성명' in header
            and '담당' in header and '참여율' in header)


def _is_common_goals(table_lines):
    header = table_lines[0] if table_lines else ''
    return '항목' in header and '목표' in header and '구체적 계획' in header


def _is_budget(table_lines):
    header = table_lines[0] if table_lines else ''
    return '비목' in header and '세목' in header and '금액' in header


def compile_custom_table(block, current_section=''):
    """커스텀 테이블을 감지하고 해당하면 오퍼레이션을 반환한다.

    Returns:
        list or None: 오퍼레이션 리스트 (커스텀이 아니면 None)
    """
    lines = block['lines']

    if _is_gantt(lines):
        return compile_gantt_as_text(lines)
    if _is_researcher(lines):
        return compile_researcher_as_text(lines)
    if _is_common_goals(lines) or _is_budget(lines):
        return compile_table_as_text(block)
    if current_section == '7':
        return compile_table_as_text(block)

    return None


def compile_gantt_as_text(table_lines):
    """간트차트 테이블 → 텍스트 오퍼레이션 (COM insert_table 대체)."""
    rows_data = parse_table(table_lines)
    if not rows_data:
        return []

    ops = []

    # 헤더 행
    headers = [strip_markdown(h) for h in rows_data[0]]
    ops.append(_para_op(align='left', line_spacing=110,
                        space_before=int(4 * MM / 10), space_after=int(2 * MM / 10),
                        indent_left=0, first_line_indent=0))
    ops.append(_char_op(font=FONT_BODY, size=9, bold=True, color=COLOR_DARK_TEXT))
    ops.append(_text_op('[' + ' / '.join(headers) + ']'))
    ops.append(_break_op())

    # 데이터 행: 활동명 + 월별 기호
    for row in rows_data[1:]:
        cells = []
        for ci, cell in enumerate(row):
            clean = strip_markdown(cell).strip()
            if ci >= 2:
                if '====' in clean:
                    clean = '\u25a0'
                elif '==' in clean:
                    clean = '\u25a1'
                elif '----' in clean:
                    clean = '\u00b7'
                elif '[A]' in clean:
                    clean = 'A'
            cells.append(clean)
        if len(cells) >= 2:
            activity = cells[0]
            months = ' '.join(cells[2:]) if len(cells) > 2 else ''
            row_text = f'  \u25a0 {activity} ({cells[1]}): {months}'
        else:
            row_text = f'  \u25a0 {cells[0]}' if cells else ''
        ops.append(_para_op(align='left', line_spacing=110,
                            space_before=int(1 * MM / 10), space_after=int(1 * MM / 10),
                            indent_left=0, first_line_indent=0))
        ops.append(_char_op(font=FONT_BODY, size=9, bold=False, color=COLOR_BLACK))
        ops.append(_text_op(row_text))
        ops.append(_break_op())

    # 범례
    ops.append(_para_op(align='left', line_spacing=110,
                        space_before=int(2 * MM / 10), space_after=int(4 * MM / 10),
                        indent_left=0, first_line_indent=0))
    ops.append(_char_op(font=FONT_BODY, size=8, bold=True, color=COLOR_BLACK))
    ops.append(_text_op('범례: ■ 주요활동  □ 착수/마무리  · 지속활동  A 기술자문위'))
    ops.append(_break_op())

    return ops


def compile_researcher_as_text(table_lines):
    """참여연구원 테이블 → 텍스트 오퍼레이션 (COM insert_table 대체)."""
    rows_data = parse_table(table_lines)
    if not rows_data or len(rows_data) < 2:
        return []

    ops = []

    # 헤더
    ops.append(_para_op(align='left', line_spacing=110,
                        space_before=int(4 * MM / 10), space_after=int(2 * MM / 10),
                        indent_left=0, first_line_indent=0))
    ops.append(_char_op(font=FONT_BODY, size=9, bold=True, color=COLOR_DARK_TEXT))
    ops.append(_text_op('[번호 / 소속 / 성명 / 직위 / 전공 / 담당분야 / 참여율 / 인력구분]'))
    ops.append(_break_op())

    # 데이터 행
    for ri, row in enumerate(rows_data[1:], start=1):
        md_구분 = strip_markdown(row[0]) if len(row) > 0 else ''
        md_성명 = strip_markdown(row[1]) if len(row) > 1 else ''
        md_직급 = strip_markdown(row[2]) if len(row) > 2 else ''
        md_전문 = strip_markdown(row[3]) if len(row) > 3 else ''
        md_담당 = strip_markdown(row[4]) if len(row) > 4 else ''
        md_참여율 = strip_markdown(row[5]).replace('%', '') if len(row) > 5 else ''

        row_text = (f'  ■ {ri}. {md_성명} ({md_직급}) — 동연에스엔티 / '
                    f'전공: {md_전문} / 담당: {md_담당} / '
                    f'참여율: {md_참여율}% / {md_구분.split("(")[0].strip()}')
        ops.append(_para_op(align='left', line_spacing=110,
                            space_before=int(1 * MM / 10), space_after=int(1 * MM / 10),
                            indent_left=0, first_line_indent=0))
        ops.append(_char_op(font=FONT_BODY, size=9, bold=False, color=COLOR_BLACK))
        ops.append(_text_op(row_text))
        ops.append(_break_op())

    return ops


def compile_researcher_table(table_lines):
    """참여연구원 6열 → 13열 확장 오퍼레이션 (레거시, 미사용)."""
    return compile_researcher_as_text(table_lines)


# ── 메인 컴파일러 ────────────────────────────────────────

# 페이지 브레이크 삽입 섹션 키워드
SECTION_BREAK_KEYWORDS = [
    '2. 관련 현황',
    '3. 과제의 목표 및 내용',
    '4. 추진 방법, 전략 및 체계',
    '5. 사업화 전략',
    '6. 사업비 비목별 세부 내역',
    '7. 사업수행기관 기본현황',
    '부속서류',
]

SKIP_H1_BREAK = ['MES 작업지시', 'MES 내부 교차']


def compile_blocks_to_ops(blocks, current_section=''):
    """블록 리스트를 COM 오퍼레이션 리스트로 컴파일한다.

    Args:
        blocks: parse_markdown() 결과
        current_section: 현재 섹션 ID (커스텀 테이블 감지용)

    Returns:
        list[dict]: COM 오퍼레이션 리스트
    """
    ops = []
    current_heading_level = 2  # 기본값: H2 수준

    for block in blocks:
        btype = block['type']

        if btype == 'header':
            current_heading_level = block['level']
            ops.extend(compile_header(block))

        elif btype == 'paragraph':
            ops.extend(compile_paragraph(block, heading_level=current_heading_level))

        elif btype == 'table':
            custom_ops = compile_custom_table(block, current_section)
            if custom_ops is not None:
                ops.extend(custom_ops)
            else:
                ops.extend(compile_table_as_text(block,
                                                 heading_level=current_heading_level))

        elif btype == 'list':
            ops.extend(compile_list(block, heading_level=current_heading_level))

        elif btype == 'blockquote':
            ops.extend(compile_blockquote(block, heading_level=current_heading_level))

        elif btype == 'code':
            ops.extend(compile_code(block, heading_level=current_heading_level))

        elif btype == 'hr':
            pass  # Skip horizontal rules

    return ops


def compile_section_ops(blocks, section_id):
    """특정 섹션의 블록을 COM 오퍼레이션으로 컴파일한다.

    Args:
        blocks: 해당 섹션의 블록 리스트 (extract_section_blocks 결과)
        section_id: 섹션 ID ('1'~'7')

    Returns:
        list[dict]: COM 오퍼레이션 리스트
    """
    return compile_blocks_to_ops(blocks, current_section=section_id)
