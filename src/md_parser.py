"""마크다운 파서 — 블록 단위 파싱 및 인라인 서식 처리.

generate_docx.py의 parse_markdown(), parse_table() 로직을 포팅하여
출력 포맷에 독립적인 순수 파서 모듈로 재구성.

Usage:
    from md_parser import parse_markdown, parse_table, extract_section_blocks

    blocks = parse_markdown(md_text)
    rows = parse_table(table_block['lines'])
    section_blocks = extract_section_blocks(blocks, '1. 과제의 필요성', '2. 관련 현황')
"""

import re
from pathlib import Path


# ── Block types ───────────────────────────────────────────────
BLOCK_CODE = 'code'
BLOCK_TABLE = 'table'
BLOCK_HEADER = 'header'
BLOCK_BLOCKQUOTE = 'blockquote'
BLOCK_LIST = 'list'
BLOCK_HR = 'hr'
BLOCK_PARAGRAPH = 'paragraph'


def parse_markdown(text):
    """마크다운 텍스트를 블록 리스트로 파싱한다.

    블록 타입:
      - code: {'type': 'code', 'content': str}
      - table: {'type': 'table', 'lines': [str]}
      - header: {'type': 'header', 'level': int, 'text': str}
      - blockquote: {'type': 'blockquote', 'content': str}
      - list: {'type': 'list', 'items': [{'indent': int, 'text': str}]}
      - hr: {'type': 'hr'}
      - paragraph: {'type': 'paragraph', 'text': str}

    Returns:
        list[dict]: 파싱된 블록 리스트
    """
    lines = text.split('\n')
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Code block
        if line.strip().startswith('```'):
            code_lines = []
            lang = line.strip()[3:].strip()
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            blocks.append({
                'type': BLOCK_CODE,
                'content': '\n'.join(code_lines),
                'lang': lang,
            })
            i += 1
            continue

        # Table
        if '|' in line and i + 1 < len(lines) and '---' in lines[i + 1]:
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            blocks.append({'type': BLOCK_TABLE, 'lines': table_lines})
            continue

        # Header
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            blocks.append({
                'type': BLOCK_HEADER,
                'level': level,
                'text': m.group(2),
            })
            i += 1
            continue

        # Blockquote
        if line.strip().startswith('>'):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                quote_lines.append(lines[i].strip().lstrip('>').strip())
                i += 1
            blocks.append({
                'type': BLOCK_BLOCKQUOTE,
                'content': '\n'.join(quote_lines),
            })
            continue

        # Bullet list
        if re.match(r'^[\s]*[-*]\s', line):
            list_items = []
            while i < len(lines) and re.match(r'^[\s]*[-*]\s', lines[i]):
                indent = len(lines[i]) - len(lines[i].lstrip())
                item_text = re.sub(r'^[\s]*[-*]\s', '', lines[i])
                list_items.append({'indent': indent, 'text': item_text})
                i += 1
            blocks.append({'type': BLOCK_LIST, 'items': list_items})
            continue

        # Horizontal rule
        if re.match(r'^---+$', line.strip()):
            blocks.append({'type': BLOCK_HR})
            i += 1
            continue

        # Empty line
        if line.strip() == '':
            i += 1
            continue

        # Skip HTML tags
        if line.strip().startswith('<'):
            i += 1
            continue

        # Regular paragraph (merge consecutive non-empty lines)
        para_lines = [line]
        i += 1
        while (i < len(lines) and lines[i].strip()
               and not lines[i].strip().startswith('#')
               and not lines[i].strip().startswith('|')
               and not lines[i].strip().startswith('```')
               and not lines[i].strip().startswith('>')
               and not re.match(r'^[\s]*[-*]\s', lines[i])
               and not re.match(r'^---+$', lines[i].strip())
               and not lines[i].strip().startswith('<')):
            para_lines.append(lines[i])
            i += 1
        blocks.append({
            'type': BLOCK_PARAGRAPH,
            'text': ' '.join(para_lines),
        })

    return blocks


def parse_table(table_lines):
    """마크다운 테이블 줄을 2D 리스트로 파싱한다.

    Args:
        table_lines: '|' 구분 테이블 줄 리스트

    Returns:
        list[list[str]]: [[cell, cell, ...], ...] 헤더 포함, 구분선 제외
    """
    rows = []
    for i, line in enumerate(table_lines):
        if i == 1 and '---' in line:
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        rows.append(cells)
    return rows


def parse_inline(text):
    """마크다운 인라인 서식을 파싱하여 run 리스트로 반환한다.

    Returns:
        list[dict]: [{'text': str, 'bold': bool, 'italic': bool}, ...]
    """
    parts = re.split(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*)', text)
    runs = []
    for part in parts:
        if not part:
            continue
        if part.startswith('***') and part.endswith('***'):
            runs.append({'text': part[3:-3], 'bold': True, 'italic': True})
        elif part.startswith('**') and part.endswith('**'):
            runs.append({'text': part[2:-2], 'bold': True, 'italic': False})
        elif part.startswith('*') and part.endswith('*'):
            runs.append({'text': part[1:-1], 'bold': False, 'italic': True})
        else:
            runs.append({'text': part, 'bold': False, 'italic': False})
    return runs


def strip_markdown(text):
    """마크다운 인라인 서식을 제거하여 일반 텍스트로 반환한다."""
    return re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', text)


# ── Section extraction ────────────────────────────────────────

# H1 헤더 텍스트에서 섹션 번호 감지
SECTION_PATTERNS = [
    ('과제 요약서', 'summary'),
    ('과제의 필요성', '1'),
    ('관련 현황', '2'),
    ('과제의 목표', '3'),
    ('추진 방법', '4'),
    ('사업화', '5'),     # '사업화 전략' or '사업화' in text
    ('사업비', '6'),
    ('사업수행기관', '7'),
    ('부속서류', 'appendix'),
]


def detect_section(header_text):
    """H1 헤더 텍스트에서 섹션 ID를 반환한다.

    Returns:
        str or None: 'summary', '1'~'7', 'appendix', or None
    """
    for keyword, section_id in SECTION_PATTERNS:
        if keyword in header_text:
            # '사업화' is in '사업화 전략' and '사업비' — check order matters
            if keyword == '사업화' and '사업비' in header_text:
                continue
            return section_id
    return None


def extract_section_blocks(blocks, start_section_id, end_section_id=None):
    """특정 섹션의 블록들을 추출한다.

    Args:
        blocks: parse_markdown() 결과
        start_section_id: 시작 섹션 ID ('1', '2', ..., 'summary', 'appendix')
        end_section_id: 종료 섹션 ID (None이면 끝까지)

    Returns:
        list[dict]: 해당 섹션의 블록 리스트 (시작 H1 헤더 제외)
    """
    result = []
    capturing = False

    for block in blocks:
        if block['type'] == BLOCK_HEADER and block.get('level') == 1:
            section = detect_section(block['text'])
            if section == start_section_id:
                capturing = True
                continue
            if capturing and end_section_id and section == end_section_id:
                break
            if capturing and section is not None and section != start_section_id:
                break

        if capturing:
            result.append(block)

    return result


def get_all_sections(blocks):
    """전체 블록에서 섹션별로 분리한다.

    Returns:
        dict: {section_id: [blocks]}
    """
    sections = {}
    current_section = None

    for block in blocks:
        if block['type'] == BLOCK_HEADER and block.get('level') == 1:
            detected = detect_section(block['text'])
            if detected:
                current_section = detected
                if current_section not in sections:
                    sections[current_section] = []
                continue

        if current_section:
            if current_section not in sections:
                sections[current_section] = []
            sections[current_section].append(block)

    return sections


def load_and_parse(md_path):
    """마크다운 파일을 읽고 파싱한다.

    Args:
        md_path: 마크다운 파일 경로

    Returns:
        list[dict]: 파싱된 블록 리스트
    """
    text = Path(md_path).read_text(encoding='utf-8')
    return parse_markdown(text)
