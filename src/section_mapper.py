"""섹션 매핑 — MD 섹션을 양식 위치에 매핑하고 콘텐츠를 추출한다.

form_content_map.json을 기반으로 business_plan_v2.md의 각 섹션을
양식 테이블 + 서술 영역에 매핑한다.

Usage:
    from section_mapper import SectionMapper
    mapper = SectionMapper('data/form_content_map.json', blocks)
    cover_data = mapper.get_cover_data()
    summary_data = mapper.get_summary_data()
    section_ops = mapper.get_section_ops('1')
"""

import json
from pathlib import Path

from src.md_parser import (
    parse_markdown, parse_table, strip_markdown,
    extract_section_blocks, get_all_sections, detect_section,
    BLOCK_HEADER, BLOCK_PARAGRAPH, BLOCK_TABLE, BLOCK_LIST,
    BLOCK_BLOCKQUOTE, BLOCK_CODE,
)


class SectionMapper:
    """MD 콘텐츠를 양식 위치에 매핑하는 매퍼."""

    def __init__(self, config_path, blocks):
        """
        Args:
            config_path: form_content_map.json 경로
            blocks: parse_markdown() 결과
        """
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        self.blocks = blocks
        self._sections = get_all_sections(blocks)

    def get_cover_data(self):
        """표지(T0)에 채울 데이터를 반환한다.

        Returns:
            dict: {field_name: value}
        """
        data = dict(self.config.get('cover_data', {}))
        # MD에서 추가 데이터 추출 가능 (현재는 고정값 사용)
        return data

    def get_summary_data(self):
        """과제요약서(T2)에 채울 데이터를 반환한다.

        Returns:
            dict: {field_name: summarized_text}
        """
        summary_blocks = self._sections.get('summary', [])
        if not summary_blocks:
            return {}

        result = {}

        # 각 H2 헤더 아래의 텍스트를 수집
        current_key = None
        current_texts = []

        def flush():
            nonlocal current_key, current_texts
            if current_key and current_texts:
                result[current_key] = '\n'.join(current_texts)
            current_key = None
            current_texts = []

        for block in summary_blocks:
            if block['type'] == BLOCK_HEADER and block.get('level') == 2:
                flush()
                text = block['text']
                # H2 텍스트에서 필드 키 매칭
                if '과 제 명' in text or '과제명' in text:
                    current_key = '과제명'
                elif '과제요약' in text:
                    current_key = '과제요약'
                elif '사업기간' in text:
                    current_key = '사업기간'
                elif '산업분야' in text:
                    current_key = '산업분야'
                elif '사업비' in text and '사 업 비' in text or '사업비' in text:
                    current_key = '사업비'
                elif '과제 목표' in text or '과제목표' in text:
                    current_key = '과제목표'
                elif '개발내용' in text:
                    current_key = '개발내용'
                elif '수행 방법' in text or '과제 수행' in text:
                    current_key = '수행방법'
                elif '사업화전략' in text or '사업화' in text:
                    current_key = '사업화전략'
                elif '최종결과물' in text:
                    current_key = '최종결과물'
                elif '기대효과' in text:
                    current_key = '기대효과'
                continue

            if current_key:
                if block['type'] == BLOCK_PARAGRAPH:
                    current_texts.append(strip_markdown(block['text']))
                elif block['type'] == BLOCK_LIST:
                    for item in block['items']:
                        current_texts.append('• ' + strip_markdown(item['text']))
                elif block['type'] == BLOCK_TABLE:
                    rows = parse_table(block['lines'])
                    for row in rows:
                        current_texts.append(' | '.join(
                            strip_markdown(c) for c in row
                        ))
                elif block['type'] == BLOCK_BLOCKQUOTE:
                    current_texts.append(strip_markdown(block['content']))

        flush()
        return result

    def get_section_blocks(self, section_id):
        """특정 섹션의 블록을 반환한다.

        Args:
            section_id: '1'~'7' 또는 'appendix'

        Returns:
            list[dict]: 블록 리스트
        """
        return self._sections.get(section_id, [])

    def get_narrative_config(self):
        """서술 섹션 설정을 반환한다.

        Returns:
            list[dict]: narrative_sections 설정
        """
        return self.config.get('narrative_sections', [])

    def get_fixed_values(self):
        """고정값 설정을 반환한다.

        Returns:
            dict: {table_key: text}
        """
        return self.config.get('fixed_values', {})

    def get_researcher_data(self):
        """참여연구원 테이블 데이터를 MD에서 추출한다.

        Returns:
            list[list[str]]: 파싱된 연구원 테이블 행 (헤더 제외)
        """
        sec4_blocks = self.get_section_blocks('4')

        for block in sec4_blocks:
            if block['type'] == BLOCK_TABLE:
                header = block['lines'][0] if block['lines'] else ''
                if '구분' in header and '성명' in header and '참여율' in header:
                    rows = parse_table(block['lines'])
                    return rows[1:] if len(rows) > 1 else []

        return []

    def get_kpi_data(self):
        """성능목표(KPI) 데이터를 MD에서 추출한다.

        Returns:
            list[dict]: [{세부목표, 목표성능, 측정방법}, ...]
        """
        sec3_blocks = self.get_section_blocks('3')

        for block in sec3_blocks:
            if block['type'] == BLOCK_TABLE:
                header = block['lines'][0] if block['lines'] else ''
                if '세부 목표' in header and '목표 성능' in header:
                    rows = parse_table(block['lines'])
                    results = []
                    for row in rows[1:]:
                        results.append({
                            '구분': strip_markdown(row[0]) if len(row) > 0 else '',
                            '세부목표': strip_markdown(row[1]) if len(row) > 1 else '',
                            '목표성능': strip_markdown(row[2]) if len(row) > 2 else '',
                            '측정방법': strip_markdown(row[3]) if len(row) > 3 else '',
                        })
                    return results
        return []

    def get_production_data(self):
        """생산계획 데이터를 MD에서 추출한다.

        Returns:
            dict: 생산계획 관련 데이터
        """
        sec5_blocks = self.get_section_blocks('5')
        # 생산계획 테이블에서 추출
        for block in sec5_blocks:
            if block['type'] == BLOCK_TABLE:
                header = block['lines'][0] if block['lines'] else ''
                if '판매량' in header or '매출' in header or '생산' in header:
                    rows = parse_table(block['lines'])
                    return rows
        return []

    def get_institution_data(self):
        """기관현황(T21) 데이터를 MD에서 추출한다.

        Returns:
            dict: {field: value}
        """
        sec7_blocks = self.get_section_blocks('7')
        data = {}

        for block in sec7_blocks:
            if block['type'] == BLOCK_TABLE:
                rows = parse_table(block['lines'])
                for row in rows:
                    if len(row) >= 2:
                        key = strip_markdown(row[0]).strip()
                        val = strip_markdown(row[1]).strip()
                        if key and val:
                            data[key] = val

        return data

    def truncate_text(self, text, max_chars=500):
        """긴 텍스트를 양식에 맞게 축약한다."""
        if len(text) <= max_chars:
            return text
        return text[:max_chars - 3] + '...'
