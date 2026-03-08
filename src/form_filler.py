"""양식 자동 채우기 파이프라인 — 메인 오케스트레이터.

Two-Pass Hybrid Pipeline:
  Pass 1 (XML): 36개 양식 테이블의 빈 셀에 데이터 직접 삽입 + 마커 주입
  Pass 2 (COM): 마커 위치에 서술형 본문 삽입 + 신규 테이블 생성

Usage:
    python3 src/form_filler.py \\
        --template ../form_to_fillout.hwpx \\
        --md ../business_plan_v2.md \\
        --output output/filled \\
        [--pass1-only] [--pass2-only] [--no-pdf]
"""

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

# lxml은 WSL에서 사용할 수 있지만, 없을 수도 있으므로 fallback 제공
try:
    from lxml import etree
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

from src.md_parser import load_and_parse, strip_markdown, parse_table, get_all_sections
from src.section_mapper import SectionMapper
from src.md_to_ops import compile_section_ops

# lxml이 필요한 모듈은 조건부 임포트
EDITOR_AVAILABLE = False
try:
    from src.hwpx_editor import HwpxEditor
    EDITOR_AVAILABLE = True
except ImportError:
    pass

from src.bridge import (
    fill_template, open_and_save_as_pdf, fix_hwpx_for_pdf,
    delete_page_content, wsl_to_win_path,
)


def format_table_as_text(text):
    """마크다운 파이프 테이블을 구조화된 텍스트로 변환한다.

    '구분 | 결과물 | 규격' 형태의 테이블을
    '■ 구분: 결과물 — 규격' 형태의 읽기 쉬운 텍스트로 변환.

    Args:
        text: 마크다운 테이블이 포함될 수 있는 텍스트

    Returns:
        변환된 텍스트
    """
    lines = text.split('\n')
    result = []
    table_lines = []
    headers = None

    def flush_table():
        nonlocal headers, table_lines
        if not table_lines:
            return
        if headers and len(headers) >= 2:
            for row in table_lines:
                cols = [c.strip() for c in row.split('|')]
                cols = [c for c in cols if c]
                if len(cols) >= 2:
                    # 첫 컬럼: 구분, 나머지는 " — "로 연결
                    label = cols[0]
                    rest = ' — '.join(cols[1:])
                    result.append(f'  ■ {label}: {rest}')
                else:
                    result.append(row)
        else:
            result.extend(table_lines)
        headers = None
        table_lines = []

    for line in lines:
        stripped = line.strip()
        # 구분선 (---|---|---) 건너뛰기
        if stripped and all(c in '-| ' for c in stripped) and '|' in stripped:
            continue

        if '|' in stripped and stripped.count('|') >= 2:
            cols = [c.strip() for c in stripped.split('|')]
            cols = [c for c in cols if c]
            if headers is None:
                # 헤더 행
                headers = cols
                if len(headers) >= 2:
                    result.append(f'[{" / ".join(headers)}]')
            else:
                table_lines.append(stripped)
        else:
            flush_table()
            result.append(line)

    flush_table()
    return '\n'.join(result)


# ── 설정 경로 ─────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = PROJECT_ROOT / 'templates' / 'gyeongnam_rbd'
DATA_DIR = PROJECT_ROOT / 'data'


def load_field_map():
    """field_map.json 로드."""
    path = TEMPLATE_DIR / 'field_map.json'
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_content_map():
    """form_content_map.json 로드."""
    path = DATA_DIR / 'form_content_map.json'
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ── Pass 1: XML 테이블 셀 채우기 ──────────────────────────

def run_pass1(template_path, md_path, output_path):
    """Pass 1: XML 직접 수정으로 양식 테이블 셀을 채운다.

    Args:
        template_path: 원본 양식 HWPX 경로
        md_path: business_plan_v2.md 경로
        output_path: 출력 HWPX 경로

    Returns:
        str: 출력 파일 경로
    """
    if not EDITOR_AVAILABLE:
        print("[Pass 1] ERROR: lxml not available. Install with: pip3 install lxml")
        return None

    print("[Pass 1] Loading template and markdown...")
    field_map = load_field_map()
    content_map = load_content_map()

    # MD 파싱
    blocks = load_and_parse(md_path)
    mapper = SectionMapper(str(DATA_DIR / 'form_content_map.json'), blocks)

    # 템플릿 복사
    shutil.copy2(template_path, output_path)
    editor = HwpxEditor(output_path)

    # ── 0. 템플릿 메모(주석) 제거 ──
    memo_count = editor.remove_memos()
    print(f"[Pass 1] Removed {memo_count} template memos (annotations)")

    total_cells = 0
    tables_config = field_map.get('tables', {})

    # ── 1. 표지(T0) 채우기 ──
    print("[Pass 1] Filling cover table (T0)...")
    total_cells += fill_cover_table(editor, tables_config.get('T0_cover', {}), mapper)

    # ── 2. 과제요약서(T2) 채우기 ──
    print("[Pass 1] Filling summary table (T2)...")
    total_cells += fill_summary_table(editor, tables_config.get('T2_summary', {}), mapper)

    # ── 3. 시장규모(T4) ──
    print("[Pass 1] Filling market table (T4)...")
    total_cells += fill_market_table(editor, tables_config.get('T4_market_size', {}))

    # ── 3b. 수요처(T5) ──
    print("[Pass 1] Filling demand table (T5)...")
    total_cells += fill_demand_table(editor, tables_config.get('T5_demand', {}))

    # ── 3c. 공통목표(T6) — 자율목표 행 ──
    print("[Pass 1] Filling common goals table (T6)...")
    total_cells += fill_common_goals_table(editor, tables_config.get('T6_common_goals', {}))

    # ── 4. 성능목표(T8) — 빈 행 채우기 ──
    print("[Pass 1] Filling KPI table (T8)...")
    total_cells += fill_kpi_table(editor, tables_config.get('T8_kpi', {}))

    # ── 4. 참여연구원-주관(T12) 채우기 ──
    print("[Pass 1] Filling researcher table (T12)...")
    total_cells += fill_researcher_table(editor, tables_config.get('T12_researcher_main', {}),
                                          mapper)

    # ── 5. 참여연구원-공동(T13) — "해당없음" ──
    print("[Pass 1] Filling co-researcher table (T13) — 단독수행...")
    t13_config = tables_config.get('T13_researcher_co', {})
    if 'fixed_text' in t13_config:
        table = editor.get_table(t13_config['table_index'])
        if table is not None:
            fixed = t13_config['fixed_text']
            # 예시 데이터 행 비우기 + 첫 행에 고정값
            for row in range(2, 12):
                for col in range(12):
                    editor.set_cell_text(table, row, col, '')
            editor.set_cell_text(table, 2, 2, fixed)
            total_cells += 1

    # ── 6. 생산계획(T14) ──
    print("[Pass 1] Filling production table (T14)...")
    total_cells += fill_production_table(editor, tables_config.get('T14_production', {}))

    # ── 7. 사업비(T18) — placeholder ──
    print("[Pass 1] Filling budget table (T18)...")
    total_cells += fill_budget_table(editor, tables_config.get('T18_budget_main', {}))

    # ── 8. 사업비-공동(T20) — "해당없음" ──
    print("[Pass 1] Filling co-budget table (T20) — 단독수행...")
    t20_config = tables_config.get('T20_budget_co', {})
    if 'fixed_text' in t20_config:
        table = editor.get_table(t20_config.get('table_index', 20))
        if table is not None:
            fixed = t20_config['fixed_text']
            # 데이터 행 비우기 + 첫 데이터 행에 고정값
            for row in range(2, 14):
                for col in range(7):
                    editor.set_cell_text(table, row, col, '')
            editor.set_cell_text(table, 2, 2, fixed)
            total_cells += 1

    # ── 9. 기관현황(T21) ──
    print("[Pass 1] Filling institution table (T21)...")
    total_cells += fill_institution_table(editor, tables_config.get('T21_institution', {}),
                                           mapper)

    # ── 10. 부속서류(T24, T27) ──
    print("[Pass 1] Filling appendix tables...")
    total_cells += fill_appendix_tables(editor, tables_config, mapper)

    # ── 11. 서술 섹션 마커 주입 (Pass 2 COM 용) ──
    print("[Pass 1] Injecting content markers for Pass 2...")
    marker_count = inject_content_markers(editor, content_map)
    print(f"[Pass 1] Injected {marker_count} markers")

    # ── 12. 빈 개요/아웃라인 단락 제거 (마커~다음 테이블 사이) ──
    print("[Pass 1] Removing outline placeholders...")
    removed = editor.remove_outline_placeholders(start_table=2)
    print(f"[Pass 1] Removed {removed} outline paragraphs")

    # 저장
    editor.save(output_path)
    print(f"[Pass 1] Done: {total_cells} cells filled")
    print(f"[Pass 1] Output: {output_path}")

    return output_path


def fill_cover_table(editor, config, mapper):
    """표지(T0) 셀 채우기."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cover = mapper.get_cover_data()
    cells = config.get('cells', {})
    count = 0

    # 고정값 매핑
    fill_map = {
        '기관명': cover.get('기관명', '동연에스엔티'),
        '업종업태': cover.get('업종업태', '소프트웨어 개발 및 공급업'),
        '사업비_기관명': cover.get('사업비_기관명', '동연에스엔티'),
    }

    # 공동수행 — 단독수행
    fill_map['공동수행_기관1'] = '해당없음 (단독수행)'
    fill_map['공동수행_기관2'] = '해당없음 (단독수행)'

    for key, value in fill_map.items():
        if key in cells:
            cell_info = cells[key]
            if editor.set_cell_text(table, cell_info['row'], cell_info['col'], value):
                count += 1

    return count


def fill_summary_table(editor, config, mapper):
    """과제요약서(T2) 셀 채우기."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    summary = mapper.get_summary_data()
    cells = config.get('cells', {})
    count = 0

    # 과제명
    project_name = '제조 지능화 구현을 위한 RDBMS 기반 지식 그래프 자동생성 및 온톨로지 관리 플랫폼 개발'
    if '과제명' in cells:
        c = cells['과제명']
        if editor.set_cell_text(table, c['row'], c['col'], project_name):
            count += 1

    # 요약서 필드 목록: (키, 최대길이)
    summary_fields = [
        ('과제요약', 800),
        ('과제목표', 600),
        ('개발내용', 600),
        ('수행방법', 600),
        ('사업화전략', 400),
        ('최종결과물', 400),
        ('기대효과', 400),
    ]

    for field_key, max_len in summary_fields:
        if field_key not in cells or field_key not in summary:
            continue
        c = cells[field_key]
        text = mapper.truncate_text(summary[field_key], max_len)
        # 마크다운 파이프 테이블을 구조화된 텍스트로 변환
        text = format_table_as_text(text)
        if editor.set_cell_text(table, c['row'], c['col'], text):
            count += 1

    return count


def fill_market_table(editor, config):
    """시장규모(T4) 채우기 — MD에서 추출한 시장 데이터."""
    if 'table_index' not in config:
        return 0
    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    # 2024/2026년 글로벌 지식 그래프 시장 데이터
    market_data = {
        '년도_1': '2024',
        '규모_1': 'USD 10.68억 (약 1.45조원)',
        '근거_1': '지식 그래프 기술 시장 급성장 (CAGR 36.6%)',
        '출처_1': 'MarketsandMarkets (2024)',
        '년도_2': '2026',
        '규모_2': 'USD 19.94억 (약 2.71조원)',
        '근거_2': '제조 AI + 지식 그래프 수요 증가, 국내 스마트제조 약 10.5조원',
        '출처_2': 'MarketsandMarkets, SK AX 리포트',
    }

    for key, value in market_data.items():
        if key in cells:
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], value):
                count += 1

    return count


def fill_demand_table(editor, config):
    """수요처(T5) 채우기 — 주요 수요처 정보."""
    if 'table_index' not in config:
        return 0
    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    demand_data = {
        '수요처_1_이름': '동국R&S',
        '수요처_1_관계': '그룹사 (자사 MES 직접 운영)',
        '수요처_1_수요량': 'MES 1식',
        '수요처_1_비고': 'Phase 1 파일럿',
        '수요처_2_이름': 'HD현대중공업',
        '수요처_2_관계': '기존 IT 서비스 고객',
        '수요처_2_수요량': 'MES 1식',
        '수요처_2_비고': 'Phase 2 확장 예정',
        '수요처_3_이름': '한국철강',
        '수요처_3_관계': '기존 MES 고객',
        '수요처_3_수요량': 'MES 1식',
        '수요처_3_비고': 'Phase 2 확장 예정',
    }

    for key, value in demand_data.items():
        if key in cells:
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], value):
                count += 1

    return count


def fill_common_goals_table(editor, config):
    """정량적 공통목표(T6) 자율목표 행 채우기."""
    if 'table_index' not in config:
        return 0
    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    goals_data = {
        '자율목표1_세부': '특허 출원',
        '자율목표1_측정방법': '특허 출원 건수',
        '자율목표1_측정자료': '특허 출원 증명서',
        '자율목표2_세부': '경남 디지털 위크 참여',
        '자율목표2_측정방법': '참여/발표 횟수',
        '자율목표2_측정자료': '참여 확인서',
    }

    for key, value in goals_data.items():
        if key in cells:
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], value):
                count += 1

    return count


def fill_kpi_table(editor, config):
    """성능목표(T8) 빈 행 채우기."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    # 추가 KPI 항목
    extra_kpis = [
        ('목표4', '온톨로지 자동 변환율', '70% 이상',
         'MES 코어 테이블 30~40개 중 수동 보정 없이 정확한 매핑 비율'),
        ('목표5', '지식 그래프 커버리지', '80% 이상',
         'MES 대상 테이블 중 지식 그래프에 반영된 비율(엔티티/관계 기준)'),
    ]

    for prefix, name, target, method in extra_kpis:
        name_key = f'{prefix}_이름'
        target_key = f'{prefix}_지표'
        method_key = f'{prefix}_방법'

        if name_key in cells:
            c = cells[name_key]
            if editor.set_cell_text(table, c['row'], c['col'], name):
                count += 1
        if target_key in cells:
            c = cells[target_key]
            if editor.set_cell_text(table, c['row'], c['col'], target):
                count += 1
        if method_key in cells:
            c = cells[method_key]
            if editor.set_cell_text(table, c['row'], c['col'], method):
                count += 1

    return count


def abbreviate_role(text, max_len=12):
    """긴 담당분야 텍스트를 양식 칸 폭에 맞게 축약한다.

    양식의 담당분야 열은 ~1.35cm로 매우 좁다.
    괄호 내용을 제거하고, 쉼표 구분 항목 중 칸에 맞는 만큼만 유지한다.
    """
    import re
    if len(text) <= max_len:
        return text

    short = re.sub(r'\([^)]*\)', '', text).strip()
    short = re.sub(r'\s+', ' ', short)

    if len(short) <= max_len:
        return short

    parts = [p.strip() for p in short.split(',') if p.strip()]
    if not parts:
        return text[:max_len]

    result = parts[0]
    if len(result) > max_len:
        words = result.split()
        result = words[0]
        for w in words[1:]:
            if len(result) + 1 + len(w) <= max_len:
                result += ' ' + w
            else:
                break
        return result

    for part in parts[1:]:
        candidate = result + '/' + part
        if len(candidate) <= max_len:
            result = candidate
        else:
            break

    return result


def fill_researcher_table(editor, config, mapper):
    """참여연구원-주관(T12) 셀 채우기.

    row 2: 연구원1 (예시 덮어쓰기), row 5: 연구원2, row 6~11: 연구원3~8.
    rows 3-4: 예시 기간분리 행 → 비우기.
    """
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    researchers = mapper.get_researcher_data()
    count = 0

    # rows 3-4 비우기 (예시 기간분리 서브행)
    clear_rows = config.get('clear_rows', [])
    for clear_row in clear_rows:
        for col in [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]:
            editor.set_cell_text(table, clear_row, col, '')

    # MD의 연구원 데이터를 양식 셀에 매핑 (r1~r8, 최대 8명)
    for ri, row_data in enumerate(researchers[:8]):
        row_prefix = f'r{ri + 1}'  # r1, r2, ..., r8

        md_구분 = strip_markdown(row_data[0]) if len(row_data) > 0 else ''
        md_성명 = strip_markdown(row_data[1]) if len(row_data) > 1 else ''
        md_직급 = strip_markdown(row_data[2]) if len(row_data) > 2 else ''
        md_전문 = strip_markdown(row_data[3]) if len(row_data) > 3 else ''
        md_담당 = strip_markdown(row_data[4]) if len(row_data) > 4 else ''
        md_참여율 = strip_markdown(row_data[5]).replace('%', '') if len(row_data) > 5 else ''

        # 담당분야 축약 (c7 폭 1.35cm — 최대 12자)
        md_담당 = abbreviate_role(md_담당, max_len=12)

        # 참여기간: 원본 양식 형식 ("4.1~") 맞춤
        if '5' in md_구분 and '선택' in md_구분:
            참여기간 = '7.1~'
        else:
            참여기간 = '4.1~'

        인력구분 = md_구분.split('(')[0].strip() if '(' in md_구분 else md_구분

        field_map = {
            f'{row_prefix}_번호': str(ri + 1),
            f'{row_prefix}_성명': md_성명,
            f'{row_prefix}_직위': md_직급,
            f'{row_prefix}_담당': md_담당,
            f'{row_prefix}_기간': 참여기간,
            f'{row_prefix}_구분': 인력구분,
            f'{row_prefix}_참여율': md_참여율,
            f'{row_prefix}_평균': md_참여율,
        }

        for key, value in field_map.items():
            if key in cells and value:
                c = cells[key]
                if editor.set_cell_text(table, c['row'], c['col'], value):
                    count += 1

    return count


def fill_production_table(editor, config):
    """생산계획(T14) — MD 기반 매출/판매 데이터 삽입."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    # MD 사업화 전략에서 추출한 데이터 (4단계 전략 → 3년 계획으로 변환)
    # 양식: 개발 종료후 1년/2년/3년 = 2027/2028/2029
    data = {
        # 국내 시장
        '국내_점유율_1': '0.1%',
        '국내_점유율_2': '0.5%',
        '국내_점유율_3': '1.0%',
        '국내_판매량_1': '3~5건',
        '국내_판매량_2': '8~12건',
        '국내_판매량_3': '15~25건',
        '국내_단가_1': '3,000~5,000만원',
        '국내_단가_2': '3,000~5,000만원',
        '국내_단가_3': '3,000~7,000만원',
        '국내_매출_1': '0.9~2.5억원',
        '국내_매출_2': '3~6억원',
        '국내_매출_3': '6~15억원',
        # 해외 시장 (Phase 4: 2029~)
        '해외_점유율_1': '-',
        '해외_점유율_2': '-',
        '해외_점유율_3': '0.01%',
        '해외_판매량_1': '-',
        '해외_판매량_2': '-',
        '해외_판매량_3': '1~3건',
        '해외_단가_1': '-',
        '해외_단가_2': '-',
        '해외_단가_3': '2,000~5,000만원',
        '해외_매출_1': '-',
        '해외_매출_2': '-',
        '해외_매출_3': '0.2~1.5억원',
        # 생산능력
        '생산능력_1': '10건/년',
        '생산능력_2': '20건/년',
        '생산능력_3': '40건/년',
    }

    for key, value in data.items():
        if key in cells:
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], value):
                count += 1

    return count


def fill_budget_table(editor, config):
    """사업비(T18) — placeholder 값 삽입."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    # 사업비는 별도 담당자가 확정하므로 placeholder
    for key in cells:
        if '내역' not in key:  # 내역 셀은 예시 텍스트가 이미 있음
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], '[추후 확정]'):
                count += 1

    return count


def fill_institution_table(editor, config, mapper):
    """기관현황(T21) 셀 채우기."""
    if 'table_index' not in config:
        return 0

    table = editor.get_table(config['table_index'])
    if table is None:
        return 0

    cells = config.get('cells', {})
    count = 0

    # 기관 고정 정보 (CLAUDE.md + MD에서 추출)
    inst_data = {
        '주관_기관명': '동연에스엔티',
        '공동_기관명': '해당없음 (단독수행)',
        '주관_대표자': '[추후 확정]',
        '주관_기업유형': '중소기업 (소프트웨어 전문)',
        '주관_설립일': '[추후 확정]',
        '주관_주생산품': '소프트웨어 개발 및 공급 (MES, CMMS, 그룹웨어, ERP 등)',
        '주관_종업원수': '101명 (기술인력 90%, 특급+고급 56%)',
        '주관_매출액': '약 76억원 (연평균)',
        '주관_주소': '경남 [추후 확정]',
    }

    # MD에서 추가 정보 추출
    md_inst = mapper.get_institution_data()
    for k, v in md_inst.items():
        if '사업자' in k:
            inst_data['주관_사업자번호'] = v
        elif '대표' in k and v != '[추후 확정]':
            inst_data['주관_대표자'] = v
        elif '설립' in k and v != '[추후 확정]':
            inst_data['주관_설립일'] = v
        elif '매출' in k:
            inst_data['주관_매출액'] = v
        elif '자본' in k and '총계' not in k:
            inst_data['주관_자본금_1'] = v
        elif '자본총계' in k:
            inst_data['주관_자본총계_1'] = v

    for key, value in inst_data.items():
        if key in cells:
            c = cells[key]
            if editor.set_cell_text(table, c['row'], c['col'], value):
                count += 1

    return count


def fill_appendix_tables(editor, tables_config, mapper):
    """부속서류 테이블 채우기."""
    count = 0
    project_name = '제조 지능화 구현을 위한 RDBMS 기반 지식 그래프 자동생성 및 온톨로지 관리 플랫폼 개발'

    # T24: 참여의사확인서
    t24 = tables_config.get('T24_agreement', {})
    if 'table_index' in t24:
        table = editor.get_table(t24['table_index'])
        if table is not None:
            cells = t24.get('cells', {})
            if '과제명' in cells:
                c = cells['과제명']
                if editor.set_cell_text(table, c['row'], c['col'], project_name):
                    count += 1
            if '수행기관' in cells:
                c = cells['수행기관']
                if editor.set_cell_text(table, c['row'], c['col'], '동연에스엔티'):
                    count += 1

    # T27: 자격점검표
    t27 = tables_config.get('T27_checklist', {})
    if 'table_index' in t27:
        table = editor.get_table(t27['table_index'])
        if table is not None:
            cells = t27.get('cells', {})
            if '과제명' in cells:
                c = cells['과제명']
                if editor.set_cell_text(table, c['row'], c['col'], project_name):
                    count += 1

    return count


def inject_content_markers(editor, content_map):
    """서술 섹션 마커를 양식 테이블 뒤에 주입한다."""
    narrative_sections = content_map.get('narrative_sections', [])
    count = 0

    for ns in narrative_sections:
        marker = ns['marker']
        after_table = ns.get('insert_after_table')
        if after_table is not None:
            if editor.inject_marker(after_table, marker):
                count += 1
                print(f"  Injected {marker} after table {after_table}")
            else:
                print(f"  WARNING: Failed to inject {marker} after table {after_table}")

    return count


# ── Pass 2: COM 서술 본문 삽입 ─────────────────────────────

def run_pass2(pass1_output, md_path, output_hwpx, output_pdf=None):
    """Pass 2: COM으로 마커 위치에 서술 본문을 삽입한다.

    Args:
        pass1_output: Pass 1 결과 HWPX 경로
        md_path: business_plan_v2.md 경로
        output_hwpx: 최종 HWPX 출력 경로
        output_pdf: PDF 출력 경로 (None이면 생략)

    Returns:
        bool: 성공 여부
    """
    print("[Pass 2] Preparing COM operations...")

    content_map = load_content_map()
    blocks = load_and_parse(md_path)
    mapper = SectionMapper(str(DATA_DIR / 'form_content_map.json'), blocks)

    # 섹션별 COM 오퍼레이션 생성
    section_ops_list = []
    for ns in content_map.get('narrative_sections', []):
        marker = ns['marker']
        section_id = ns['md_section']
        section_blocks = mapper.get_section_blocks(section_id)

        if not section_blocks:
            print(f"  WARNING: No blocks for section {section_id}")
            continue

        ops = compile_section_ops(section_blocks, section_id)
        section_ops_list.append({
            'marker': marker,
            'ops': ops,
        })
        print(f"  Section {section_id}: {len(section_blocks)} blocks → {len(ops)} ops")

    if not section_ops_list:
        print("[Pass 2] No sections to insert. Copying Pass 1 output.")
        shutil.copy2(pass1_output, output_hwpx)
        return True

    # PrintMethod 수정 (PDF 페이지 수 보장)
    print("[Pass 2] Fixing PrintMethod...")
    fix_hwpx_for_pdf(pass1_output)

    # COM 실행
    print(f"[Pass 2] Executing {len(section_ops_list)} section(s) via COM...")
    success = fill_template(
        pass1_output,
        section_ops_list,
        output_hwpx,
        output_pdf=output_pdf,
        timeout=600,
    )

    if success:
        print("[Pass 2] Done!")
    else:
        print("[Pass 2] FAILED — check COM output")

    return success


# ── 메인 파이프라인 ────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='HWPX 양식 자동 채우기 파이프라인'
    )
    parser.add_argument('--template', required=True,
                        help='원본 양식 HWPX 파일 경로')
    parser.add_argument('--md', required=True,
                        help='business_plan_v2.md 경로')
    parser.add_argument('--output', default='output/filled',
                        help='출력 디렉토리 (기본: output/filled)')
    parser.add_argument('--pass1-only', action='store_true',
                        help='Pass 1(XML)만 실행')
    parser.add_argument('--pass2-only', action='store_true',
                        help='Pass 2(COM)만 실행 (Pass 1 결과 필요)')
    parser.add_argument('--no-pdf', action='store_true',
                        help='PDF 생성 건너뛰기')
    args = parser.parse_args()

    # 출력 디렉토리 생성
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    template_path = os.path.abspath(args.template)
    md_path = os.path.abspath(args.md)

    pass1_output = str(output_dir / 'form_pass1.hwpx')
    final_output = str(output_dir / 'business_plan_v2_filled.hwpx')
    pdf_output = str(output_dir / 'business_plan_v2_filled.pdf') if not args.no_pdf else None

    if args.pass2_only:
        # Pass 2만 실행 (Pass 1 결과가 이미 있어야 함)
        if not os.path.exists(pass1_output):
            print(f"ERROR: Pass 1 output not found: {pass1_output}")
            sys.exit(1)
        success = run_pass2(pass1_output, md_path, final_output, pdf_output)
    elif args.pass1_only:
        # Pass 1만 실행
        run_pass1(template_path, md_path, pass1_output)
        print(f"\nPass 1 complete. Run Pass 2 with: --pass2-only")
    else:
        # 전체 파이프라인
        print("=" * 60)
        print("HWPX 양식 자동 채우기 — Two-Pass Pipeline")
        print("=" * 60)

        # Pass 1
        result = run_pass1(template_path, md_path, pass1_output)
        if result is None:
            print("Pass 1 failed. Aborting.")
            sys.exit(1)

        print()

        # Pass 2
        success = run_pass2(pass1_output, md_path, final_output, pdf_output)
        if not success:
            print("\nPass 2 failed. Pass 1 output available at:", pass1_output)
            sys.exit(1)

        print()
        print("=" * 60)
        print("Pipeline complete!")
        print(f"  HWPX: {final_output}")
        if pdf_output:
            print(f"  PDF:  {pdf_output}")
        print("=" * 60)


if __name__ == '__main__':
    main()
