"""Smoke tests for the current Dongkuk HWPX/HWP pipeline inputs."""

import os
import json
import shutil
from pathlib import Path

import pytest

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.bridge import check_com_available, convert_document, inspect_document
from src.document_pipeline import main as pipeline_main
from src.hwpx_editor import HwpxEditor


PROJECT_ROOT = Path(__file__).resolve().parents[1]
WORKSPACE_ROOT = PROJECT_ROOT.parent
RESULT_DIR = WORKSPACE_ROOT / '동국산업_자율형공장_중간산출물_문서' / 'result'
SOURCE_DIR = WORKSPACE_ROOT / '동국산업_자율형공장_중간산출물_문서'
RESULT_HWPX = sorted(RESULT_DIR.glob('*.hwpx'))
WEEKLY_HWPX = RESULT_DIR / '★17 SF-PM2 주간보고서 양식_V1.2.hwpx'
WEEKLY_MAP = PROJECT_ROOT / 'templates' / 'dongkuk_reports' / 'weekly_report.json'
RESULT_PDFS = sorted(RESULT_DIR.glob('*.pdf'))


@pytest.mark.parametrize('path', RESULT_HWPX)
def test_current_result_hwpx_package_inspection(path):
    info = HwpxEditor.inspect_package(str(path))
    assert info['ok'], info
    assert info['mimetype'] == 'application/hwp+zip'
    assert info['section_count'] >= 1
    assert info['table_count'] > 0


@pytest.mark.parametrize('path', RESULT_HWPX)
def test_current_result_hwpx_table_inventory(path):
    editor = HwpxEditor(str(path))
    tables = editor.inspect_tables()
    assert len(tables) == editor.get_table_count()
    assert tables
    assert all(t['cell_count'] > 0 for t in tables)


def test_set_cell_text_by_index_on_current_hwpx_copy(tmp_path):
    if not RESULT_HWPX:
        pytest.skip('current result HWPX files not found')

    src = RESULT_HWPX[0]
    dst = tmp_path / src.name
    shutil.copy2(src, dst)

    editor = HwpxEditor(str(dst))
    target = None
    for table in editor.inspect_tables(include_empty=True):
        if table['empty_cells']:
            empty = table['empty_cells'][0]
            target = (table['index'], empty['row'], empty['col'])
            break
    if target is None:
        pytest.skip('no empty cell candidate found')

    table_index, row, col = target
    assert editor.set_cell_text_by_index(table_index, row, col, '스모크테스트')
    editor.save(str(dst))

    reloaded = HwpxEditor(str(dst))
    table = reloaded.get_table(table_index)
    cell = reloaded.get_cell(table, row, col)
    assert HwpxEditor._cell_text(cell) == '스모크테스트'

    info = HwpxEditor.inspect_package(str(dst))
    assert info['ok'], info


def test_document_pipeline_inspect_no_com_on_current_hwpx(capsys):
    if not RESULT_HWPX:
        pytest.skip('current result HWPX files not found')

    rc = pipeline_main([
        'inspect',
        '--input', str(RESULT_HWPX[0]),
        '--no-com',
        '--json',
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert '"hwpx"' in captured.out
    assert '"ok": true' in captured.out


def test_document_pipeline_batch_inspect_no_com(capsys):
    rc = pipeline_main([
        'batch-inspect',
        '--input-root', str(SOURCE_DIR),
        '--no-com',
        '--no-root-hwp',
        '--json',
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert f'"count": {len(RESULT_HWPX)}' in captured.out
    assert '"failures": []' in captured.out


def test_document_pipeline_fill_cells_dry_run(tmp_path, capsys):
    if not WEEKLY_HWPX.exists():
        pytest.skip(f'weekly HWPX not found: {WEEKLY_HWPX}')

    data = {
        'current_week_results': '금주 실적 테스트',
        'next_week_plan': '차주 계획 테스트',
        'risk1_content': '위험 내용 테스트',
        'risk1_status': '처리 현황 테스트',
    }
    data_path = tmp_path / 'weekly_data.json'
    data_path.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')

    rc = pipeline_main([
        'fill-cells',
        '--input', str(WEEKLY_HWPX),
        '--output', str(tmp_path / 'weekly_out.hwpx'),
        '--map', str(WEEKLY_MAP),
        '--data', str(data_path),
        '--dry-run',
        '--json',
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert '"dry_run": true' in captured.out
    assert '"current_week_results"' in captured.out


def test_document_pipeline_fill_cells_on_copy(tmp_path):
    if not WEEKLY_HWPX.exists():
        pytest.skip(f'weekly HWPX not found: {WEEKLY_HWPX}')

    data = {
        'current_week_results': '금주 실적 테스트',
        'next_week_plan': '차주 계획 테스트',
    }
    data_path = tmp_path / 'weekly_data.json'
    data_path.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    out_path = tmp_path / 'weekly_out.hwpx'

    rc = pipeline_main([
        'fill-cells',
        '--input', str(WEEKLY_HWPX),
        '--output', str(out_path),
        '--map', str(WEEKLY_MAP),
        '--data', str(data_path),
        '--json',
    ])
    assert rc == 0

    editor = HwpxEditor(str(out_path))
    table = editor.get_table(1)
    assert HwpxEditor._cell_text(editor.get_cell(table, 1, 0)) == '금주 실적 테스트'
    assert HwpxEditor._cell_text(editor.get_cell(table, 1, 1)) == '차주 계획 테스트'


def test_document_pipeline_verify_existing_pdf(capsys):
    if not RESULT_PDFS:
        pytest.skip('current result PDFs not found')

    rc = pipeline_main([
        'verify-pdf',
        '--input', str(RESULT_PDFS[0]),
        '--json',
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert '"ok": true' in captured.out


@pytest.mark.skipif(
    os.environ.get('RUN_HWP_COM_TESTS') != '1',
    reason='set RUN_HWP_COM_TESTS=1 to run Windows Hangul COM smoke tests',
)
def test_windows_com_health_and_conversion(tmp_path):
    source_hwp = SOURCE_DIR / '★19 SF-PM4 회의록 양식_V1.0.hwp'
    if not source_hwp.exists():
        pytest.skip(f'source HWP not found: {source_hwp}')

    health = check_com_available(timeout=60)
    assert health['ok'], health

    inspected = inspect_document(str(source_hwp), timeout=120)
    assert inspected['ok'], inspected
    assert inspected['pages'] > 0

    work = WORKSPACE_ROOT / '.tmp_hwpx_generator_com_tests'
    work.mkdir(parents=True, exist_ok=True)
    out_hwpx = work / 'converted.hwpx'
    out_pdf = work / 'converted.pdf'
    hwpx = convert_document(str(source_hwp), str(out_hwpx), format='HWPX', timeout=300)
    pdf = convert_document(str(source_hwp), str(out_pdf), format='PDF', timeout=300)
    assert hwpx['ok'], hwpx
    assert pdf['ok'], pdf
    assert out_hwpx.exists() and out_hwpx.stat().st_size > 0
    assert out_pdf.exists() and out_pdf.stat().st_size > 0
