"""HwpxEditor 단위 테스트."""

import os
import tempfile
import zipfile

import pytest
from lxml import etree

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from src.hwpx_editor import HwpxEditor, NAMESPACES

REF_HWPX = os.path.join(os.path.dirname(__file__), '..', 'ref', 'test_01.hwpx')


@pytest.fixture
def editor():
    return HwpxEditor(REF_HWPX)


@pytest.fixture
def cover_table(editor):
    return editor.get_table(0)


def test_get_table(editor):
    tbl = editor.get_table(0)
    assert tbl is not None
    assert tbl.get('rowCnt') == '35'
    assert tbl.get('colCnt') == '11'


def test_get_table_out_of_range(editor):
    assert editor.get_table(99999) is None


def test_get_cell(editor, cover_table):
    cell = editor.get_cell(cover_table, 1, 3)
    assert cell is not None
    # This cell contains the project name text
    t_elems = cell.findall('.//hp:t', NAMESPACES)
    texts = [t.text for t in t_elems if t.text]
    combined = ''.join(texts)
    assert '클라우드' in combined


def test_get_cell_not_found(editor, cover_table):
    assert editor.get_cell(cover_table, 999, 999) is None


def test_set_cell_text(editor, cover_table):
    # Cell (6, 3) is empty in the reference (기업명 value)
    cell = editor.get_cell(cover_table, 6, 3)
    run = cell.find('.//hp:run', NAMESPACES)
    assert run.find('hp:t', NAMESPACES) is None  # empty before

    ok = editor.set_cell_text(cover_table, 6, 3, '테스트기업')
    assert ok is True

    t = run.find('hp:t', NAMESPACES)
    assert t is not None
    assert t.text == '테스트기업'


def test_set_cell_text_missing_cell(editor, cover_table):
    assert editor.set_cell_text(cover_table, 999, 999, 'fail') is False


def test_fill_cells(editor, cover_table):
    data = {
        (6, 3): '기업A',
        (7, 3): '111-22-33333',
        (8, 3): '김대표',
    }
    count = editor.fill_cells(cover_table, data)
    assert count == 3

    for (row, col), text in data.items():
        cell = editor.get_cell(cover_table, row, col)
        t = cell.find('.//hp:t', NAMESPACES)
        assert t is not None
        assert t.text == text


def test_save_and_reload(editor, cover_table):
    editor.set_cell_text(cover_table, 6, 3, '저장테스트')

    with tempfile.NamedTemporaryFile(suffix='.hwpx', delete=False) as f:
        tmp_path = f.name

    try:
        editor.save(tmp_path)
        assert os.path.exists(tmp_path)

        editor2 = HwpxEditor(tmp_path)
        tbl2 = editor2.get_table(0)
        cell2 = editor2.get_cell(tbl2, 6, 3)
        t2 = cell2.find('.//hp:t', NAMESPACES)
        assert t2 is not None
        assert t2.text == '저장테스트'
    finally:
        os.unlink(tmp_path)


def test_namespace_preservation(editor, cover_table):
    editor.set_cell_text(cover_table, 6, 3, 'ns test')

    with tempfile.NamedTemporaryFile(suffix='.hwpx', delete=False) as f:
        tmp_path = f.name

    try:
        editor.save(tmp_path)
        with zipfile.ZipFile(tmp_path, 'r') as z:
            xml_bytes = z.read('Contents/section0.xml')
        xml_str = xml_bytes.decode('utf-8')

        for prefix, uri in NAMESPACES.items():
            assert uri in xml_str, f'Namespace {prefix}={uri} missing'
    finally:
        os.unlink(tmp_path)


def test_charPrIDRef_preserved(editor, cover_table):
    cell = editor.get_cell(cover_table, 6, 3)
    run = cell.find('.//hp:run', NAMESPACES)
    original_ref = run.get('charPrIDRef')
    assert original_ref is not None

    editor.set_cell_text(cover_table, 6, 3, 'style test')

    assert run.get('charPrIDRef') == original_ref


def test_mimetype_stored(editor, cover_table):
    editor.set_cell_text(cover_table, 6, 3, 'zip test')

    with tempfile.NamedTemporaryFile(suffix='.hwpx', delete=False) as f:
        tmp_path = f.name

    try:
        editor.save(tmp_path)
        with zipfile.ZipFile(tmp_path, 'r') as z:
            info = z.getinfo('mimetype')
            assert info.compress_type == zipfile.ZIP_STORED
    finally:
        os.unlink(tmp_path)
