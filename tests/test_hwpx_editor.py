"""HwpxEditor 단위 테스트."""

import os
import tempfile
import zipfile

import pytest
from lxml import etree

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from src.hwpx_editor import HwpxEditor, NAMESPACES

_PROJECT_ROOT = os.path.join(os.path.dirname(__file__), '..', '..')
REF_HWPX = os.path.join(os.path.dirname(__file__), '..', 'ref', 'test_01.hwpx')
FORM_HWPX = os.path.join(_PROJECT_ROOT, 'form_to_fillout.hwpx')


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


@pytest.fixture
def form_editor():
    if not os.path.exists(FORM_HWPX):
        pytest.skip('form_to_fillout.hwpx not found')
    return HwpxEditor(FORM_HWPX)


def test_inject_marker_no_nested_p(form_editor):
    """inject_marker는 hs:sec 레벨에 hp:p를 삽입해야 한다 (hp:p 중첩 금지)."""
    ok = form_editor.inject_marker(0, '##TEST_MARKER##')
    assert ok is True

    # 마커 텍스트 확인
    found = False
    for t in form_editor.root.findall('.//hp:t', NAMESPACES):
        if t.text == '##TEST_MARKER##':
            found = True
            # 마커의 hp:p 부모가 hs:sec의 직속 자식인지 확인
            run = t.getparent()
            p = run.getparent()
            sec = p.getparent()
            assert p.tag.endswith('}p'), f"Expected hp:p, got {p.tag}"
            assert sec.tag.endswith('}sec'), \
                f"Marker hp:p parent should be hs:sec, got {sec.tag}"
    assert found, "Marker text not found in document"

    # hp:p 중첩이 없어야 한다
    nested = form_editor.root.findall('.//hp:p/hp:p', NAMESPACES)
    assert len(nested) == 0, f"Found {len(nested)} nested hp:p elements"


def test_inject_marker_out_of_range(form_editor):
    assert form_editor.inject_marker(-1, '##BAD##') is False
    assert form_editor.inject_marker(9999, '##BAD##') is False


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
