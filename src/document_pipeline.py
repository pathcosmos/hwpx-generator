"""HWP/HWPX document pipeline CLI.

Current focus:
  - Check Windows Hangul COM availability from WSL.
  - Inspect HWP/HWPX files.
  - Convert HWP/HWPX to PDF/HWPX/HWP through Hangul COM.
  - Apply a narrow HWPX table-cell text edit, then optionally export PDF.
"""

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))
DEFAULT_DOC_ROOT = PROJECT_DIR.parent / '동국산업_자율형공장_중간산출물_문서'

from src.bridge import (  # noqa: E402
    check_com_available,
    cleanup_hwp_processes,
    convert_document,
    fix_hwpx_for_pdf,
    inspect_document,
    list_hwp_processes,
)
from src.hwpx_editor import HwpxEditor  # noqa: E402


def _print_json(data):
    print(json.dumps(data, ensure_ascii=False, indent=2))


def _as_abs(path):
    return str(Path(path).expanduser().resolve())


def _exit_code_for(result):
    return 0 if result.get('ok') else 1


def _doc_targets(input_root, include_root_hwp=True, include_result_hwpx=True):
    root = Path(input_root).expanduser().resolve()
    targets = []
    if include_root_hwp:
        targets.extend(sorted(root.glob('*.hwp')))
    if include_result_hwpx:
        targets.extend(sorted((root / 'result').glob('*.hwpx')))
    return targets


def _safe_output_name(input_path, format):
    stem = Path(input_path).stem
    suffix = '.' + str(format).lower()
    return stem + suffix


def _load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def _write_json(path, data):
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def _print_health(result):
    windows_python = result.get('windows_python') or ''
    windows_python = windows_python.splitlines()[0] if windows_python else ''
    print(f"ok:             {result.get('ok')}")
    print(f"windows_python: {windows_python}")
    print(f"pywin32:        {result.get('pywin32')}")
    print(f"hwp_com:        {result.get('hwp_com')}")
    if 'hwp_process_count_before' in result:
        print(f"hwp_before:     {result.get('hwp_process_count_before')}")
        print(f"hwp_after:      {result.get('hwp_process_count_after')}")
    if result.get('error'):
        print(f"error:          {result['error']}")


def cmd_health(args):
    result = check_com_available(timeout=args.timeout)
    if args.json:
        _print_json(result)
    else:
        _print_health(result)
    return _exit_code_for(result)


def _inspect_hwpx(path, include_tables=False):
    package = HwpxEditor.inspect_package(path)
    if include_tables and package.get('ok'):
        editor = HwpxEditor(path)
        package['tables'] = editor.inspect_tables(include_empty=True)
    return package


def _print_inspect(result):
    print(f"input:  {result['input']}")
    print(f"exists: {result['exists']}")
    if result.get('hwpx') is not None:
        hwpx = result['hwpx']
        print(f"hwpx_ok:       {hwpx.get('ok')}")
        print(f"hwpx_sections: {hwpx.get('section_count')}")
        print(f"hwpx_tables:   {hwpx.get('table_count')}")
        if hwpx.get('errors'):
            print(f"hwpx_errors:   {hwpx['errors']}")
        if hwpx.get('warnings'):
            print(f"hwpx_warnings: {hwpx['warnings']}")
    if result.get('com') is not None:
        com = result['com']
        print(f"com_ok:        {com.get('ok')}")
        if com.get('ok'):
            print(f"com_format:    {com.get('format')}")
            print(f"com_pages:     {com.get('pages')}")
        if com.get('error'):
            print(f"com_error:     {com['error']}")


def cmd_inspect(args):
    path = _as_abs(args.input)
    suffix = Path(path).suffix.lower()
    result = {
        'input': path,
        'exists': os.path.exists(path),
        'hwpx': None,
        'com': None,
        'ok': False,
    }
    if not result['exists']:
        result['error'] = f'file not found: {path}'
        if args.json:
            _print_json(result)
        else:
            _print_inspect(result)
            print(f"error:  {result['error']}")
        return 1

    if suffix == '.hwpx':
        result['hwpx'] = _inspect_hwpx(path, include_tables=args.tables)

    if not args.no_com:
        result['com'] = inspect_document(
            path, timeout=args.timeout, include_controls=args.controls
        )

    checks = []
    if result['hwpx'] is not None:
        checks.append(result['hwpx'].get('ok', False))
    if result['com'] is not None:
        checks.append(result['com'].get('ok', False))
    result['ok'] = all(checks) if checks else result['exists']

    if args.json:
        _print_json(result)
    else:
        _print_inspect(result)
    return _exit_code_for(result)


def cmd_convert(args):
    result = convert_document(
        _as_abs(args.input),
        _as_abs(args.output),
        format=args.format.upper(),
        timeout=args.timeout,
    )
    if args.json:
        _print_json(result)
    else:
        print(f"ok:     {result.get('ok')}")
        print(f"input:  {result.get('input_wsl')}")
        print(f"output: {result.get('output_wsl')}")
        if result.get('output'):
            print(f"bytes:  {result['output'].get('bytes'):,}")
        if result.get('error'):
            print(f"error:  {result['error']}")
    return _exit_code_for(result)


def cmd_fill_cell(args):
    input_path = Path(_as_abs(args.input))
    output_path = Path(_as_abs(args.output))
    result = {
        'ok': False,
        'input': str(input_path),
        'output': str(output_path),
        'cell': {
            'table': args.table,
            'row': args.row,
            'col': args.col,
            'text': args.text,
        },
    }

    if input_path.suffix.lower() != '.hwpx':
        result['error'] = 'fill-cell only supports HWPX input'
    elif not input_path.exists():
        result['error'] = f'file not found: {input_path}'
    else:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if input_path.resolve() != output_path.resolve():
            shutil.copy2(input_path, output_path)

        before = HwpxEditor.inspect_package(str(output_path))
        result['before'] = before
        if not before.get('ok'):
            result['error'] = f'HWPX package validation failed: {before.get("errors")}'
        else:
            editor = HwpxEditor(str(output_path))
            ok = editor.set_cell_text_by_index(args.table, args.row, args.col, args.text)
            if not ok:
                result['error'] = (
                    f'cell not found: table={args.table} row={args.row} col={args.col}'
                )
            else:
                editor.save(str(output_path))
                result['after'] = HwpxEditor.inspect_package(str(output_path))
                result['ok'] = result['after'].get('ok', False)

                if args.pdf:
                    fix_hwpx_for_pdf(str(output_path))
                    result['pdf'] = convert_document(
                        str(output_path),
                        _as_abs(args.pdf),
                        format='PDF',
                        timeout=args.timeout,
                    )
                    result['ok'] = result['ok'] and result['pdf'].get('ok', False)

    if args.json:
        _print_json(result)
    else:
        print(f"ok:     {result.get('ok')}")
        print(f"output: {result.get('output')}")
        if args.pdf:
            print(f"pdf:    {_as_abs(args.pdf)}")
        if result.get('error'):
            print(f"error:  {result['error']}")
        if result.get('pdf', {}).get('error'):
            print(f"pdf_error: {result['pdf']['error']}")
    return _exit_code_for(result)


def cmd_cleanup_com(args):
    if args.kill:
        result = cleanup_hwp_processes(kill=True, timeout=args.timeout)
    else:
        result = list_hwp_processes(timeout=args.timeout)

    if args.json:
        _print_json(result)
    else:
        if args.kill:
            print(f"kill_requested: {result.get('kill_requested')}")
            print(f"before:         {result.get('count_before')}")
            print(f"after:          {result.get('count_after')}")
        else:
            print(f"count:          {result.get('count')}")
        if result.get('error'):
            print(f"error:          {result['error']}")
    return _exit_code_for(result)


def _inspect_one(path, timeout, include_com=True, include_controls=False, include_tables=False):
    suffix = Path(path).suffix.lower()
    entry = {
        'path': str(Path(path).resolve()),
        'name': Path(path).name,
        'suffix': suffix,
        'exists': Path(path).exists(),
        'hwpx': None,
        'com': None,
        'ok': False,
    }
    if suffix == '.hwpx':
        entry['hwpx'] = _inspect_hwpx(str(path), include_tables=include_tables)
    if include_com:
        entry['com'] = inspect_document(
            str(path), timeout=timeout, include_controls=include_controls
        )
    checks = []
    if entry['hwpx'] is not None:
        checks.append(entry['hwpx'].get('ok', False))
    if entry['com'] is not None:
        checks.append(entry['com'].get('ok', False))
    entry['ok'] = all(checks) if checks else entry['exists']
    return entry


def cmd_batch_inspect(args):
    targets = _doc_targets(
        args.input_root,
        include_root_hwp=not args.no_root_hwp,
        include_result_hwpx=not args.no_result_hwpx,
    )
    result = {
        'ok': False,
        'input_root': _as_abs(args.input_root),
        'count': len(targets),
        'items': [],
    }
    for path in targets:
        result['items'].append(_inspect_one(
            path,
            timeout=args.timeout,
            include_com=not args.no_com,
            include_controls=args.controls,
            include_tables=args.tables,
        ))
    result['ok'] = bool(targets) and all(item['ok'] for item in result['items'])
    result['failures'] = [item for item in result['items'] if not item['ok']]

    if args.output_report:
        _write_json(args.output_report, result)

    if args.json:
        _print_json(result)
    else:
        print(f"input_root: {_as_abs(args.input_root)}")
        print(f"count:      {result['count']}")
        print(f"ok:         {result['ok']}")
        for item in result['items']:
            pages = item.get('com', {}).get('pages') if item.get('com') else None
            print(f"{'OK' if item['ok'] else 'FAIL'}  {item['name']}  pages={pages}")
    return _exit_code_for(result)


def cmd_batch_convert(args):
    fmt = args.format.upper()
    include_result_hwpx = not args.no_result_hwpx and fmt == 'PDF'
    targets = _doc_targets(
        args.input_root,
        include_root_hwp=not args.no_root_hwp,
        include_result_hwpx=include_result_hwpx,
    )
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    result = {
        'ok': False,
        'input_root': _as_abs(args.input_root),
        'output_dir': str(output_dir),
        'format': fmt,
        'count': len(targets),
        'items': [],
    }

    for path in targets:
        out_path = output_dir / _safe_output_name(path, fmt)
        item = {
            'input': str(Path(path).resolve()),
            'output': str(out_path),
            'format': fmt,
        }
        if out_path.exists() and not args.overwrite:
            item['ok'] = False
            item['error'] = 'output exists; pass --overwrite'
        else:
            converted = convert_document(
                str(path), str(out_path), format=fmt, timeout=args.timeout
            )
            item.update(converted)
            item['ok'] = converted.get('ok', False)
            if item['ok'] and args.verify_pdf and fmt == 'PDF':
                item['pdf_verify'] = verify_pdf_file(
                    str(out_path), require_text=not args.no_require_text
                )
                item['ok'] = item['pdf_verify'].get('ok', False)
        result['items'].append(item)

    result['ok'] = bool(targets) and all(item.get('ok') for item in result['items'])
    result['failures'] = [item for item in result['items'] if not item.get('ok')]

    if args.output_report:
        _write_json(args.output_report, result)

    if args.json:
        _print_json(result)
    else:
        print(f"output_dir: {output_dir}")
        print(f"count:      {result['count']}")
        print(f"ok:         {result['ok']}")
        for item in result['items']:
            print(f"{'OK' if item.get('ok') else 'FAIL'}  {Path(item['input']).name} -> {Path(item['output']).name}")
            if item.get('error'):
                print(f"  error: {item['error']}")
    return _exit_code_for(result)


def _plan_field_fills(field_map, data):
    fields = field_map.get('fields', [])
    planned = []
    errors = []
    for field in fields:
        name = field.get('name')
        if not name:
            errors.append(f'field without name: {field}')
            continue
        has_value = name in data and data[name] is not None and data[name] != ''
        if not has_value:
            if field.get('required', False):
                errors.append(f'required field missing: {name}')
            continue
        try:
            planned.append({
                'name': name,
                'table': int(field['table']),
                'row': int(field['row']),
                'col': int(field['col']),
                'text': str(data[name]),
            })
        except KeyError as exc:
            errors.append(f'field {name} missing coordinate key: {exc}')
        except (TypeError, ValueError) as exc:
            errors.append(f'field {name} invalid coordinate: {exc}')
    return planned, errors


def cmd_fill_cells(args):
    input_path = Path(_as_abs(args.input))
    output_path = Path(_as_abs(args.output))
    field_map = _load_json(args.map)
    data = _load_json(args.data)
    planned, errors = _plan_field_fills(field_map, data)
    result = {
        'ok': False,
        'input': str(input_path),
        'output': str(output_path),
        'map': _as_abs(args.map),
        'data': _as_abs(args.data),
        'planned': planned,
        'errors': errors,
        'dry_run': args.dry_run,
    }

    if input_path.suffix.lower() != '.hwpx':
        result['errors'].append('fill-cells only supports HWPX input')
    if not input_path.exists():
        result['errors'].append(f'file not found: {input_path}')
    if not planned:
        result['errors'].append('no fields to apply')

    if not result['errors']:
        package = HwpxEditor.inspect_package(str(input_path))
        result['before'] = package
        if not package.get('ok'):
            result['errors'].append(f'HWPX package validation failed: {package.get("errors")}')

    if result['errors']:
        if args.json:
            _print_json(result)
        else:
            print(f"ok:     {result['ok']}")
            for error in result['errors']:
                print(f"error:  {error}")
        return 1

    if args.dry_run:
        result['ok'] = True
    else:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if input_path.resolve() != output_path.resolve():
            shutil.copy2(input_path, output_path)
        editor = HwpxEditor(str(output_path))
        apply_errors = []
        for fill in planned:
            ok = editor.set_cell_text_by_index(
                fill['table'], fill['row'], fill['col'], fill['text']
            )
            if not ok:
                apply_errors.append(
                    f"{fill['name']}: cell not found table={fill['table']} row={fill['row']} col={fill['col']}"
                )
        if apply_errors:
            result['errors'].extend(apply_errors)
        else:
            editor.save(str(output_path))
            result['after'] = HwpxEditor.inspect_package(str(output_path))
            result['ok'] = result['after'].get('ok', False)
            if args.pdf:
                fix_hwpx_for_pdf(str(output_path))
                result['pdf'] = convert_document(
                    str(output_path), _as_abs(args.pdf), format='PDF', timeout=args.timeout
                )
                result['ok'] = result['ok'] and result['pdf'].get('ok', False)
                if result['ok'] and args.verify_pdf:
                    result['pdf_verify'] = verify_pdf_file(
                        _as_abs(args.pdf), require_text=not args.no_require_text
                    )
                    result['ok'] = result['pdf_verify'].get('ok', False)

    if args.json:
        _print_json(result)
    else:
        print(f"ok:      {result['ok']}")
        print(f"planned: {len(planned)}")
        print(f"output:  {output_path}")
        if args.pdf:
            print(f"pdf:     {_as_abs(args.pdf)}")
        for error in result.get('errors', []):
            print(f"error:   {error}")
    return _exit_code_for(result)


def verify_pdf_file(path, require_text=True, compare=None, output_dir=None, ssim=False):
    pdf_path = Path(path).expanduser().resolve()
    result = {
        'ok': False,
        'path': str(pdf_path),
        'exists': pdf_path.exists(),
        'bytes': 0,
        'pages': None,
        'text_chars': None,
        'text_extractable': None,
        'warnings': [],
        'errors': [],
    }
    if not result['exists']:
        result['errors'].append(f'file not found: {pdf_path}')
        return result
    result['bytes'] = pdf_path.stat().st_size
    if result['bytes'] <= 0:
        result['errors'].append('PDF file is empty')
        return result

    try:
        import fitz  # PyMuPDF
        doc = fitz.open(str(pdf_path))
        try:
            result['pages'] = doc.page_count
            text_chars = 0
            for page in doc:
                text_chars += len(page.get_text() or '')
            result['text_chars'] = text_chars
            result['text_extractable'] = text_chars > 0
        finally:
            doc.close()
    except ImportError:
        result['warnings'].append('PyMuPDF not installed; page/text checks skipped')
    except Exception as exc:
        result['errors'].append(f'PDF parse failed: {type(exc).__name__}: {exc}')

    if result['pages'] is not None and result['pages'] <= 0:
        result['errors'].append('PDF has no pages')
    if require_text and result['text_extractable'] is False:
        result['errors'].append('PDF text extraction returned no text')

    if compare:
        if ssim:
            try:
                from src.pdf_compare import PdfComparator
                compare_dir = output_dir or str(pdf_path.parent / 'pdf_compare')
                comparator = PdfComparator(str(compare), str(pdf_path), dpi=150)
                result['comparison'] = comparator.compare(output_dir=compare_dir)
                if not result['comparison'].get('pass'):
                    result['errors'].append('PDF comparison failed threshold')
            except Exception as exc:
                result['errors'].append(f'PDF comparison failed: {type(exc).__name__}: {exc}')
        else:
            ref = verify_pdf_file(compare, require_text=require_text)
            result['reference'] = ref
            if not ref.get('ok'):
                result['errors'].append('reference PDF verification failed')
            elif result.get('pages') is not None and ref.get('pages') is not None:
                if result['pages'] != ref['pages']:
                    result['errors'].append(
                        f'page count mismatch: generated={result["pages"]} reference={ref["pages"]}'
                    )

    result['ok'] = not result['errors']
    return result


def cmd_verify_pdf(args):
    result = verify_pdf_file(
        _as_abs(args.input),
        require_text=not args.no_require_text,
        compare=_as_abs(args.compare) if args.compare else None,
        output_dir=_as_abs(args.output_dir) if args.output_dir else None,
        ssim=args.ssim,
    )
    if args.json:
        _print_json(result)
    else:
        print(f"ok:       {result['ok']}")
        print(f"bytes:    {result['bytes']:,}")
        print(f"pages:    {result['pages']}")
        print(f"text:     {result['text_chars']}")
        for error in result.get('errors', []):
            print(f"error:    {error}")
        for warning in result.get('warnings', []):
            print(f"warning:  {warning}")
    return _exit_code_for(result)


def build_parser():
    parser = argparse.ArgumentParser(
        description='HWP/HWPX document pipeline',
    )
    sub = parser.add_subparsers(dest='command', required=True)

    p_health = sub.add_parser('health', help='check Windows Hangul COM availability')
    p_health.add_argument('--timeout', type=int, default=60)
    p_health.add_argument('--json', action='store_true')
    p_health.set_defaults(func=cmd_health)

    p_cleanup = sub.add_parser('cleanup-com', help='report or terminate Hwp.exe processes')
    p_cleanup.add_argument('--kill', action='store_true', help='actually terminate Hwp.exe')
    p_cleanup.add_argument('--timeout', type=int, default=60)
    p_cleanup.add_argument('--json', action='store_true')
    p_cleanup.set_defaults(func=cmd_cleanup_com)

    p_inspect = sub.add_parser('inspect', help='inspect an HWP/HWPX document')
    p_inspect.add_argument('--input', required=True)
    p_inspect.add_argument('--timeout', type=int, default=120)
    p_inspect.add_argument('--no-com', action='store_true')
    p_inspect.add_argument('--controls', action='store_true')
    p_inspect.add_argument('--tables', action='store_true')
    p_inspect.add_argument('--json', action='store_true')
    p_inspect.set_defaults(func=cmd_inspect)

    p_convert = sub.add_parser('convert', help='convert through Hangul COM')
    p_convert.add_argument('--input', required=True)
    p_convert.add_argument('--output', required=True)
    p_convert.add_argument(
        '--format',
        required=True,
        choices=['pdf', 'hwpx', 'hwp', 'html', 'text'],
    )
    p_convert.add_argument('--timeout', type=int, default=300)
    p_convert.add_argument('--json', action='store_true')
    p_convert.set_defaults(func=cmd_convert)

    p_fill = sub.add_parser('fill-cell', help='set one HWPX table cell')
    p_fill.add_argument('--input', required=True)
    p_fill.add_argument('--output', required=True)
    p_fill.add_argument('--table', required=True, type=int)
    p_fill.add_argument('--row', required=True, type=int)
    p_fill.add_argument('--col', required=True, type=int)
    p_fill.add_argument('--text', required=True)
    p_fill.add_argument('--pdf')
    p_fill.add_argument('--timeout', type=int, default=300)
    p_fill.add_argument('--json', action='store_true')
    p_fill.set_defaults(func=cmd_fill_cell)

    p_batch_inspect = sub.add_parser('batch-inspect', help='inspect current document set')
    p_batch_inspect.add_argument('--input-root', default=str(DEFAULT_DOC_ROOT))
    p_batch_inspect.add_argument('--timeout', type=int, default=120)
    p_batch_inspect.add_argument('--no-com', action='store_true')
    p_batch_inspect.add_argument('--no-root-hwp', action='store_true')
    p_batch_inspect.add_argument('--no-result-hwpx', action='store_true')
    p_batch_inspect.add_argument('--controls', action='store_true')
    p_batch_inspect.add_argument('--tables', action='store_true')
    p_batch_inspect.add_argument('--output-report')
    p_batch_inspect.add_argument('--json', action='store_true')
    p_batch_inspect.set_defaults(func=cmd_batch_inspect)

    p_batch_convert = sub.add_parser('batch-convert', help='convert current document set')
    p_batch_convert.add_argument('--input-root', default=str(DEFAULT_DOC_ROOT))
    p_batch_convert.add_argument(
        '--output-dir',
        default=str(DEFAULT_DOC_ROOT / 'result' / 'batch_check'),
    )
    p_batch_convert.add_argument('--format', choices=['pdf', 'hwpx'], default='pdf')
    p_batch_convert.add_argument('--timeout', type=int, default=300)
    p_batch_convert.add_argument('--no-root-hwp', action='store_true')
    p_batch_convert.add_argument('--no-result-hwpx', action='store_true')
    p_batch_convert.add_argument('--overwrite', action='store_true')
    p_batch_convert.add_argument('--verify-pdf', action='store_true')
    p_batch_convert.add_argument('--no-require-text', action='store_true')
    p_batch_convert.add_argument('--output-report')
    p_batch_convert.add_argument('--json', action='store_true')
    p_batch_convert.set_defaults(func=cmd_batch_convert)

    p_fill_cells = sub.add_parser('fill-cells', help='fill multiple HWPX cells from JSON')
    p_fill_cells.add_argument('--input', required=True)
    p_fill_cells.add_argument('--output', required=True)
    p_fill_cells.add_argument('--map', required=True)
    p_fill_cells.add_argument('--data', required=True)
    p_fill_cells.add_argument('--pdf')
    p_fill_cells.add_argument('--verify-pdf', action='store_true')
    p_fill_cells.add_argument('--no-require-text', action='store_true')
    p_fill_cells.add_argument('--dry-run', action='store_true')
    p_fill_cells.add_argument('--timeout', type=int, default=300)
    p_fill_cells.add_argument('--json', action='store_true')
    p_fill_cells.set_defaults(func=cmd_fill_cells)

    p_verify_pdf = sub.add_parser('verify-pdf', help='verify generated PDF')
    p_verify_pdf.add_argument('--input', required=True)
    p_verify_pdf.add_argument('--compare')
    p_verify_pdf.add_argument('--output-dir')
    p_verify_pdf.add_argument('--ssim', action='store_true')
    p_verify_pdf.add_argument('--no-require-text', action='store_true')
    p_verify_pdf.add_argument('--json', action='store_true')
    p_verify_pdf.set_defaults(func=cmd_verify_pdf)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == '__main__':
    raise SystemExit(main())
