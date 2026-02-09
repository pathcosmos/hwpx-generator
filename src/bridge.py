"""
WSL <-> Windows Python 브릿지

WSL 환경에서 Windows Python을 호출하여 COM 자동화 스크립트를 실행합니다.
"""
import subprocess
import json
import os
import sys
import zipfile

WIN_PYTHON = "/mnt/c/Users/lanco/AppData/Local/Microsoft/WindowsApps/python.exe"


def wsl_to_win_path(wsl_path):
    """WSL 경로를 Windows 경로로 변환

    Args:
        wsl_path: WSL 경로 (예: /mnt/d/project/file.hwpx)

    Returns:
        str: Windows 경로 (예: D:\\project\\file.hwpx)
    """
    wsl_path = os.path.abspath(wsl_path)
    if wsl_path.startswith("/mnt/"):
        parts = wsl_path.split("/")
        drive = parts[2].upper()
        rest = "\\".join(parts[3:])
        return f"{drive}:\\{rest}"
    raise ValueError(f"Cannot convert non-/mnt/ path to Windows path: {wsl_path}")


def win_to_wsl_path(win_path):
    """Windows 경로를 WSL 경로로 변환

    Args:
        win_path: Windows 경로 (예: D:\\project\\file.hwpx)

    Returns:
        str: WSL 경로 (예: /mnt/d/project/file.hwpx)
    """
    win_path = win_path.replace("/", "\\")
    if len(win_path) >= 2 and win_path[1] == ":":
        drive = win_path[0].lower()
        rest = win_path[2:].replace("\\", "/")
        return f"/mnt/{drive}{rest}"
    raise ValueError(f"Cannot convert path to WSL format: {win_path}")


def run_com_script(script_path, *args, timeout=120):
    """Windows Python으로 COM 스크립트 실행

    Args:
        script_path: 실행할 Python 스크립트의 WSL 경로
        *args: 스크립트에 전달할 인자
        timeout: 실행 제한 시간 (초)

    Returns:
        subprocess.CompletedProcess: 실행 결과

    Raises:
        subprocess.TimeoutExpired: 시간 초과
        subprocess.CalledProcessError: 스크립트 실행 실패
    """
    win_script = wsl_to_win_path(script_path)
    cmd = [WIN_PYTHON, win_script] + [str(a) for a in args]
    result = subprocess.run(
        cmd,
        capture_output=True,
        timeout=timeout,
    )
    # Decode with fallback for mixed encodings (cp949/utf-8)
    result.stdout = result.stdout.decode("utf-8", errors="replace")
    result.stderr = result.stderr.decode("utf-8", errors="replace")
    return result


def fix_hwpx_for_pdf(input_hwpx, output_hwpx=None):
    """HWPX 파일의 인쇄 설정을 수정하여 올바른 PDF 출력을 생성하도록 함

    수정 내용:
    - PrintMethod=4 (2페이지/장) → 0 (1페이지/장)

    Args:
        input_hwpx: 원본 HWPX 파일 경로
        output_hwpx: 수정된 HWPX 저장 경로 (None이면 input_hwpx 덮어쓰기)

    Returns:
        str: 수정된 HWPX 파일 경로
    """
    if output_hwpx is None:
        output_hwpx = input_hwpx

    # 임시 파일로 먼저 쓰고 교체 (덮어쓰기 지원)
    tmp_path = output_hwpx + ".tmp"
    fixed = False

    with zipfile.ZipFile(input_hwpx, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == 'settings.xml':
                    text = data.decode('utf-8')
                    if '"PrintMethod" type="short">4<' in text:
                        text = text.replace(
                            '"PrintMethod" type="short">4<',
                            '"PrintMethod" type="short">0<'
                        )
                        data = text.encode('utf-8')
                        item.file_size = len(data)
                        fixed = True

                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data, compress_type=item.compress_type)

    # 임시 파일을 최종 위치로 이동
    if os.path.exists(output_hwpx) and output_hwpx != input_hwpx:
        os.unlink(output_hwpx)
    os.replace(tmp_path, output_hwpx)

    return output_hwpx


def open_and_save_as_pdf(hwpx_path, pdf_path, timeout=300):
    """HWPX/HWP를 열어서 PDF로 저장

    Args:
        hwpx_path: 원본 HWPX/HWP 파일의 WSL 경로
        pdf_path: 출력 PDF 파일의 WSL 경로
        timeout: 실행 제한 시간 (초, 기본 300=5분)

    Returns:
        bool: 성공 여부
    """
    win_hwpx = wsl_to_win_path(hwpx_path)
    win_pdf = wsl_to_win_path(pdf_path)

    script = f'''\
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.hwp_com import HwpController
hwp = HwpController(visible=False)
try:
    start = time.time()
    hwp.open(r"{win_hwpx}")
    pages = hwp.get_page_count()
    print(f"Opened: {{pages}} pages ({{time.time()-start:.1f}}s)", flush=True)
    t = time.time()
    hwp.save_as_pdf(r"{win_pdf}")
    print(f"PDF saved ({{time.time()-t:.1f}}s)", flush=True)
    print("OK", flush=True)
finally:
    hwp.quit()
'''
    return _run_inline_script(script, timeout=timeout)


def open_and_replace(template_path, replacements, output_hwpx, output_pdf=None,
                     timeout=120):
    """템플릿을 열고 텍스트 교체 후 저장

    Args:
        template_path: 템플릿 HWPX 파일의 WSL 경로
        replacements: dict {찾을_텍스트: 바꿀_텍스트, ...}
        output_hwpx: 출력 HWPX 파일의 WSL 경로
        output_pdf: 출력 PDF 파일의 WSL 경로 (None이면 PDF 미생성)
        timeout: 실행 제한 시간 (초)

    Returns:
        bool: 성공 여부
    """
    win_template = wsl_to_win_path(template_path)
    win_output = wsl_to_win_path(output_hwpx)
    replacements_json = json.dumps(replacements, ensure_ascii=False)

    pdf_line = ""
    if output_pdf:
        win_pdf = wsl_to_win_path(output_pdf)
        pdf_line = f'    hwp.save_as_pdf(r"{win_pdf}")'

    script = f'''\
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.hwp_com import HwpController
replacements = json.loads({repr(replacements_json)})
hwp = HwpController(visible=False)
try:
    hwp.open(r"{win_template}")
    hwp.find_and_replace_all(replacements)
    hwp.save_as(r"{win_output}", "HWPX")
{pdf_line}
    print("OK", flush=True)
finally:
    hwp.quit()
'''
    return _run_inline_script(script, timeout=timeout)


def create_document(operations, output_path, output_pdf=None, timeout=120):
    """새 문서 생성 (JSON 기반 명령어)

    Args:
        operations: 명령어 리스트 [{"op": "insert_text", "text": "..."}, ...]
            지원 명령어:
            - {"op": "insert_text", "text": "..."}
            - {"op": "line_break"}
            - {"op": "set_char_shape", "font": "...", "size": N, "bold": bool, "color": N}
            - {"op": "set_para_shape", "align": "center", "line_spacing": N}
            - {"op": "insert_table", "rows": N, "cols": N}
            - {"op": "fill_table", "data": [[...], ...]}
        output_path: 출력 HWPX 파일의 WSL 경로
        output_pdf: 출력 PDF 파일의 WSL 경로 (None이면 PDF 미생성)
        timeout: 실행 제한 시간 (초)

    Returns:
        bool: 성공 여부
    """
    win_output = wsl_to_win_path(output_path)
    ops_json = json.dumps(operations, ensure_ascii=False)

    pdf_line = ""
    if output_pdf:
        win_pdf = wsl_to_win_path(output_pdf)
        pdf_line = f'    hwp.save_as_pdf(r"{win_pdf}")'

    script = f'''\
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.hwp_com import HwpController
operations = json.loads({repr(ops_json)})
hwp = HwpController(visible=False)
try:
    for op in operations:
        cmd = op["op"]
        if cmd == "insert_text":
            hwp.insert_text(op["text"])
        elif cmd == "line_break":
            hwp.insert_line_break()
        elif cmd == "set_char_shape":
            kwargs = {{k: v for k, v in op.items() if k != "op"}}
            hwp.set_char_shape(**kwargs)
        elif cmd == "set_para_shape":
            kwargs = {{k: v for k, v in op.items() if k != "op"}}
            hwp.set_para_shape(**kwargs)
        elif cmd == "insert_table":
            hwp.insert_table(op["rows"], op["cols"])
        elif cmd == "fill_table":
            hwp.fill_table(op["data"])
    hwp.save_as(r"{win_output}", "HWPX")
{pdf_line}
    print("OK", flush=True)
finally:
    hwp.quit()
'''
    return _run_inline_script(script, timeout=timeout)


def _run_inline_script(script_code, timeout=120):
    """인라인 Python 스크립트를 Windows Python으로 실행

    Args:
        script_code: 실행할 Python 코드
        timeout: 실행 제한 시간 (초)

    Returns:
        bool: 성공 여부 (stdout에 "OK" 포함 시 True)
    """
    # Write script to a temp file in the project directory (accessible from Windows)
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    tmp_path = os.path.join(project_dir, "_bridge_tmp.py")
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(script_code)

        win_tmp = wsl_to_win_path(tmp_path)
        proc = subprocess.Popen(
            [WIN_PYTHON, win_tmp],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        try:
            stdout_bytes, stderr_bytes = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            stdout_bytes, stderr_bytes = proc.communicate()
            print("[bridge] Process timed out", file=sys.stderr)
            return False

        stdout = stdout_bytes.decode("utf-8", errors="replace")
        stderr = stderr_bytes.decode("utf-8", errors="replace")

        if proc.returncode != 0:
            print(f"[bridge] STDERR: {stderr}", file=sys.stderr)

        return "OK" in stdout
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
