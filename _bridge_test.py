import sys
print(f"PYTHON_VERSION={sys.version}", flush=True)
try:
    import win32com.client
    print("PYWIN32=OK", flush=True)
except ImportError:
    print("PYWIN32=MISSING", flush=True)
print("OK", flush=True)
