# -*- coding: utf-8 -*-
r"""干净实验：只用 GCL_HBRBACKGROUND 类画刷，无任何 WndProc 子类化。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_class_brush.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="classbrush-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

GCL_HBRBACKGROUND = -10
COLOR_WINDOW = 5

user32.SetClassLongPtrW.restype = ctypes.c_ssize_t
user32.SetClassLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]

from ui_main import DuplexPrintApp, _create_root  # noqa: E402

# 禁用子类化方案，只测类画刷
import win32_flicker  # noqa: E402
win32_flicker.install_erase_background = lambda *a, **k: None

root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+80+80")
root.update()
time.sleep(0.3)
root.update()
print("built", flush=True)

hwnd = int(root.wm_frame(), 16)
rgb = app.colors["bg"].lstrip("#")
colorref = int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)
brush = gdi32.CreateSolidBrush(colorref)
old = user32.SetClassLongPtrW(ctypes.c_void_p(hwnd), GCL_HBRBACKGROUND, brush)
print("SetClassLongPtrW old=%s brush=%s" % (old, brush), flush=True)

root.update()
print("update ok", flush=True)

WM_SYSCOMMAND = 0x0112
SC_MAXIMIZE = 0xF030
SC_RESTORE = 0xF120
user32.PostMessageW.restype = ctypes.c_int
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                ctypes.c_ssize_t]
hwnd_c = ctypes.c_void_p(hwnd)
user32.PostMessageW(hwnd_c, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
t0 = time.perf_counter()
while time.perf_counter() - t0 < 1.0:
    root.update()
    time.sleep(0.02)
print("maximize ok", flush=True)
user32.PostMessageW(hwnd_c, WM_SYSCOMMAND, SC_RESTORE, 0)
t0 = time.perf_counter()
while time.perf_counter() - t0 < 1.0:
    root.update()
    time.sleep(0.02)
print("restore ok", flush=True)

for value in range(140, 300, 8):
    app._set_log_height_during_drag(value)
    root.update()
print("log drag ok", flush=True)

app.shutting_down = True
root.destroy()
print("DONE", flush=True)
