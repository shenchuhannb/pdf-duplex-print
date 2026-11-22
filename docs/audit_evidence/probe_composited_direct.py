# -*- coding: utf-8 -*-
r"""隔离实验：绕过 after 回调，在 update() 外直接安装 COMPOSITED。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_composited_direct.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="comp-direct-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
GWL_EXSTYLE = -20
WS_EX_COMPOSITED = 0x02000000
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]

import win32_flicker  # noqa: E402
from ui_main import DuplexPrintApp, _create_root  # noqa: E402

root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+80+80")
print("built", flush=True)

# 关键：先让首帧完成，并且不依赖 after 回调
root.update()
print("first update ok", flush=True)
time.sleep(0.3)
root.update()
print("second update ok", flush=True)

# 现在在事件循环外直接安装
win32_flicker.install_erase_background(root, app.colors["bg"])
hwnd = ctypes.c_void_p(int(root.wm_frame(), 16))
ex = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
print("installed, exstyle=0x%08X composited=%s" % (
    ex, bool(ex & WS_EX_COMPOSITED)), flush=True)

root.update()
print("update after install ok", flush=True)

WM_SYSCOMMAND = 0x0112
SC_MAXIMIZE = 0xF030
SC_RESTORE = 0xF120
user32.PostMessageW.restype = ctypes.c_int
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                ctypes.c_ssize_t]
user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
t0 = time.perf_counter()
while time.perf_counter() - t0 < 0.8:
    root.update()
    time.sleep(0.02)
print("maximize ok", flush=True)
user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
t0 = time.perf_counter()
while time.perf_counter() - t0 < 0.8:
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
