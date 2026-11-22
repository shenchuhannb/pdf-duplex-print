# -*- coding: utf-8 -*-
r"""对照实验：不装 WM_ERASEBKGND 子类化，单开 WS_EX_COMPOSITED 是否可用。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_composited_only.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="composited-only-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32

GWL_EXSTYLE = -20
WS_EX_COMPOSITED = 0x02000000

user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]

from ui_main import DuplexPrintApp, _create_root  # noqa: E402

root = _create_root()

# 关键差异：不装子类化
def _no_fix(_event):
    return


DuplexPrintApp._on_root_map_erase_fix = _no_fix

app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+80+80")
root.update()
time.sleep(0.4)
root.update()

hwnd = ctypes.c_void_p(int(root.wm_frame(), 16))
exstyle = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
print("before exstyle=0x%08X composited=%s" % (
    exstyle, bool(exstyle & WS_EX_COMPOSITED)))

user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, exstyle | WS_EX_COMPOSITED)
user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                    0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0100)
deadline = time.time() + 3
ok = True
try:
    for i in range(5):
        root.update()
        time.sleep(0.05)
except Exception as exc:
    ok = False
    print("UPDATE RAISED:", exc)
print("updates_ok:", ok, "alive:", root.winfo_exists())
if ok:
    root.geometry("1200x800+60+60")
    root.update()
    print("geometry ok:", root.winfo_width(), root.winfo_height())
    now = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
    print("after exstyle=0x%08X composited=%s" % (
        now, bool(now & WS_EX_COMPOSITED)))
    app.shutting_down = True
root.destroy()
print("DONE")
