# -*- coding: utf-8 -*-
r"""检查窗口样式位与 DWM 合成状态，并实验 WS_EX_COMPOSITED 是否可安全启用。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_ws_composited.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="composited-probe-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32

GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_EX_COMPOSITED = 0x02000000
WS_EX_LAYERED = 0x00080000
WS_EX_NOREDIRECTIONBITMAP = 0x02000000

user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]

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
root.update()
time.sleep(0.4)
root.update()

hwnd = ctypes.c_void_p(int(root.wm_frame(), 16))
style = user32.GetWindowLongPtrW(hwnd, GWL_STYLE)
exstyle = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
print("hwnd: 0x%X" % hwnd.value)
print("style: 0x%08X" % style)
print("exstyle: 0x%08X" % exstyle)
print("has WS_EX_COMPOSITED:", bool(exstyle & WS_EX_COMPOSITED))
print("has WS_EX_LAYERED:", bool(exstyle & WS_EX_LAYERED))
print("has WS_POPUP:", bool(style & 0x80000000))

# 实验：启用 WS_EX_COMPOSITED，观察窗口是否还能正常绘制与更新
new_ex = exstyle | WS_EX_COMPOSITED
user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_ex)
user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                    0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0100)  # NOMOVE NOSIZE NOZORDER FRAMECHANGED NOACTIVATE
root.update()
time.sleep(0.3)

# 验证窗口仍可交互、控件仍可见（离屏像素检查走 app 自身截图能力太重，直接看 Tk 是否存活）
for i in range(5):
    root.update()
    time.sleep(0.05)
print("after COMPOSITED: exstyle=0x%08X" % user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE))
print("window alive, tk ok:", root.winfo_exists())

# 尝试改变几何验证重绘正常
root.geometry("1200x800+60+60")
root.update()
time.sleep(0.2)
root.update()
print("geometry change ok:", root.winfo_width(), root.winfo_height())

app.shutting_down = True
root.destroy()
print("DONE")
