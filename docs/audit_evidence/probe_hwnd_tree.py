# -*- coding: utf-8 -*-
r"""探明 Tk toplevel 的 HWND 层级结构：wrapper 与子窗口类名、矩形。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_hwnd_tree.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="hwnd-tree-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

import win32con
import win32gui
import win32api

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

wrapper = int(root.wm_frame(), 16)
print("wrapper hwnd=0x%X class=%s" % (wrapper, win32gui.GetClassName(wrapper)))


def dump(hwnd, depth):
    cls = win32gui.GetClassName(hwnd)
    rect = win32gui.GetWindowRect(hwnd)
    text = win32gui.GetWindowText(hwnd)[:30]
    print("%s hwnd=0x%X cls=%s rect=%s text=%r" % (
        "  " * depth, hwnd, cls, rect, text))
    children = []

    def collect(child, _param):
        children.append(child)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, collect, None)
    except win32gui.error:
        pass
    for child in children:
        dump(child, depth + 1)


dump(wrapper, 0)
app.shutting_down = True
root.destroy()
print("DONE")
