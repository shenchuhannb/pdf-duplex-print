# -*- coding: utf-8 -*-
r"""诊断：验证 win32_flicker 子类化在真实应用路径上是否生效。

用与桌面快捷方式相同的解释器运行：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_live_subclass.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="flicker-probe-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

import tkinter as tk

import win32_flicker
from ui_main import DuplexPrintApp, _create_root
import theme

print("python:", sys.executable)
root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
root.update()  # 触发 Map → _on_root_map_erase_fix → install

print("installed_entries:", len(win32_flicker._installed))
hwnd = int(root.wm_frame(), 16)
print("hwnd: 0x%X" % hwnd)
print("in_installed:", hwnd in win32_flicker._installed)

# 直接验证 WndProc 已被替换且我们的处理生效：
# 向窗口发送 WM_ERASEBKGND，用一个内存 DC 检查是否被填充为底色。
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

WM_ERASEBKGND = 0x0014


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


hwnd_c = ctypes.c_void_p(hwnd)
user32.GetDC.restype = ctypes.c_void_p
user32.GetDC.argtypes = [ctypes.c_void_p]
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                ctypes.c_size_t, ctypes.c_ssize_t]
gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.GetPixel.restype = ctypes.c_ulong
gdi32.GetPixel.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
rect = RECT()
user32.GetClientRect(hwnd_c, ctypes.byref(rect))
w, h = rect.right - rect.left, rect.bottom - rect.top
hdc = user32.GetDC(hwnd_c)
memdc = gdi32.CreateCompatibleDC(hdc)
membmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
old = gdi32.SelectObject(memdc, membmp)
# 先填成纯红，便于检测我们的画刷是否覆盖它
red = gdi32.CreateSolidBrush(0x0000FF)
gdi32.CreateSolidBrush.restype = ctypes.c_void_p
user32.FillRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
user32.FillRect(memdc, ctypes.byref(rect), red)
res = user32.SendMessageW(hwnd_c, WM_ERASEBKGND, memdc, 0)
print("wm_erasebkgnd_result:", res)

center = gdi32.GetPixel(memdc, w // 2, h // 2)
r, g, b = center & 0xFF, (center >> 8) & 0xFF, (center >> 16) & 0xFF
print("center_rgb_after_erase: #%02X%02X%02X" % (r, g, b))
expected = app.colors["bg"]
print("expected_bg:", expected)
gdi32.SelectObject(memdc, old)
gdi32.DeleteObject(membmp)
gdi32.DeleteDC(memdc)
user32.ReleaseDC(hwnd_c, hdc)

match = "#%02X%02X%02X" % (r, g, b)
print("ERASE_WORKS:", match.upper() == expected.upper())

app.shutting_down = True
root.destroy()
