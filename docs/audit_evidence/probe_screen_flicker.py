# -*- coding: utf-8 -*-
r"""屏幕级黑块/残影复现探针。

用真实应用路径启动（与桌面快捷方式一致），从屏幕 DC 连续抓帧，
在最大化瞬间与日志分隔条拖动期间检测异常像素。

运行：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_screen_flicker.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="flicker-screen-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.GetWindowRect.restype = ctypes.c_int
user32.GetDC.restype = ctypes.c_void_p
user32.GetDC.argtypes = [ctypes.c_void_p]
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
gdi32.BitBlt.restype = ctypes.c_int
gdi32.BitBlt.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                         ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                         ctypes.c_uint]
gdi32.GetDIBits.restype = ctypes.c_int
gdi32.GetDIBits.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint,
                            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
user32.PostMessageW.restype = ctypes.c_int
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                ctypes.c_ssize_t]
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                ctypes.c_ssize_t]


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint), ("biWidth", ctypes.c_int),
                ("biHeight", ctypes.c_int), ("biPlanes", ctypes.c_ushort),
                ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
                ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_int),
                ("biYPelsPerMeter", ctypes.c_int), ("biClrUsed", ctypes.c_uint),
                ("biClrImportant", ctypes.c_uint)]


def capture(hwnd, screen_dc):
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return w, h, None
    memdc = gdi32.CreateCompatibleDC(screen_dc)
    bmp = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
    old = gdi32.SelectObject(memdc, bmp)
    gdi32.BitBlt(memdc, 0, 0, w, h, screen_dc, rect.left, rect.top, 0x00CC0020)
    bmi = BITMAPINFOHEADER()
    bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.biWidth = w
    bmi.biHeight = -h
    bmi.biPlanes = 1
    bmi.biBitCount = 32
    bmi.biCompression = 0
    buf = ctypes.create_string_buffer(w * h * 4)
    taken = gdi32.GetDIBits(memdc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.SelectObject(memdc, old)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(memdc)
    if not taken:
        return w, h, None
    return w, h, buf.raw


def analyze(raw, w, h):
    """返回(近黑像素比例, 深暗但非纯黑比例)。"""
    dark = 0
    dim = 0
    total = w * h
    step = 4  # 每 4 像素采样一次
    sampled = 0
    for off in range(0, len(raw) - 3, 4 * step):
        b, g, r = raw[off], raw[off + 1], raw[off + 2]
        sampled += 1
        if r < 16 and g < 16 and b < 16:
            dark += 1
        elif r < 64 and g < 64 and b < 64:
            dim += 1
    return dark / sampled, dim / sampled


# ---------------------------------------------------------------- 启动真实应用
import tkinter as tk  # noqa: E402

from ui_main import DuplexPrintApp, _create_root  # noqa: E402

root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+60+60")
root.update()
hwnd = int(root.wm_frame(), 16)
print("hwnd: 0x%X" % hwnd)
time.sleep(0.6)
root.update()

screen_dc = user32.GetDC(None)

# ------------------------------------------------------------- 1) 最大化黑块
frames = []
WM_SYSCOMMAND = 0x0112
SC_MAXIMIZE = 0xF030
SC_RESTORE = 0xF120
user32.PostMessageW(ctypes.c_void_p(hwnd), WM_SYSCOMMAND, SC_MAXIMIZE, 0)
t0 = time.perf_counter()
worst = (0.0, 0.0, -1.0)
while time.perf_counter() - t0 < 1.6:
    root.update()
    w, h, raw = capture(hwnd, screen_dc)
    if raw:
        ratio_dark, ratio_dim = analyze(raw, w, h)
        t = time.perf_counter() - t0
        frames.append((t, ratio_dark, ratio_dim))
        if ratio_dark > worst[0]:
            worst = (ratio_dark, ratio_dim, t)
    time.sleep(1 / 60)
print("MAXIMIZE frames:", len(frames))
for t, d, m in frames[:12]:
    print("  t=%.3f dark=%.4f dim=%.4f" % (t, d, m))
print("WORST dark=%.4f dim=%.4f at t=%.3f" % worst)
user32.PostMessageW(ctypes.c_void_p(hwnd), WM_SYSCOMMAND, SC_RESTORE, 0)
time.sleep(0.8)
root.update()

# --------------------------------------------------- 2) 日志分隔条快速拖动
# 直接调用与真实拖动相同的路径：_set_log_height_during_drag
import win32con  # noqa: E402

drag_frames = []
heights = list(range(134, 320, 6)) + list(range(320, 134, -6))
t0 = time.perf_counter()
for value in heights:
    app._set_log_height_during_drag(value)
    root.update()
    w, h, raw = capture(hwnd, screen_dc)
    if raw:
        d, m = analyze(raw, w, h)
        drag_frames.append((d, m))
    time.sleep(1 / 90)
print("DRAG frames:", len(drag_frames))
worst_drag = max(drag_frames) if drag_frames else (0, 0)
print("WORST drag dark=%.4f dim=%.4f" % worst_drag)

user32.ReleaseDC(None, screen_dc)
app.shutting_down = True
root.destroy()
print("DONE")
