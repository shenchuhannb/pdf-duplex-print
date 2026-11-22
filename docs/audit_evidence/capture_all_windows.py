# -*- coding: utf-8 -*-
r"""枚举所有顶层窗口，找到指定标题的全部实例并逐个截图。

用法：python capture_all_windows.py <title> <outdir>
"""
import ctypes
import os
import sys
import time

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

user32.GetWindowTextW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
user32.EnumWindows.argtypes = [ctypes.c_void_p, ctypes.c_ssize_t]
user32.GetDC.restype = ctypes.c_void_p
user32.GetDC.argtypes = [ctypes.c_void_p]
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
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


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


matches = []


@ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_ssize_t)
def on_window(hwnd, _lparam):
    if not user32.IsWindowVisible(hwnd):
        return True
    buf = ctypes.create_unicode_buffer(256)
    user32.GetWindowTextW(hwnd, buf, 256)
    if buf.value == TARGET:
        matches.append(hwnd)
    return True


def grab(hwnd, out):
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return
    hdc = user32.GetDC(None)
    memdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    old = gdi32.SelectObject(memdc, bmp)
    gdi32.BitBlt(memdc, 0, 0, w, h, hdc, rect.left, rect.top, 0x00CC0020)
    from PIL import Image
    import ctypes as ct

    class BMIH(ctypes.Structure):
        _fields_ = [("biSize", ct.c_uint), ("biWidth", ct.c_int),
                    ("biHeight", ct.c_int), ("biPlanes", ct.c_ushort),
                    ("biBitCount", ct.c_ushort), ("biCompression", ct.c_uint),
                    ("biSizeImage", ct.c_uint), ("biXPelsPerMeter", ct.c_int),
                    ("biYPelsPerMeter", ct.c_int), ("biClrUsed", ct.c_uint),
                    ("biClrImportant", ct.c_uint)]

    gdi32.GetDIBits.restype = ctypes.c_int
    gdi32.GetDIBits.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ct.c_uint, ct.c_uint,
                                ct.c_void_p, ct.c_void_p, ct.c_uint]
    bmi = BMIH()
    bmi.biSize = ctypes.sizeof(BMIH)
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
    user32.ReleaseDC(None, hdc)
    if taken:
        img = Image.frombytes("RGBA", (w, h), buf.raw, "raw", "BGRA")
        img.convert("RGB").save(out)
        print("saved:", out, f"({w}x{h} at {rect.left},{rect.top})")


TARGET = sys.argv[1] if len(sys.argv) > 1 else "PDF 批量双面打印工具"
outdir = sys.argv[2] if len(sys.argv) > 2 else "."
user32.EnumWindows(on_window, 0)
print("found windows:", len(matches))
for i, hwnd in enumerate(matches):
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.8)
    grab(hwnd, os.path.join(outdir, f"win-{i}-0x{hwnd:016x}.png"))
