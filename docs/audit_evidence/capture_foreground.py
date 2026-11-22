# -*- coding: utf-8 -*-
r"""强制前台 + BitBlt 屏幕截图（WebView2 内容需真实屏幕采集）。

用法：python capture_foreground.py <title> <out.png>
"""
import ctypes
import sys
import time

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
shcore = ctypes.windll.shcore

# DPI 感知：否则 GetWindowRect 返回虚拟像素，高 DPI 屏上截图只抓到局部
try:
    shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
except Exception:
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

user32.FindWindowW.restype = ctypes.c_void_p
user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
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


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def main():
    title = sys.argv[1] if len(sys.argv) > 1 else "PDF 批量双面打印工具"
    out = sys.argv[2] if len(sys.argv) > 2 else "window.png"
    hwnd = user32.FindWindowW(None, title)
    if not hwnd:
        print("window not found")
        return 1
    user32.ShowWindow(ctypes.c_void_p(hwnd), 9)  # SW_RESTORE
    user32.SetForegroundWindow(ctypes.c_void_p(hwnd))
    time.sleep(1.2)  # 等 DWM 合成稳定
    rect = RECT()
    user32.GetWindowRect(ctypes.c_void_p(hwnd), ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    print("rect:", rect.left, rect.top, w, h)
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
    if not taken:
        print("GetDIBits failed")
        return 1
    img = Image.frombytes("RGBA", (w, h), buf.raw, "raw", "BGRA")
    img.convert("RGB").save(out)
    print("saved:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
