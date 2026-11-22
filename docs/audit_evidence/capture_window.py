# -*- coding: utf-8 -*-
r"""截取指定标题窗口到 PNG。

用法：.venv\Scripts\python.exe docs\audit_evidence\capture_window.py <title-substr> <out.png>
"""
import ctypes
import sys
import time

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

user32.FindWindowW.restype = ctypes.c_void_p
user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
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
        print("window not found:", title)
        return 1
    # PrintWindow：直接从窗口表面读取，不受遮挡/前台状态影响（PW_RENDERFULLCONTENT=2）
    user32.PrintWindow.restype = ctypes.c_int
    user32.PrintWindow.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
    rect = RECT()
    user32.GetWindowRect(ctypes.c_void_p(hwnd), ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    print("rect:", rect.left, rect.top, w, h)
    hdc = user32.GetDC(None)
    memdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    old = gdi32.SelectObject(memdc, bmp)
    ok = user32.PrintWindow(ctypes.c_void_p(hwnd), ctypes.c_void_p(memdc), 2)
    if not ok:
        print("PrintWindow failed, fallback to BitBlt")
        user32.SetForegroundWindow(ctypes.c_void_p(hwnd))
        time.sleep(0.6)
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
