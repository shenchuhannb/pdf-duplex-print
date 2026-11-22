# -*- coding: utf-8 -*-
"""离屏验证最大化首帧不再出现大面积黑色；不显示窗口、不抓屏幕。"""
import ctypes
import os
import pathlib
import sys
import tempfile
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

import win32gui
import win32ui
from PIL import Image

TEMP = tempfile.mkdtemp(prefix="maximize-smoke-")
os.environ["APPDATA"] = TEMP

from ui_main import DuplexPrintApp, _create_root  # noqa: E402


def print_window(root) -> Image.Image:
    hwnd = win32gui.GetAncestor(root.winfo_id(), 2)
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width, height = right - left, bottom - top
    hdc = win32gui.GetDC(0)
    source = win32ui.CreateDCFromHandle(hdc)
    memory = source.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source, width, height)
    memory.SelectObject(bitmap)
    try:
        ctypes.windll.user32.PrintWindow(
            ctypes.c_void_p(hwnd), ctypes.c_void_p(memory.GetSafeHdc()), 2)
        info = bitmap.GetInfo()
        image = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                                 bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1)
        return image
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory.DeleteDC()
        source.DeleteDC()
        win32gui.ReleaseDC(0, hdc)


def black_ratio(image: Image.Image) -> float:
    pixels = image.convert("RGB").resize((160, 120)).getdata()
    dark = sum(1 for r, g, b in pixels if r < 18 and g < 18 and b < 18)
    return dark / (160 * 120)


root = _create_root()
root.withdraw()
app = DuplexPrintApp(root, layout_scheme="stacked")
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    poll_id = getattr(app, attr, None)
    if poll_id is not None:
        root.after_cancel(poll_id)
        setattr(app, attr, None)
root.attributes("-alpha", 0.0)
root.deiconify()
try:
    root.geometry("1024x720+40+40")
    root.update()
    before = black_ratio(print_window(root))
    root.state("zoomed")
    samples = []
    for _ in range(8):
        root.update()
        time.sleep(0.04)
        samples.append(black_ratio(print_window(root)))
    print(f"before={before:.4f} after_samples={[round(v, 4) for v in samples]}")
    print("ok=True" if max(samples) < 0.15 else "ok=False")
finally:
    app.shutting_down = True
    root.destroy()