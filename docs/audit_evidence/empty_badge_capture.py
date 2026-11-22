# -*- coding: utf-8 -*-
"""离屏输出当前空状态图标截图，供人工复核轮廓和中心。"""
import ctypes
import os
import pathlib
import sys
import tempfile

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

import win32gui
import win32ui
from PIL import Image

os.environ["APPDATA"] = tempfile.mkdtemp(prefix="empty-badge-smoke-")
from ui_main import DuplexPrintApp, _create_root  # noqa: E402


def capture(root) -> Image.Image:
    hwnd = win32gui.GetAncestor(root.winfo_id(), 2)
    _l, _t, right, bottom = win32gui.GetClientRect(hwnd)
    hdc = win32gui.GetDC(0)
    source = win32ui.CreateDCFromHandle(hdc)
    memory = source.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source, right, bottom)
    memory.SelectObject(bitmap)
    try:
        ctypes.windll.user32.PrintWindow(
            ctypes.c_void_p(hwnd), ctypes.c_void_p(memory.GetSafeHdc()), 2)
        info = bitmap.GetInfo()
        return Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                                bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory.DeleteDC()
        source.DeleteDC()
        win32gui.ReleaseDC(0, hdc)


root = _create_root()
root.withdraw()
app = DuplexPrintApp(root, layout_scheme="stacked", ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    poll_id = getattr(app, attr, None)
    if poll_id is not None:
        root.after_cancel(poll_id)
        setattr(app, attr, None)
root.attributes("-alpha", 0.0)
root.geometry("1024x720+40+40")
root.deiconify()
try:
    root.update()
    full = capture(root)
    crop = full
    target = pathlib.Path(tempfile.gettempdir()) / "codearts-current-empty-badge.png"
    crop.save(target)
    print(target)
finally:
    app.shutting_down = True
    root.destroy()