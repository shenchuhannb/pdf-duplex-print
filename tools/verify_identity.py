# -*- coding: utf-8 -*-
"""任务栏图标/AUMID 验证（全程离屏，不上屏）。

流程：
  1. 设置进程 AUMID 并读回；
  2. 以透明窗口（alpha=0）构建主界面，由 app_identity.apply_window_icon 设置
     WM_SETICON(ICON_BIG/ICON_SMALL) 与窗口类图标；
  3. 查询窗口实际 HICON（WM_GETICON 三种 + GCLP_HICON/GCLP_HICONS），
     导出 docs/taskbar_icon_big.png / docs/taskbar_icon_small.png；
  4. 与 app_icon.ico 同尺寸帧做像素对比，确认不是 pythonw 图标。

用法：python tools/verify_identity.py [--assert]
"""

import ctypes
import os
import sys
import tempfile
from ctypes import wintypes

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32gui
import win32ui
from PIL import Image, ImageChops, ImageStat

import app_identity
from ui_main import DuplexPrintApp, _create_root

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
GA_ROOT = 2
WM_GETICON = 0x007F
ICON_SMALL, ICON_BIG, ICON_SMALL2 = 0, 1, 2
GCLP_HICON, GCLP_HICONS = -14, -34

_user32 = ctypes.windll.user32
_user32.SendMessageW.restype = ctypes.c_ssize_t
_user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_user32.GetClassLongPtrW.restype = ctypes.c_ssize_t
_user32.DrawIcon.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.HANDLE]


def export_hicon(hicon: int, size: int, out_path: str) -> bool:
    """把 HICON 画到白底位图并保存 PNG（离屏，不依赖屏幕像素）。"""
    if not hicon:
        return False
    screen_dc = win32gui.GetDC(0)
    source = win32ui.CreateDCFromHandle(screen_dc)
    memory = source.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source, size, size)
    memory.SelectObject(bitmap)
    memory.FillSolidRect((0, 0, size, size), 0xFFFFFF)
    ok = bool(_user32.DrawIcon(memory.GetSafeHdc(), 0, 0, hicon))
    bmp_path = out_path + ".bmp"
    try:
        bitmap.SaveBitmapFile(memory, bmp_path)
        image = Image.open(bmp_path).convert("RGB")
        if size != image.size[0]:
            image = image.resize((size, size), Image.LANCZOS)
        image.save(out_path, optimize=True)
    finally:
        try:
            os.remove(bmp_path)
        except OSError:
            pass
        win32gui.DeleteObject(bitmap.GetHandle())
        memory.DeleteDC()
        source.DeleteDC()
        win32gui.ReleaseDC(0, screen_dc)
    return ok


def ico_frame(size: int, path: str) -> Image.Image:
    """取 .ico 中指定尺寸的帧（缺失时用最大帧缩放），合成到白底。"""
    image = Image.open(path)
    image.size = (size, size)
    image.load()
    return image.convert("RGB")


def main() -> int:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    app_identity.set_process_app_user_model_id()
    aumid = app_identity.get_process_app_user_model_id()

    root = _create_root()
    root.withdraw()
    app = DuplexPrintApp(root, style="A")
    root.geometry("1280x960+40+40")
    root.attributes("-alpha", 0.0)
    root.update_idletasks()
    root.deiconify()
    root.update()

    hwnd = _user32.GetAncestor(int(root.winfo_id()), GA_ROOT)
    handles = {
        "WM_GETICON ICON_BIG": _user32.SendMessageW(hwnd, WM_GETICON, ICON_BIG, 0),
        "WM_GETICON ICON_SMALL": _user32.SendMessageW(hwnd, WM_GETICON, ICON_SMALL, 0),
        "WM_GETICON ICON_SMALL2": _user32.SendMessageW(hwnd, WM_GETICON, ICON_SMALL2, 0),
        "GCLP_HICON": _user32.GetClassLongPtrW(hwnd, GCLP_HICON),
        "GCLP_HICONS": _user32.GetClassLongPtrW(hwnd, GCLP_HICONS),
    }
    ico = os.path.join(ROOT, "app_icon.ico")
    big = os.path.join(DOCS, "taskbar_icon_big.png")
    small = os.path.join(DOCS, "taskbar_icon_small.png")
    os.makedirs(DOCS, exist_ok=True)
    big_ok = export_hicon(handles["WM_GETICON ICON_BIG"] or handles["GCLP_HICON"], 32, big)
    small_ok = export_hicon(
        handles["WM_GETICON ICON_SMALL2"] or handles["WM_GETICON ICON_SMALL"]
        or handles["GCLP_HICONS"], 16, small)

    diffs = {}
    for name, path, size in (("big", big, 32), ("small", small, 16)):
        if not os.path.isfile(path):
            continue
        actual = Image.open(path).convert("RGB")
        reference = ico_frame(size, ico)
        diff = ImageStat.Stat(ImageChops.difference(actual, reference)).mean
        diffs[name] = sum(diff) / len(diff)

    print("AUMID =", repr(aumid))
    for key, value in handles.items():
        print(f"{key} = 0x{int(value or 0):X}")
    app_handles = getattr(app, "icon_handles", {})
    print("apply_window_icon handles =",
          {k: f"0x{int(v):X}" for k, v in app_handles.items()})
    from_file = (handles["WM_GETICON ICON_BIG"] == app_handles.get("big")
                 and (handles["WM_GETICON ICON_SMALL2"] == app_handles.get("small")
                      or handles["WM_GETICON ICON_SMALL"] == app_handles.get("small")))
    print("window HICON from app_icon.ico :", from_file)
    print(f"export big={big_ok} small={small_ok}")
    print("pixel diff vs app_icon.ico :", {k: round(v, 2) for k, v in diffs.items()})

    app.shutting_down = True
    app.preview_enabled = False
    root.destroy()

    if "--assert" in sys.argv:
        ok = (aumid == app_identity.APP_USER_MODEL_ID and big_ok and small_ok and from_file
              and diffs.get("big", 999) < 40)
        print("ASSERT", "PASS" if ok else "FAIL")
        return 0 if ok else 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
