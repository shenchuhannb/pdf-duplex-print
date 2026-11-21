# -*- coding: utf-8 -*-
"""重建桌面快捷方式 + 刷新 Shell 图标缓存 + 提取“系统实际取到的图标”。

用法（通常在 venv 里运行）：
    python tools/create_shortcut.py --target ".venv\\Scripts\\pythonw.exe"

做的事：
  1. 用 WScript.Shell COM *重新创建* .lnk（不是只改属性），写回后逐项读回校验；
  2. 触发 SHChangeNotify(SHCNE_ASSOCCHANGED)，让资源管理器丢掉旧的图标缓存；
     绝不重启 explorer.exe（会打断用户当前桌面会话）；
  3. 用 SHGetFileInfo 取资源管理器“真正会显示”的 32/16px 图标，
     并把 ICO 里的 256/48 条目一起拼成 docs/shortcut_icon_check.png，便于人工比对；
  4. 打印 ICO 逐条格式校验结果（小尺寸必须 BMP，256 才允许 PNG）。
"""

import argparse
import ctypes
import os
import struct
import sys

from ctypes import wintypes

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from app_identity import APP_USER_MODEL_ID  # noqa: E402
from make_icons import verify_ico  # noqa: E402

DEFAULT_LNK = os.path.join(os.path.expanduser("~"), "Desktop", "PDF双面打印.lnk")

SHCNE_ASSOCCHANGED = 0x08000000
SHCNF_IDLIST = 0x0000
SHGFI_ICON = 0x00000100
SHGFI_LARGEICON = 0x00000000
SHGFI_SMALLICON = 0x00000001
SHGFI_USEFILEATTRIBUTES = 0x00000010


class SHFILEINFO(ctypes.Structure):
    _fields_ = [
        ("hIcon", wintypes.HICON),
        ("iIcon", ctypes.c_int),
        ("dwAttributes", wintypes.DWORD),
        ("szDisplayName", wintypes.WCHAR * 260),
        ("szTypeName", wintypes.WCHAR * 80),
    ]


def rebuild_shortcut(lnk_path, target, arguments, workdir, icon, description, style):
    import pythoncom
    from win32com.client import Dispatch

    pythoncom.CoInitialize()
    try:
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        shortcut.TargetPath = target
        shortcut.Arguments = arguments
        shortcut.WorkingDirectory = workdir
        shortcut.IconLocation = icon
        shortcut.Description = description
        shortcut.WindowStyle = style
        shortcut.Save()
        return {
            "TargetPath": shortcut.TargetPath,
            "Arguments": shortcut.Arguments,
            "WorkingDirectory": shortcut.WorkingDirectory,
            "IconLocation": shortcut.IconLocation,
            "Description": shortcut.Description,
            "WindowStyle": shortcut.WindowStyle,
        }
    finally:
        pythoncom.CoUninitialize()


def set_lnk_aumid(lnk_path: str, aumid: str) -> None:
    """写入 System.AppUserModel.ID：任务栏/开始菜单固定项与进程 AUMID 一致，
    图标才不会回退到宿主 pythonw.exe 的图标。"""
    import pythoncom
    from win32com.propsys import propsys, pscon

    pythoncom.CoInitialize()
    try:
        store = propsys.SHGetPropertyStoreFromParsingName(
            lnk_path, None, 2, propsys.IID_IPropertyStore)   # 2 = GPS_READWRITE
        store.SetValue(pscon.PKEY_AppUserModel_ID, propsys.PROPVARIANTType(aumid))
        store.Commit()
    finally:
        pythoncom.CoUninitialize()


def read_lnk_aumid(lnk_path: str) -> str:
    """读回 .lnk 的 System.AppUserModel.ID（未设置时返回空字符串）。"""
    from win32com.propsys import propsys, pscon

    store = propsys.SHGetPropertyStoreFromParsingName(
        lnk_path, None, 2, propsys.IID_IPropertyStore)
    value = store.GetValue(pscon.PKEY_AppUserModel_ID)
    try:
        return value.ToString() if value.vt else ""
    except Exception:
        return ""


def notify_shell():
    """SHCNE_ASSOCCHANGED：让资源管理器重新读取图标缓存（不重启 explorer）。"""
    ctypes.windll.shell32.SHChangeNotify(
        SHCNE_ASSOCCHANGED, SHCNF_IDLIST, ctypes.c_void_p(0), ctypes.c_void_p(0))


def _hicon_to_rgba(hicon, size):
    """把 HICON 画到 32bpp 位图上，返回随 RGB 覆盖的 RGBA 图像。"""
    import win32con
    import win32gui
    import win32ui
    from PIL import Image

    hdc = win32gui.GetDC(0)
    source = win32ui.CreateDCFromHandle(hdc)
    mem = source.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(source, size, size)
    mem.SelectObject(bmp)
    mem.FillSolidRect((0, 0, size, size), 0x00FFFFFF)  # 白底，便于人工比对
    win32gui.DrawIconEx(mem.GetSafeHdc(), 0, 0, hicon, size, size,
                        0, None, win32con.DI_NORMAL)
    info = bmp.GetInfo()
    bits = bmp.GetBitmapBits(True)
    image = Image.frombuffer("RGBA", (info["bmWidth"], info["bmHeight"]),
                             bits, "raw", "BGRA", 0, 1)
    # 兼容位图的 alpha 恒为 0；这里以白底结果为准确内容，统一置为不透明。
    image.putalpha(255)
    mem.DeleteDC()
    source.DeleteDC()
    win32gui.ReleaseDC(0, hdc)
    win32gui.DeleteObject(bmp.GetHandle())
    return image


def shell_icon(path, small=False):
    """SHGetFileInfo：返回资源管理器对该路径实际会显示的图标（HICON）。"""
    info = SHFILEINFO()
    flags = SHGFI_ICON | (SHGFI_SMALLICON if small else SHGFI_LARGEICON)
    ok = ctypes.windll.shell32.SHGetFileInfoW(
        ctypes.c_wchar_p(path), 0, ctypes.byref(info), ctypes.sizeof(info), flags)
    if not ok or not info.hIcon:
        return None, 0
    size = 16 if small else 32
    image = _hicon_to_rgba(info.hIcon, size)
    ctypes.windll.user32.DestroyIcon(info.hIcon)
    return image, size


def _label_font(size=14):
    from PIL import ImageFont

    for candidate in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf"):
        if os.path.isfile(candidate):
            try:
                return ImageFont.truetype(candidate, size)
            except OSError:
                pass
    return ImageFont.load_default()


def compose_check(shell_large, shell_small, ico_images, out_path):
    """拼成对照图：系统取到的 32/16px + ICO 内的 256/48px。"""
    from PIL import Image, ImageDraw

    cell, pad, label_h = 200, 16, 28
    panels = [
        ("SHGetFileInfo(.lnk) 32px", shell_large),
        ("SHGetFileInfo(.lnk) 16px", shell_small),
        ("ICO 256px", ico_images.get(256)),
        ("ICO 48px", ico_images.get(48)),
    ]
    width = pad + len(panels) * (cell + pad)
    height = pad + label_h + cell + pad
    canvas = Image.new("RGB", (width, height), (238, 241, 245))
    draw = ImageDraw.Draw(canvas)
    font = _label_font(14)
    for index, (label, image) in enumerate(panels):
        x = pad + index * (cell + pad)
        draw.text((x, pad), label, fill=(31, 41, 51), font=font)
        box_y = pad + label_h
        draw.rectangle((x, box_y, x + cell, box_y + cell), fill=(255, 255, 255),
                       outline=(215, 221, 229))
        if image is not None:
            limit = cell - 16
            if image.width <= limit // 2:
                # 小图标放大展示（最近邻，避免糊成一片）
                factor = max(2, limit // max(1, image.width))
                shown = image.resize((image.width * factor, image.height * factor),
                                     Image.NEAREST)
            elif image.width > limit:
                shown = image.resize((limit, limit), Image.LANCZOS)
            else:
                shown = image
            canvas.paste(shown,
                         (x + (cell - shown.width) // 2, box_y + (cell - shown.height) // 2),
                         shown if shown.mode == "RGBA" else None)
    canvas.save(out_path)
    return out_path


def ico_entries(path, keep=(256, 48)):
    """从 ICO 里按尺寸取出条目（256 走 PNG，48 走 BMP），用于对照。"""
    from PIL import Image

    with open(path, "rb") as handle:
        data = handle.read()
    count = struct.unpack("<HHH", data[:6])[2]
    result = {}
    for index in range(count):
        entry = data[6 + 16 * index: 6 + 16 * (index + 1)]
        width_b, height_b, colors, resv, planes, bitcount, size, offset = struct.unpack(
            "<BBBBHHII", entry)
        width = width_b or 256
        if width not in keep:
            continue
        blob = data[offset: offset + size]
        if blob[:8] == b"\x89PNG\r\n\x1a\n":
            import io

            result[width] = Image.open(io.BytesIO(blob)).convert("RGBA")
        else:
            from make_icons import _bmp_entry  # 仅用于结构参考

            result[width] = _dib_to_image(blob, width)
    return result


def _dib_to_image(blob, width):
    from PIL import Image

    height = struct.unpack("<i", blob[8:12])[0] // 2
    xor = blob[40:40 + width * height * 4]
    image = Image.frombuffer("RGBA", (width, height), xor, "raw", "BGRA", 0, -1)
    return image.transpose(Image.FLIP_TOP_BOTTOM)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--lnk", default=DEFAULT_LNK)
    parser.add_argument("--target", default=os.path.join(
        ROOT, ".venv", "Scripts", "pythonw.exe"))
    parser.add_argument("--arguments", default="pdf_duplex_print.py")
    parser.add_argument("--workdir", default=ROOT)
    parser.add_argument("--icon", default=os.path.join(ROOT, "app_icon.ico"))
    parser.add_argument("--description", default="PDF 批量手动双面打印工具")
    parser.add_argument("--aumid", default=APP_USER_MODEL_ID)
    parser.add_argument("--style", type=int, default=1)
    parser.add_argument("--no-notify", action="store_true")
    args = parser.parse_args()

    ico = os.path.abspath(args.icon)
    print("ICO 逐条校验:")
    for item in verify_ico(ico):
        print("  {0}x{1} {2} {3}bpp {4} bytes".format(*item))

    props = rebuild_shortcut(args.lnk, os.path.abspath(args.target), args.arguments,
                             os.path.abspath(args.workdir), f"{ico},0",
                             args.description, args.style)
    print("快捷方式已重新创建:", args.lnk)
    for key, value in props.items():
        print(f"  {key} = {value}")

    set_lnk_aumid(args.lnk, args.aumid)
    print(f"  System.AppUserModel.ID = {read_lnk_aumid(args.lnk)!r}")

    if not args.no_notify:
        notify_shell()
        print("已触发 SHChangeNotify(SHCNE_ASSOCCHANGED)")

    os.makedirs(DOCS, exist_ok=True)
    out = os.path.join(DOCS, "shortcut_icon_check.png")
    large, _ = shell_icon(args.lnk)
    small, _ = shell_icon(args.lnk, small=True)
    compose_check(large, small, ico_entries(ico), out)
    print("系统实际图标对照已保存:", out)


if __name__ == "__main__":
    main()
