# -*- coding: utf-8 -*-
"""Word（.doc/.docx）→ 临时 PDF 转换。

为什么不在“直接打印 Word”上继续：直接打印时，Word 把页面交给打印机驱动的
GDI 路径，字体是否可用、是否被驱动替换取决于打印机自带字体/系统字体；WPS 与
Office 的排版引擎也有差异。结果是字体替换、符号乱码、分页漂移。先转 PDF 把
排版固定为字形轮廓，再走本工具现有的 PDF 排版/打印流程，输出才可控。

转换能力优先级（依据见 README/汇报）：
  1. Word COM  ``ExportAsFixedFormat``：保真最好，不依赖 Acrobat（本机已装
     Microsoft Office 2021 x64，WINWORD.EXE 存在）；
  2. WPS COM    ``ExportAsFixedFormat``，失败再 ``SaveAs2(..., 17)``（本机未装）；
  3. Acrobat PDFMaker：依赖 Word 加载项且 ProgID 未注册，仅探测不启用。

所有转换都在独立子进程里执行：
  * 可以强制超时（kill 子进程 + 结束本次新起的 WINWORD/WPS），不会卡住界面；
  * COM 初始化发生在这个子进程自己的线程里（CoInitialize），不污染 GUI 线程。
"""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from typing import Callable, Dict, List, Optional

WORD_PROGIDS = ("Word.Application", "Word.Application.16")
WPS_PROGIDS = ("KWPS.Application", "WPS.Application")
WD_EXPORT_FORMAT_PDF = 17          # wdExportFormatPDF
WD_DO_NOT_SAVE_CHANGES = 0
OFFICE_PROCESS_NAMES = ("WINWORD.EXE", "wps.exe")
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

# 转换产物的专用临时目录：固定前缀 + 时间戳 + PID，便于启动时识别上次残留。
# 不与系统临时目录里的其它文件混放，清理时只删除本工具自己的目录。
WORD_TEMP_PREFIX = "pdf_duplex_word_"
# 打印阶段生成的排版/合并 PDF 也放在这个专用根目录下：
# 崩溃/断电后留下的目录能被启动清理识别（否则会一直堆在系统临时目录里）。
PRINT_TEMP_PREFIX = "pdf_duplex_print_"
TEMP_PREFIXES = (WORD_TEMP_PREFIX, PRINT_TEMP_PREFIX)
WORD_TEMP_ROOT_NAME = "PDFDuplexPrint-temp"


class WordConvertError(Exception):
    """转换失败；message 为可直接展示给用户的中文提示。"""


def is_wps_file(path: str) -> bool:
    return str(path).strip().lower().endswith(".wps")


# ------------------------------------------------------- 临时目录与残留清理
def temp_root() -> str:
    """本工具专用的临时根目录（%TEMP%\\PDFDuplexPrint-temp）。"""
    return os.path.join(tempfile.gettempdir(), WORD_TEMP_ROOT_NAME)


def create_temp_dir(prefix: str = WORD_TEMP_PREFIX) -> str:
    """新建带前缀 + 时间戳 + PID 的临时目录，返回绝对路径。"""
    root = temp_root()
    os.makedirs(root, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    return tempfile.mkdtemp(prefix=f"{prefix}{stamp}_{os.getpid()}_", dir=root)


def _pid_alive(pid: int) -> bool:
    """只查询进程是否存在（不结束、不影响任何用户进程）。"""
    if pid <= 0:
        return False
    try:
        import ctypes

        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:
        return False


def _dir_pid(name: str) -> int:
    """从目录名解析创建者 PID；解析失败返回 0。

    目录名格式：pdf_duplex_{word|print}_YYYYmmdd_HHMMSS_<PID>_<随机后缀>
    """
    parts = name.split("_")
    if len(parts) < 6:
        return 0
    try:
        return int(parts[5])
    except ValueError:
        return 0


def cleanup_stale_temp_dirs() -> List[str]:
    """启动清理：删除上次运行残留的转换/打印临时目录。

    仍被其它运行中的本工具实例占用的目录（PID 存活）会跳过，
    只删除进程已退出的残留（含上次崩溃/被强制关闭的情况）。
    """
    root = temp_root()
    if not os.path.isdir(root):
        return []
    removed: List[str] = []
    for name in os.listdir(root):
        if not name.startswith(TEMP_PREFIXES):
            continue
        path = os.path.join(root, name)
        if not os.path.isdir(path):
            continue
        pid = _dir_pid(name)
        if pid == os.getpid():
            continue                      # 本次运行自己的目录（启动时通常为空）
        if pid and _pid_alive(pid):
            continue                      # 其它正在运行实例的目录
        shutil.rmtree(path, ignore_errors=True)
        if not os.path.exists(path):
            removed.append(name)
    return removed


# ------------------------------------------------------------------ 能力探测
def progid_registered(progid: str) -> bool:
    """只读注册表判断 COM ProgID 是否可用，不启动任何应用进程。"""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"):
            return True
    except OSError:
        return False


def available_converters() -> List[str]:
    """返回本机可用的转换器名称列表（Word / WPS）。"""
    found: List[str] = []
    if any(progid_registered(p) for p in WORD_PROGIDS):
        found.append("Word")
    if any(progid_registered(p) for p in WPS_PROGIDS):
        found.append("WPS")
    return found


# ------------------------------------------------------------------ 字体预检
# 缺字体的典型表现是导出 PDF 时被替换成默认字体，进而在打印稿上出现方框/乱码。
# 这里在转换前做一次轻量预检：
#   * .docx：直接读 zip 内的 fontTable.xml / document.xml / styles.xml，
#     并解析 theme1.xml 的主题字体，不启动 Office；
#   * .doc：只能在 Word COM 打开后从 Styles/Content 的 Font.Name 取样；
#   * 与系统字体（注册表 Fonts 键 + GDI 字体枚举）比对。
# 预检只用于提示，默认不阻断转换；预检失败按“未能预检”处理，不上报误报。
_FONT_NAME_CACHE: Optional[set] = None
_PSEUDO_FONT_NAMES = frozenset({
    "defaultparagraphfont", "normal", "正文", "默认段落字体",
})

# 常见 PDF 基础字体：未嵌入也不会造成方框，避免误报噪音。
_BASE14_FONTS = frozenset({
    "helvetica", "helvetica-bold", "helvetica-oblique", "helvetica-boldoblique",
    "times-roman", "times-bold", "times-italic", "times-bolditalic",
    "courier", "courier-bold", "courier-oblique", "courier-boldoblique",
    "symbol", "zapfdingbats",
})


def _normalize_font(name) -> str:
    """字体名归一化：忽略大小写/空格/连字符，并去掉竖排前缀 @。"""
    text = html.unescape(str(name or "")).strip()
    while text.startswith("@"):
        text = text[1:]
    text = text.casefold()
    return re.sub(r"[ _\-]+", "", text)


def _clean_font(name) -> str:
    """保留展示用的字体名；过滤空串/竖排前缀/Word 伪字体。"""
    text = html.unescape(str(name or "")).strip().lstrip("@").strip()
    if not text:
        return ""
    if _normalize_font(text) in _PSEUDO_FONT_NAMES:
        return ""
    return text


def _add_registry_font(names: set, raw) -> None:
    text = str(raw or "").strip()
    if not text:
        return
    if re.search(r"\.(ttf|ttc|otf|fon|fnt|pfb)$", text, re.I):
        base = os.path.splitext(os.path.basename(text.replace("\\", "/")))[0]
        key = _normalize_font(base)
        if key:
            names.add(key)
        return
    # 注册表值形如 “SimSun & NSimSun (TrueType)”，按 & 拆成多个族名
    for part in text.split("&"):
        clean = re.sub(r"\s*\([^)]*\)\s*$", "", part).strip()
        key = _normalize_font(clean)
        if key:
            names.add(key)


def _registry_font_names() -> set:
    """从注册表 Fonts 键收集字体族名（含英文名，如 SimSun/Microsoft YaHei）。"""
    names: set = set()
    try:
        import winreg
    except Exception:
        return names
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, key_path)
        except OSError:
            continue
        try:
            count = winreg.QueryInfoKey(key)[1]
            for index in range(count):
                try:
                    entry, value, _kind = winreg.EnumValue(key, index)
                except OSError:
                    continue
                _add_registry_font(names, entry)
                _add_registry_font(names, value)
        finally:
            try:
                winreg.CloseKey(key)
            except Exception:
                pass
    return names


def _gdi_font_names() -> set:
    """用 GDI EnumFontFamiliesEx 枚举字体族名（含中文名，如 宋体/微软雅黑）。"""
    names: set = set()
    try:
        import ctypes
        from ctypes import wintypes

        class LOGFONTW(ctypes.Structure):
            _fields_ = [
                ("lfHeight", ctypes.c_long),
                ("lfWidth", ctypes.c_long),
                ("lfEscapement", ctypes.c_long),
                ("lfOrientation", ctypes.c_long),
                ("lfWeight", ctypes.c_long),
                ("lfItalic", ctypes.c_byte),
                ("lfUnderline", ctypes.c_byte),
                ("lfStrikeOut", ctypes.c_byte),
                ("lfCharSet", ctypes.c_byte),
                ("lfOutPrecision", ctypes.c_byte),
                ("lfClipPrecision", ctypes.c_byte),
                ("lfQuality", ctypes.c_byte),
                ("lfPitchAndFamily", ctypes.c_byte),
                ("lfFaceName", ctypes.c_wchar * 32),
            ]

        class ENUMLOGFONTEXW(ctypes.Structure):
            _fields_ = [
                ("elfLogFont", LOGFONTW),
                ("elfFullName", ctypes.c_wchar * 64),
                ("elfStyle", ctypes.c_wchar * 32),
                ("elfScript", ctypes.c_wchar * 32),
            ]

        callback_type = ctypes.WINFUNCTYPE(
            ctypes.c_int, ctypes.POINTER(ENUMLOGFONTEXW), ctypes.c_void_p,
            wintypes.DWORD, wintypes.LPARAM)

        def collect(elf, _metric, _font_type, _param):
            face = elf.contents.elfLogFont.lfFaceName.strip()
            full = elf.contents.elfFullName.strip()
            for value in (face, full):
                key = _normalize_font(value)
                if key:
                    names.add(key)
            return 1

        callback = callback_type(collect)
        hdc = ctypes.windll.user32.GetDC(None)
        logfont = LOGFONTW()
        logfont.lfCharSet = 1  # DEFAULT_CHARSET
        ctypes.windll.gdi32.EnumFontFamiliesExW(
            hdc, ctypes.byref(logfont), callback, 0, 0)
        ctypes.windll.user32.ReleaseDC(None, hdc)
    except Exception:
        return names
    return names


def installed_font_families(refresh: bool = False) -> set:
    """本机字体族名集合（已归一化）。注册表与 GDI 双来源，覆盖中英文名。"""
    global _FONT_NAME_CACHE
    if _FONT_NAME_CACHE is None or refresh:
        found = _registry_font_names() | _gdi_font_names()
        _FONT_NAME_CACHE = found
    return set(_FONT_NAME_CACHE)


def _xml_attr(tag: str, name: str) -> str:
    match = re.search(rf"""\b{name}\s*=\s*["']([^"']*)["']""", tag or "")
    return html.unescape(match.group(1)) if match else ""


def _theme_typefaces(theme_xml: str) -> Dict[tuple, str]:
    """解析 theme1.xml，返回 {(major|minor, latin|ea|cs): 字体名}。"""
    fonts: Dict[tuple, str] = {}
    for scope, tag in (("major", "majorFont"), ("minor", "minorFont")):
        match = re.search(rf"<a:{tag}\b[^>]*>(.*?)</a:{tag}>", theme_xml, re.S)
        if not match:
            continue
        block = match.group(1)
        for kind in ("latin", "ea", "cs"):
            slot = re.search(rf"<a:{kind}\b[^>]*>", block)
            name = _clean_font(_xml_attr(slot.group(0), "typeface")) if slot else ""
            if name:
                fonts[(scope, kind)] = name
        if ("ea" not in {key[1] for key in fonts if key[0] == scope}):
            # theme 的 <a:ea typeface=""> 常为空，中文回退到 Hans 脚本字体
            for tag_text in re.findall(r"<a:font\b[^>]*>", block):
                if _xml_attr(tag_text, "script").casefold() == "hans":
                    name = _clean_font(_xml_attr(tag_text, "typeface"))
                    if name:
                        fonts[(scope, "ea")] = name
                    break
    return fonts


def _theme_ref_font(ref: str, theme_fonts: Dict[tuple, str]) -> str:
    value = str(ref or "").strip().casefold()
    mapping = {
        "minorhansi": ("minor", "latin"), "majorhansi": ("major", "latin"),
        "minorascii": ("minor", "latin"), "majorascii": ("major", "latin"),
        "minoreastasia": ("minor", "ea"), "majoreastasia": ("major", "ea"),
        "minorbidi": ("minor", "cs"), "majorbidi": ("major", "cs"),
    }
    slot = mapping.get(value)
    return theme_fonts.get(slot, "") if slot else ""


def _fonts_from_rfonts(xml: str, theme_xml: str = ""):
    """只解析 ``<w:rFonts>`` 元素，返回 ``(字体集合, 是否出现该元素)``。

    不能直接匹配 ``w:eastAsia="..."``：``<w:lang w:eastAsia="zh-CN">`` 也带
    同名属性，会把语言代码误报成字体。
    """
    names: set = set()
    theme_refs: set = set()
    tags = re.findall(r"<w:rFonts\b[^>]*>", xml or "")
    for tag in tags:
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            clean = _clean_font(_xml_attr(tag, "w:" + attr))
            if clean:
                names.add(clean)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            ref = _xml_attr(tag, "w:" + attr).strip()
            if ref:
                theme_refs.add(ref.casefold())
    theme_fonts = _theme_typefaces(theme_xml or "")
    for ref in theme_refs:
        resolved = _theme_ref_font(ref, theme_fonts)
        if resolved:
            names.add(resolved)
    return names, bool(tags)


def _font_table_names(xml: str):
    """解析 fontTable 的 ``<w:font w:name>``，跳过已嵌入字体的条目。"""
    names: set = set()
    for match in re.finditer(r"<w:font\b([^>]*)>(.*?)</w:font>", xml or "", re.S):
        attrs, body = match.group(1), match.group(2)
        if re.search(r"<w:embed(?:Regular|Bold|Italic|BoldItalic)\b", body):
            continue                        # 字体已嵌入文档，不依赖系统安装
        clean = _clean_font(_xml_attr(attrs, "w:name"))
        if clean:
            names.add(clean)
    for match in re.finditer(r"<w:font\b([^>]*)/>", xml or ""):
        clean = _clean_font(_xml_attr(match.group(1), "w:name"))
        if clean:
            names.add(clean)
    return names, bool(re.search(r"<w:font\b", xml or ""))


def _docx_font_names(path: str):
    """从 .docx 的 zip 中收集文档引用的字体名，返回 ``(字体集合, 是否有可判源)``。

    ``有可判源`` 指确实读到了 rFonts 或 fontTable：只有在这种情况下
    “没有缺失”才是可靠结论；纯空文档仍按“未能预检”。
    """
    parts: Dict[str, str] = {}
    with zipfile.ZipFile(path) as archive:
        for part in ("word/document.xml", "word/styles.xml",
                     "word/fontTable.xml", "word/theme/theme1.xml"):
            try:
                parts[part] = archive.read(part).decode("utf-8", "ignore")
            except KeyError:
                continue
            except Exception:
                continue

    theme_xml = parts.get("word/theme/theme1.xml", "")
    names: set = set()
    for part in ("word/document.xml", "word/styles.xml"):
        found, _has_tags = _fonts_from_rfonts(parts.get(part, ""), theme_xml)
        names |= found
    table_names, has_table = _font_table_names(parts.get("word/fontTable.xml", ""))
    names |= table_names
    looked = bool(names) or has_table
    return names, looked


def inspect_document_fonts(path: str, installed=None) -> dict:
    """预检文档字体：返回 ``{"checked", "missing", "fonts"}``，永不抛异常。

    ``checked`` 为 False 表示“未能预检”（格式不支持、文件损坏、未取到字体），
    调用方不应据此报告缺失，也不要输出噪音日志。
    """
    result = {"checked": False, "missing": [], "fonts": []}
    if not str(path).strip().lower().endswith(".docx"):
        return result
    try:
        names, looked = _docx_font_names(path)
    except Exception:
        return result
    if not looked:
        return result
    known = installed_font_families() if installed is None else {
        _normalize_font(item) for item in installed}
    missing = sorted(name for name in names
                     if _normalize_font(name) not in known)
    result.update(checked=True, missing=missing, fonts=sorted(names))
    return result


def _inspect_doc_fonts_com(doc) -> dict:
    """在已打开的 .doc 上做轻量取样。

    优先一次读取 ``Content.WordOpenXML``（实测约 0.15s，含 rFonts 与
    fontTable），避免逐样式 COM 往返（400 个样式约 3s，会拖慢批量）。
    读不到时才退回 Content.Font.Name + 少量样式取样。
    """
    result = {"checked": False, "missing": [], "fonts": [], "source": "com"}
    names: set = set()
    try:
        xml = doc.Content.WordOpenXML
    except Exception:
        xml = ""
    if xml:
        found, _has_tags = _fonts_from_rfonts(xml)
        table_names, _has_table = _font_table_names(xml)
        names |= found | table_names
    if not names:
        # 退化路径：只取正文首段字体 + 前 20 个样式（保持轻量）
        try:
            clean = _clean_font(doc.Content.Font.Name)
            if clean:
                names.add(clean)
        except Exception:
            pass
        try:
            styles = doc.Styles
            for index in range(1, min(int(styles.Count or 0), 20) + 1):
                try:
                    clean = _clean_font(styles.Item(index).Font.Name)
                except Exception:
                    continue
                if clean:
                    names.add(clean)
        except Exception:
            pass
    if not names:
        return result
    known = installed_font_families()
    result.update(checked=True,
                  missing=sorted(n for n in names if _normalize_font(n) not in known),
                  fonts=sorted(names))
    return result


def font_warning_text(display_name: str, missing, limit: int = 6) -> str:
    """把缺失字体列表转成一条用户提示；无缺失时返回空串（不产生日志）。"""
    names = [str(item).strip() for item in (missing or []) if str(item).strip()]
    if not names:
        return ""
    shown = "、".join(names[:limit])
    if len(names) > limit:
        shown += f" 等 {len(names)} 个"
    prefix = f"{display_name}：" if display_name else ""
    return (f"{prefix}缺少字体 {shown}。已按“缺字体位图化”导出，"
            "若打印预览仍出现方框/乱码，请安装这些字体后重新导入。")


def pdf_font_warning_text(display_name: str, unembedded, limit: int = 6) -> str:
    """PDF 未嵌入字体提示；为空时返回空串。"""
    names = [str(item).strip() for item in (unembedded or []) if str(item).strip()]
    if not names:
        return ""
    shown = "、".join(names[:limit])
    if len(names) > limit:
        shown += f" 等 {len(names)} 个"
    prefix = f"{display_name}：" if display_name else ""
    return (f"{prefix}导出 PDF 中有未嵌入字体（{shown}），"
            "阅读器/打印机替换时可能出现字形差异。")


def _merge_font_info(result: dict, info: Optional[dict]) -> None:
    """把若干次预检结果合并到逐文件结果；全部未能预检时标记 unavailable。"""
    checked = bool(result.get("font_check") == "checked")
    missing = set(result.get("missing_fonts") or [])
    info = info or {}
    if info.get("checked"):
        checked = True
        missing |= {str(item).strip() for item in (info.get("missing") or [])
                    if str(item).strip()}
    if checked:
        result["font_check"] = "checked"
        if missing:
            result["missing_fonts"] = sorted(missing)
    else:
        result["font_check"] = "unavailable"


# ------------------------------------------------------------------ COM 转换
def _dispatch(progids):
    import win32com.client

    last_error: Optional[Exception] = None
    for name in progids:
        try:
            return win32com.client.DispatchEx(name)
        except Exception as exc:  # 逐个 ProgID 尝试，例如 Word.Application.16
            last_error = exc
    raise WordConvertError(f"无法创建 COM 实例 {progids[0]}：{last_error}")


_EXPORT_BASE_KWARGS = dict(ExportFormat=WD_EXPORT_FORMAT_PDF, OpenAfterExport=False)


def _export_pdf(doc, dst: str) -> bool:
    """导出 PDF，优先启用“缺字体位图化”。

    返回是否成功启用了 ``BitmapMissingFonts``；个别 Office 版本不接受该
    参数时捕获异常并无参重试（降级为原有调用），不改变既有导出流程。
    """
    kwargs = dict(OutputFileName=dst, **_EXPORT_BASE_KWARGS)
    try:
        doc.ExportAsFixedFormat(BitmapMissingFonts=True, **kwargs)
        return True
    except Exception:
        _remove_quietly(dst)
        doc.ExportAsFixedFormat(**kwargs)
        return False


def convert_with_word(src: str, dst: str, app=None,
                      font_info: Optional[dict] = None) -> str:
    """用 Word COM 导出 PDF；全程不可见、关闭不保存、结束后退出实例。"""
    created = app is None
    if created:
        app = _dispatch(WORD_PROGIDS)
    doc = None
    try:
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False,
                                 Visible=False)
        if font_info is not None:
            # .doc 没有可读的 OOXML 字体表，只能在 COM 打开后从样式/正文取样；
            # 取不到时保持 checked=False（“未能预检”），不误报。
            if str(src).strip().lower().endswith(".doc"):
                try:
                    font_info.update(_inspect_doc_fonts_com(doc))
                except Exception:
                    pass
        bitmap_ok = _export_pdf(doc, dst)
        if font_info is not None:
            font_info["bitmap_missing_fonts"] = bitmap_ok
    finally:
        if doc is not None:
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if created:
            try:
                app.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
    return "Word"


def convert_with_wps(src: str, dst: str, app=None,
                     font_info: Optional[dict] = None) -> str:
    """用 WPS 文字导出 PDF：优先 ExportAsFixedFormat，退化到 SaveAs2。

    ``font_info`` 仅为与 Word 路径保持统一签名；WPS 无字体位图化参数，
    字体预检也覆盖不到，保持“未能预检”。
    """
    created = app is None
    if created:
        app = _dispatch(WPS_PROGIDS)
    doc = None
    try:
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(OutputFileName=dst,
                                    ExportFormat=WD_EXPORT_FORMAT_PDF,
                                    OpenAfterExport=False)
        except Exception:
            doc.SaveAs2(dst, WD_EXPORT_FORMAT_PDF)
    finally:
        if doc is not None:
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if created:
            try:
                app.Quit()
            except Exception:
                pass
    return "WPS"


def convert_inprocess(src: str, dst: str,
                      font_info: Optional[dict] = None) -> str:
    """在当前进程内尝试转换，返回实际使用的转换器名称。

    * .wps 只走 WPS：装了 Word 也可能无法正确打开 WPS 专有格式；
    * 其余文档按 Word → WPS 优先级尝试。
    * ``font_info`` 用于回传字体预检结果（.doc 在 COM 内取样）。
    """
    import pythoncom

    pythoncom.CoInitialize()
    try:
        errors: List[str] = []
        candidates = (
            (("WPS", convert_with_wps, WPS_PROGIDS),) if is_wps_file(src)
            else (("Word", convert_with_word, WORD_PROGIDS),
                  ("WPS", convert_with_wps, WPS_PROGIDS))
        )
        for converter, func, progids in candidates:
            if not any(progid_registered(p) for p in progids):
                errors.append(f"{converter} 未安装")
                continue
            try:
                func(src, dst, font_info=font_info)
            except Exception as exc:
                errors.append(f"{converter} 转换失败：{exc}")
                continue
            if os.path.isfile(dst) and os.path.getsize(dst) > 0:
                return converter
            errors.append(f"{converter} 未生成 PDF")
        if is_wps_file(src) and not any(
                progid_registered(p) for p in WPS_PROGIDS):
            raise WordConvertError(
                "未检测到 WPS，无法转换 .wps 文件。请安装 WPS，"
                "或用 WPS 打开后另存为 PDF/Word 再导入。")
        raise WordConvertError(
            "Word 文档转换 PDF 失败：" + "；".join(errors)
            + "。请安装 Microsoft Word 或 WPS，或先手动另存为 PDF 再导入。")
    finally:
        pythoncom.CoUninitialize()


# ------------------------------------------------------- 批量常驻实例转换
# 实测（本机 Word 2021）：每文件单独起子进程 + 新 Word 实例约 6~8 秒，
# 其中 Word/COM 冷启动与进程回收占 4 秒以上；一个实例连续转换 4 个文件
# 合计约 2.8 秒。因此多文件时改为“一个子进程 + 一个常驻实例”批量转换，
# 同时保留子进程可被 kill 的超时能力与逐文件失败隔离。
BATCH_MANIFEST_NAME = "manifest.json"
BATCH_RESULTS_NAME = "results.jsonl"


class BatchCancelled(Exception):
    """批量转换被用户中止：调用方停止后续处理并清理未完成的临时文件。"""


def _batch_candidates(kind: str):
    """按 .wps/.doc 选择转换器顺序，与 convert_inprocess 保持一致。"""
    if kind == "wps" or is_wps_file(kind):
        return (("WPS", convert_with_wps, WPS_PROGIDS),)
    return (("Word", convert_with_word, WORD_PROGIDS),
            ("WPS", convert_with_wps, WPS_PROGIDS))


def _get_app(name: str, progids, apps: dict):
    """取（或首次创建）常驻实例；同一子进程内复用，避免重复冷启动。"""
    app = apps.get(name)
    if app is None:
        app = _dispatch(progids)
        try:
            app.Visible = False
            app.DisplayAlerts = 0
        except Exception:
            pass
        apps[name] = app
    return app


def _drop_app(apps: dict, name: str) -> None:
    """退出并移除某个实例；实例异常后下一次会重新创建（失败降级）。"""
    app = apps.pop(name, None)
    if app is None:
        return
    try:
        app.Quit(WD_DO_NOT_SAVE_CHANGES)
    except Exception:
        pass


def _unembedded_font_names(doc) -> List[str]:
    """列出 PDF 中未嵌入的字体（过滤基础 14 字体，避免误报噪音）。"""
    gaps: set = set()
    base14 = {_normalize_font(name) for name in _BASE14_FONTS}
    try:
        for page in doc:
            try:
                fonts = page.get_fonts()
            except Exception:
                continue
            for item in fonts or []:
                if not item or len(item) < 5:
                    continue
                xref, ext, basefont = item[0], str(item[1] or "").lower(), str(item[3] or "")
                embedded = bool(xref) and ext not in ("n/a", "none", "")
                if embedded:
                    continue
                label = basefont.split("+")[-1].strip()
                if not label or _normalize_font(label) in base14:
                    continue
                gaps.add(label)
    except Exception:
        return []
    return sorted(gaps)


def _probe_output(dst: str):
    """校验导出结果，返回 ``(页数, 未嵌入字体列表)``（0 页视为失败）。"""
    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        raise WordConvertError("未生成 PDF 文件")
    import fitz

    with fitz.open(dst) as doc:
        pages = doc.page_count
        gaps = _unembedded_font_names(doc)
    if pages <= 0:
        raise WordConvertError("生成的 PDF 没有页面")
    return pages, gaps


def _run_batch(manifest_path: str, results_path: str) -> int:
    """子进程入口：一个常驻 Word/WPS 实例依次处理清单中的文件。

    每个文件的结果单独追加到 results.jsonl（父进程轮询读取，实现进度可见
    与逐文件失败隔离）；文件失败不会中断后续文件。
    """
    import pythoncom

    with open(manifest_path, "r", encoding="utf-8") as handle:
        items = json.load(handle)

    pythoncom.CoInitialize()
    apps: dict = {}
    all_ok = True
    try:
        with open(results_path, "a", encoding="utf-8") as out:
            for item in items:
                started = time.perf_counter()
                result = {"iid": item.get("iid")}
                errors: List[str] = []
                # 转换前预检（.docx 走 zip 解析；.doc 在 COM 内取样）：
                # 单文件预检失败只影响该文件提示，不中断整批。
                precheck = inspect_document_fonts(item.get("src", ""))
                font_info: dict = {}
                try:
                    for name, func, progids in _batch_candidates(item.get("kind", "word")):
                        if not any(progid_registered(p) for p in progids):
                            errors.append(f"{name} 未安装")
                            continue
                        try:
                            app = _get_app(name, progids, apps)
                            func(item["src"], item["dst"], app=app,
                                 font_info=font_info)
                            pages, gaps = _probe_output(item["dst"])
                        except Exception as exc:
                            errors.append(f"{name} 转换失败：{exc}")
                            _drop_app(apps, name)     # 实例可能已损坏，下次重建
                            _remove_quietly(item["dst"])
                            continue
                        result.update(ok=True, converter=name, pages=pages,
                                      seconds=round(time.perf_counter() - started, 2))
                        if gaps:
                            result["unembedded_fonts"] = gaps
                        break
                    if not result.get("ok"):
                        all_ok = False
                        result.update(ok=False, error="；".join(errors) or "转换失败")
                except Exception as exc:
                    all_ok = False
                    result.update(ok=False, error=f"转换失败：{exc}")
                _merge_font_info(result, precheck)
                _merge_font_info(result, font_info)
                out.write(json.dumps(result, ensure_ascii=False) + "\n")
                out.flush()
    finally:
        for name in list(apps):
            _drop_app(apps, name)
        pythoncom.CoUninitialize()
    return 0 if all_ok else 4


def _retry_remaining(remaining: List[dict], results: dict,
                     timeout_per_file: float, python_exe: Optional[str],
                     stop_event) -> None:
    """批量整体异常/超时后，对未完成文件回退单文件路径各重试一次。

    成功的批量结果不受影响；重试写回同一 dst，先清掉可能的不完整文件，
    不新建临时目录，避免重复转换与临时文件泄漏。
    """
    for item in remaining:
        if stop_event is not None and stop_event.is_set():
            break
        iid = item.get("iid")
        started = time.perf_counter()
        _remove_quietly(item["dst"])
        try:
            info = convert_file(item["src"], item["dst"], timeout=timeout_per_file,
                                python_exe=python_exe)
        except Exception as exc:
            previous = str(results.get(iid, {}).get("error") or "")
            results[iid] = {"iid": iid, "ok": False, "retried": True,
                            "error": (previous + f"；单文件重试仍失败：{exc}").lstrip("；")}
            continue
        entry = {"iid": iid, "ok": True, "retried": True,
                 "converter": info.get("converter"),
                 "pages": info.get("pages"),
                 "seconds": round(time.perf_counter() - started, 2)}
        for key in ("missing_fonts", "font_check", "unembedded_fonts"):
            if info.get(key):
                entry[key] = info[key]
        results[iid] = entry


def convert_batch(items: List[dict], timeout_per_file: float = 180.0,
                  on_result: Optional[Callable[[dict], None]] = None,
                  stop_event=None, python_exe: Optional[str] = None) -> dict:
    """批量转换（单子进程 + 常驻实例），返回 {iid: 结果}。

    * ``items``：``[{"iid","src","dst","kind"}]``，kind 为 word/wps；
    * 逐文件结果通过 ``on_result`` 实时回调，便于界面增量更新；
    * 任一文件超过 ``timeout_per_file`` 秒无进展 → 结束子进程并把未完成项
      标记为超时（不会无限挂起）；未完成项再回退单文件路径各重试一次；
    * 用户中止（``stop_event``）抛 ``BatchCancelled``，临时目录与 Office
      子进程在 finally 中清理，不会残留。
    """
    if not items:
        return {}
    for item in items:
        item["src"] = os.path.abspath(item["src"])
        item["dst"] = os.path.abspath(item["dst"])
        folder = os.path.dirname(item["dst"])
        if folder:
            os.makedirs(folder, exist_ok=True)

    batch_dir = create_temp_dir()
    manifest_path = os.path.join(batch_dir, BATCH_MANIFEST_NAME)
    results_path = os.path.join(batch_dir, BATCH_RESULTS_NAME)
    payload = [{"iid": item.get("iid"), "src": item["src"], "dst": item["dst"],
                "kind": item.get("kind", "word")} for item in items]
    with open(manifest_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False)
    open(results_path, "w", encoding="utf-8").close()

    before = conversion_processes()
    command = [python_exe or sys.executable, os.path.abspath(__file__),
               "--batch", manifest_path, "--results", results_path]
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               creationflags=CREATE_NO_WINDOW)
    results: dict = {}
    handled = 0
    last_activity = time.time()
    timed_out = False
    try:
        while True:
            if stop_event is not None and stop_event.is_set():
                raise BatchCancelled("用户中止了转换。")
            lines = []
            try:
                with open(results_path, "r", encoding="utf-8") as handle:
                    lines = handle.read().splitlines()
            except OSError:
                lines = []
            while handled < len(lines):
                try:
                    one = json.loads(lines[handled])
                except ValueError:
                    handled += 1
                    continue
                handled += 1
                results[one.get("iid")] = one
                last_activity = time.time()
                if on_result is not None:
                    on_result(one)
            if len(results) >= len(items):
                break
            if process.poll() is not None:
                break
            if time.time() - last_activity > timeout_per_file:
                timed_out = True
                break
            time.sleep(0.2)
    finally:
        if process.poll() is None:
            try:
                process.kill()
            except Exception:
                pass
        try:
            out, err = process.communicate(timeout=15)
        except Exception:
            out, err = b"", b""
        _terminate_new_office(before)
        shutil.rmtree(batch_dir, ignore_errors=True)

    remaining = [item for item in items if item.get("iid") not in results]
    if remaining:
        if timed_out:
            message = (f"Word 转换超时（超过 {int(timeout_per_file)} 秒无进展），"
                       "已终止批量转换进程。")
        else:
            detail = _last_line(err) or _last_line(out) or "转换进程异常退出"
            message = f"批量转换失败：{detail}"
        for item in remaining:
            results[item.get("iid")] = {"iid": item.get("iid"), "ok": False,
                                        "error": message}
        # 批量子进程整体异常/超时：已成功的文件保留，未完成项回退单文件路径
        # 各重试一次（同一 dst，不新建临时目录），重试结果统一再回调界面。
        _retry_remaining(remaining, results, timeout_per_file, python_exe, stop_event)
        if on_result is not None:
            for item in remaining:
                outcome = results.get(item.get("iid"))
                if outcome is not None:
                    on_result(outcome)
    return results


# ------------------------------------------------------------- 进程与超时控制
def _list_pids(name: str) -> set:
    """列出指定进程名的 PID（tasklist 输出，失败返回空集合）。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {name}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15, encoding="mbcs",
            errors="ignore", creationflags=CREATE_NO_WINDOW)
    except Exception:
        return set()
    pids = set()
    for line in (result.stdout or "").splitlines():
        parts = [piece.strip('"') for piece in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == name.upper():
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def _kill_pids(pids) -> None:
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True,
                           timeout=10, encoding="mbcs", errors="ignore",
                           creationflags=CREATE_NO_WINDOW)
        except Exception:
            pass


def conversion_processes() -> set:
    pids = set()
    for name in OFFICE_PROCESS_NAMES:
        pids |= _list_pids(name)
    return pids


def _terminate_new_office(before: set) -> None:
    """结束“本次转换新起”的 WINWORD/WPS，避免残留后台进程。

    只结束转换前不存在的 PID：用户原本开着的 Word 不会被误杀。
    """
    current = conversion_processes()
    _kill_pids(current - before)


def convert_file(src: str, dst: str, timeout: float = 180.0,
                 python_exe: Optional[str] = None) -> dict:
    """子进程执行转换：返回 {"converter": "Word"/"WPS", "pages": 页数}。

    超时会终止子进程，并清理本次新起的 Office 进程与不完整的输出文件。
    """
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise WordConvertError(f"找不到 Word 文件：{src}")
    folder = os.path.dirname(dst)
    if folder:
        os.makedirs(folder, exist_ok=True)

    before = conversion_processes()
    command = [python_exe or sys.executable, os.path.abspath(__file__),
               "--src", src, "--dst", dst]
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               creationflags=CREATE_NO_WINDOW)
    try:
        out, err = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate()
        _terminate_new_office(before)
        _remove_quietly(dst)
        raise WordConvertError(
            f"Word 转换超时（超过 {int(timeout)} 秒），已终止转换进程。")
    _terminate_new_office(before)

    if process.returncode != 0:
        _remove_quietly(dst)
        detail = _last_line(err) or _last_line(out) or f"退出码 {process.returncode}"
        raise WordConvertError(f"Word 转换失败：{detail}")
    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        raise WordConvertError("Word 转换失败：未生成 PDF 文件。")

    converter = ""
    font_summary: dict = {"checked": False, "missing": []}
    for line in (out or "").decode("utf-8", "ignore").splitlines():
        if line.startswith("converter="):
            converter = line.split("=", 1)[1].strip()
        elif line.startswith("font_info="):
            try:
                parsed = json.loads(line.split("=", 1)[1])
            except ValueError:
                parsed = {}
            if isinstance(parsed, dict):
                font_summary = parsed
    pages = 0
    gaps: List[str] = []
    try:
        import fitz

        with fitz.open(dst) as doc:
            pages = doc.page_count
            gaps = _unembedded_font_names(doc)
    except Exception as exc:
        _remove_quietly(dst)
        raise WordConvertError(f"转换结果不是有效 PDF：{exc}")
    if pages <= 0:
        _remove_quietly(dst)
        raise WordConvertError("Word 转换失败：生成的 PDF 没有页面。")
    info = {"converter": converter or "Word", "pages": pages,
            "font_check": "checked" if font_summary.get("checked") else "unavailable"}
    missing = [str(item).strip() for item in (font_summary.get("missing") or [])
               if str(item).strip()]
    if missing:
        info["missing_fonts"] = sorted(set(missing))
    if gaps:
        info["unembedded_fonts"] = gaps
    return info


def _last_line(blob) -> str:
    text = (blob or b"").decode("utf-8", "ignore").strip()
    return text.splitlines()[-1] if text else ""


def _remove_quietly(path: str) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


def _main(argv=None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Word/WPS 转 PDF（子进程入口）")
    parser.add_argument("--src")
    parser.add_argument("--dst")
    parser.add_argument("--batch", metavar="MANIFEST",
                        help="批量清单 JSON 路径（与 --results 配合）")
    parser.add_argument("--results", help="批量结果 JSONL 输出路径")
    args = parser.parse_args(argv)
    if args.batch:
        if not args.results:
            print("批量模式需要 --results 参数", file=sys.stderr)
            return 2
        try:
            return _run_batch(args.batch, args.results)
        except Exception as exc:  # 批量入口异常也要以可读信息结束
            print(f"批量转换异常：{exc}", file=sys.stderr)
            return 3
    if not args.src or not args.dst:
        print("单文件模式需要 --src 与 --dst 参数", file=sys.stderr)
        return 2
    try:
        font_info: dict = {}
        precheck = inspect_document_fonts(args.src)
        converter = convert_inprocess(args.src, args.dst, font_info=font_info)
    except WordConvertError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    except Exception as exc:  # 未预期异常也转成可读错误
        print(f"Word 转换异常：{exc}", file=sys.stderr)
        return 3
    merged: dict = {}
    _merge_font_info(merged, precheck)
    _merge_font_info(merged, font_info)
    summary = {"checked": merged.get("font_check") == "checked",
               "missing": merged.get("missing_fonts") or []}
    print(f"converter={converter}")
    print("font_info=" + json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(_main())
