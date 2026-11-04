# -*- coding: utf-8 -*-
"""Windows 打印后端：GDI 栅格打印、DEVMODE 设置、打印机状态与队列监听。

说明：
  * 每一页都用 PyMuPDF 栅格化后经 GDI 送入打印队列，页序完全由程序控制；
  * 打印前通过 DEVMODE 强制单面（dmDuplex=SIMPLEX）、单份、一页一版（dmNup），
    避免打印机驱动自带双面/多版功能破坏手动双面流程；
  * 驱动不支持的 DEVMODE 选项会被驱动忽略或回退，界面会给出降级提示。
"""

import ctypes
import logging
import math
import os
import re
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

_logger = logging.getLogger("pdf_duplex")

try:
    import win32print  # pywin32
except ImportError:  # pragma: no cover
    win32print = None

try:
    from PIL import Image  # Pillow：灰度/黑白/省墨后处理
except ImportError:  # pragma: no cover
    Image = None

from core_utils import PrintError, PrinterOfflineError, PrintTimeoutError, CancelledByUser

JOB_POLL_INTERVAL = 1.0              # 打印队列轮询间隔（秒）
JOB_TIMEOUT_SECONDS = 1800.0         # 单个打印任务等待上限（秒）
POST_JOB_SETTLE_SECONDS = 2.0        # 任务完成后的额外等待，给机械出纸留余量

# 稳定性参数（对标 PDFtoPrinter / winprint 的提交重试做法）：
#   * 仅在“尚未成功出纸一面”时重试整个任务提交，出过纸的任务绝不重试，
#     否则会出现重复页；
#   * 队列 OpenPrinter 偶发失败（打印后台服务/安全软件瞬时占用）时短暂退避后
#     重试，避免整批判为失败。
SUBMIT_RETRY_ATTEMPTS = 1            # 提交阶段额外重试次数（总计 2 次尝试）
OPEN_PRINTER_RETRY_ATTEMPTS = 3      # 队列句柄打开失败的重试次数
RETRY_BACKOFF_SECONDS = 0.5          # 退避基数（线性递增）

# GDI / Win32 常量
HORZRES = 8
VERTRES = 10
MM_TEXT = 1
STRETCH_HALFTONE = 4
SRCCOPY = 0x00CC0020
BI_RGB = 0
MAX_RENDER_PIXELS = 60_000_000      # 单页位图像素数上限，防止超大页面耗尽内存

# DEVMODE 字段/取值常量
DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DM_PAPERLENGTH = 0x00000004
DM_PAPERWIDTH = 0x00000008
DM_NUP = 0x00000040
DM_COPIES = 0x00000100
DM_DEFAULTSOURCE = 0x00000200
DM_PRINTQUALITY = 0x00000400
DM_COLOR = 0x00000800
DM_DUPLEX = 0x00001000
DM_IN_BUFFER = 0x00000008
DM_OUT_BUFFER = 0x00000002

DMDUP_SIMPLEX = 1
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2
DMCOLOR_COLOR = 1
DMCOLOR_MONOCHROME = 2
DMNUP_ONEUP = 2
DMPAPER_USER = 256

# 打印质量（负值为 GDI 标准分辨率枚举）
DMRES_DRAFT = -1
DMRES_MEDIUM = -3
DMRES_HIGH = -4

# DeviceCapabilities 能力编号
DC_BINS = 6
DC_BINNAMES = 12
DC_DUPLEX = 7
DC_COLLATE = 22

# 打印机状态位（用 getattr 兼容不同 pywin32 版本）
_STATUS_BITS: Tuple[Tuple[str, str], ...] = (
    ("PRINTER_STATUS_PAPER_JAM", "卡纸"),
    ("PRINTER_STATUS_PAPER_OUT", "缺纸"),
    ("PRINTER_STATUS_PAPER_PROBLEM", "纸张问题"),
    ("PRINTER_STATUS_OFFLINE", "脱机"),
    ("PRINTER_STATUS_PAUSED", "已暂停"),
    ("PRINTER_STATUS_ERROR", "错误"),
    ("PRINTER_STATUS_NO_TONER", "缺墨粉"),
    ("PRINTER_STATUS_TONER_LOW", "墨粉不足"),
    ("PRINTER_STATUS_DOOR_OPEN", "门未关"),
    ("PRINTER_STATUS_USER_INTERVENTION", "需要人工处理"),
    ("PRINTER_STATUS_OUTPUT_BIN_FULL", "出纸仓已满"),
    ("PRINTER_STATUS_PRINTING", "正在打印"),
    ("PRINTER_STATUS_PROCESSING", "正在处理"),
    ("PRINTER_STATUS_BUSY", "忙"),
    ("PRINTER_STATUS_WAITING", "等待中"),
    ("PRINTER_STATUS_INITIALIZING", "正在初始化"),
    ("PRINTER_STATUS_WARMING_UP", "预热中"),
    ("PRINTER_STATUS_POWER_SAVE", "省电模式"),
)


# ----------------------------------------------------------------------------
# Win32 结构与 API 绑定
# ----------------------------------------------------------------------------
class _DOCINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_int),
        ("lpszDocName", wintypes.LPCWSTR),
        ("lpszOutput", wintypes.LPCWSTR),
        ("lpszDatatype", wintypes.LPCWSTR),
        ("fwType", wintypes.DWORD),
    ]


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class _DEVMODEW(ctypes.Structure):
    """DEVMODEW 的打印相关字段。"""
    _fields_ = [
        ("dmDeviceName", ctypes.c_wchar * 32),
        ("dmSpecVersion", ctypes.c_ushort),
        ("dmDriverVersion", ctypes.c_ushort),
        ("dmSize", ctypes.c_ushort),
        ("dmDriverExtra", ctypes.c_ushort),
        ("dmFields", ctypes.c_uint32),
        ("dmOrientation", ctypes.c_short),
        ("dmPaperSize", ctypes.c_short),
        ("dmPaperLength", ctypes.c_short),
        ("dmPaperWidth", ctypes.c_short),
        ("dmScale", ctypes.c_short),
        ("dmCopies", ctypes.c_short),
        ("dmDefaultSource", ctypes.c_short),
        ("dmPrintQuality", ctypes.c_short),
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", ctypes.c_wchar * 32),
        ("dmLogPixels", ctypes.c_ushort),
        ("dmBitsPerPel", ctypes.c_uint32),
        ("dmPelsWidth", ctypes.c_uint32),
        ("dmPelsHeight", ctypes.c_uint32),
        # dmDisplayFlags 与 dmNup 共用同一位置
        ("dmDisplayFlags", ctypes.c_uint32),
        ("dmDisplayFrequency", ctypes.c_uint32),
        ("dmICMMethod", ctypes.c_uint32),
        ("dmICMIntent", ctypes.c_uint32),
        ("dmMediaType", ctypes.c_uint32),
        ("dmDitherType", ctypes.c_uint32),
        ("dmReserved1", ctypes.c_uint32),
        ("dmReserved2", ctypes.c_uint32),
        ("dmPanningWidth", ctypes.c_uint32),
        ("dmPanningHeight", ctypes.c_uint32),
    ]


_spool = ctypes.WinDLL("winspool.drv", use_last_error=True)
_spool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE), ctypes.c_void_p]
_spool.OpenPrinterW.restype = wintypes.BOOL
_spool.ClosePrinter.argtypes = [wintypes.HANDLE]
_spool.ClosePrinter.restype = wintypes.BOOL
_spool.DocumentPropertiesW.argtypes = [
    wintypes.HWND, wintypes.HANDLE, wintypes.LPWSTR,
    ctypes.POINTER(_DEVMODEW), ctypes.POINTER(_DEVMODEW), wintypes.DWORD,
]
_spool.DocumentPropertiesW.restype = ctypes.c_long

_gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
_gdi32.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
_gdi32.CreateDCW.restype = wintypes.HDC
_gdi32.DeleteDC.argtypes = [wintypes.HDC]
_gdi32.DeleteDC.restype = wintypes.BOOL
_gdi32.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
_gdi32.GetDeviceCaps.restype = ctypes.c_int
_gdi32.SetMapMode.argtypes = [wintypes.HDC, ctypes.c_int]
_gdi32.SetMapMode.restype = ctypes.c_int
_gdi32.SetStretchBltMode.argtypes = [wintypes.HDC, ctypes.c_int]
_gdi32.SetStretchBltMode.restype = ctypes.c_int
_gdi32.SetBrushOrgEx.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_void_p]
_gdi32.SetBrushOrgEx.restype = wintypes.BOOL
_gdi32.StartDocW.argtypes = [wintypes.HDC, ctypes.POINTER(_DOCINFOW)]
_gdi32.StartDocW.restype = ctypes.c_int
_gdi32.StartPage.argtypes = [wintypes.HDC]
_gdi32.StartPage.restype = ctypes.c_int
_gdi32.EndPage.argtypes = [wintypes.HDC]
_gdi32.EndPage.restype = ctypes.c_int
_gdi32.EndDoc.argtypes = [wintypes.HDC]
_gdi32.EndDoc.restype = ctypes.c_int
_gdi32.AbortDoc.argtypes = [wintypes.HDC]
_gdi32.AbortDoc.restype = ctypes.c_int
_gdi32.StretchDIBits.argtypes = [
    wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_void_p, ctypes.POINTER(_BITMAPINFOHEADER), wintypes.DWORD, wintypes.DWORD,
]
_gdi32.StretchDIBits.restype = ctypes.c_int


# ----------------------------------------------------------------------------
# 基础工具
# ----------------------------------------------------------------------------
def _win_error_text(code: Optional[int] = None) -> str:
    """把 Win32 错误码翻译为可读文本。"""
    if code is None:
        code = ctypes.get_last_error()
    try:
        text = ctypes.FormatError(code)
    except Exception:
        text = ""
    text = (text or "").strip()
    return text or f"Windows 错误码 {code}"


def _job_status_flags(job: dict) -> int:
    """兼容 pywin32 不同版本：把打印任务状态统一成位标志。"""
    if win32print is None:
        return 0
    status = job.get("Status")
    if isinstance(status, int):
        return status
    text = f"{status or ''} {job.get('pStatus') or ''}".lower()
    # 单词按词匹配，避免厂商自定义状态文本的子串误判（如 "...error..." 类提示语）；
    # 双词短语仍按子串匹配（分词后无法直接比较）。
    words = set(re.findall(r"[a-z]+", text))
    flags = 0
    mapping = (
        ("printing", win32print.JOB_STATUS_PRINTING),
        ("printed", win32print.JOB_STATUS_PRINTED),
        ("spooling", win32print.JOB_STATUS_SPOOLING),
        ("offline", win32print.JOB_STATUS_OFFLINE),
        ("error", win32print.JOB_STATUS_ERROR),
        ("paper out", win32print.JOB_STATUS_PAPEROUT),
        ("paperout", win32print.JOB_STATUS_PAPEROUT),
        ("deleted", win32print.JOB_STATUS_DELETED),
        ("complete", win32print.JOB_STATUS_COMPLETE),
        ("blocked", win32print.JOB_STATUS_BLOCKED_DEVQ),
    )
    for key, value in mapping:
        if (" " in key and key in text) or (" " not in key and key in words):
            flags |= value
    return flags


def _fit_rect(src_w: int, src_h: int, dst_w: int, dst_h: int) -> Tuple[int, int, int, int]:
    """等比缩放并居中，保证页面内容完整、不被裁切。"""
    scale = min(dst_w / src_w, dst_h / src_h)
    width = max(1, int(src_w * scale))
    height = max(1, int(src_h * scale))
    x = max(0, (dst_w - width) // 2)
    y = max(0, (dst_h - height) // 2)
    return x, y, width, height


def _prepare_dib(pixmap):
    """把 PyMuPDF 的紧密 RGB 采样补齐为 GDI 要求的 4 字节对齐扫描行。

    BITMAPINFOHEADER 规定：未压缩 DIB 的每一行按 DWORD（4 字节）对齐。
    返回 (位图数据指针, 对齐后的每行字节数, 需要保持引用的对象)。
    """
    width, height = pixmap.width, pixmap.height
    src_stride = pixmap.stride
    dst_stride = (width * 3 + 3) & ~3
    samples = pixmap.samples
    samples_ref = ctypes.c_char_p(samples)

    if src_stride == dst_stride:
        return ctypes.cast(samples_ref, ctypes.c_void_p), dst_stride, samples_ref

    buffer = ctypes.create_string_buffer(dst_stride * height)
    src_addr = ctypes.cast(samples_ref, ctypes.c_void_p).value
    for row in range(height):
        ctypes.memmove(
            ctypes.addressof(buffer) + row * dst_stride,
            src_addr + row * src_stride,
            src_stride,
        )
    return ctypes.cast(buffer, ctypes.c_void_p), dst_stride, buffer


def _render_page(page, dpi: int, annots: bool = True):
    """把 PDF 页面渲染成 RGB 位图（DPI 可配置，过大时自动降采样保护内存）。

    annots=False 时不输出批注/表单/戳记（由 PyMuPDF 渲染层控制）。
    """
    import fitz
    scale = max(int(dpi), 72) / 72.0
    width_px = page.rect.width * scale
    height_px = page.rect.height * scale
    if width_px * height_px > MAX_RENDER_PIXELS:
        scale *= math.sqrt(MAX_RENDER_PIXELS / (width_px * height_px))
        # 用户可能拿到比预期模糊的输出，必须在日志里留下原因
        _logger.warning(
            "页面位图超过 %d 万像素上限，已自动降采样（等效约 %.0f DPI），输出清晰度会下降",
            MAX_RENDER_PIXELS // 1_000_000, scale * 72)
    matrix = fitz.Matrix(scale, scale)
    return page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False, annots=annots)


def _rasterize(page, dpi: int, color_mode: str = "color", ink_saver: bool = False, annots: bool = True):
    """渲染页面并按图像模式做后处理，返回带 width/height/stride/samples 的对象。

    color / gray / mono 以及省墨都在栅格层完成，不依赖打印机驱动是否支持，
    因此任何打印机都能得到一致效果（这是对驱动能力不足的降级实现）。
    """
    import fitz
    from types import SimpleNamespace

    pixmap = _render_page(page, dpi, annots=annots)
    width, height = pixmap.width, pixmap.height
    samples = pixmap.samples

    if (color_mode != "color" or ink_saver) and Image is not None:
        image = Image.frombytes("RGB", (width, height), samples)
        if color_mode == "gray":
            image = image.convert("L").convert("RGB")
        elif color_mode == "mono":
            # 黑白：按亮度阈值二值化，适合纯文字文档
            image = image.convert("L").point(lambda v: 0 if v < 128 else 255).convert("RGB")
        if ink_saver:
            # 省墨：整体淡化深色像素，减少墨粉消耗（降级实现，不是驱动力度控制）
            image = image.point(lambda v: min(255, int(v * 0.72 + 55)))
        samples = image.tobytes()

    return SimpleNamespace(width=width, height=height, stride=width * 3, samples=samples)


# ----------------------------------------------------------------------------
# 打印机枚举 / 状态
# ----------------------------------------------------------------------------
def _printer_port(printer_name: str) -> str:
    """读取打印机端口名（DeviceCapabilities 需要）。"""
    if win32print is None:
        return ""
    handle = None
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2) or {}
        return info.get("pPortName") or ""
    except Exception:
        return ""
    finally:
        if handle:
            try:
                win32print.ClosePrinter(handle)
            except Exception:
                pass


def list_printers() -> List[str]:
    """枚举本机与已连接的打印机。"""
    if win32print is None:
        raise PrintError("缺少 pywin32 库，请执行：pip install pywin32")
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 2
    )
    return [item.get("pPrinterName") or "" for item in printers if item.get("pPrinterName")]


def get_default_printer() -> str:
    """获取系统默认打印机名称（失败返回空串）。"""
    if win32print is None:
        return ""
    try:
        return win32print.GetDefaultPrinter() or ""
    except Exception:
        return ""


def get_printer_sources(printer_name: str) -> List[Tuple[str, int]]:
    """获取纸张来源/纸盒列表 [(名称, DEVMODE 代码)]。

    通过 DeviceCapabilities(DC_BINNAMES + DC_BINS) 让驱动自己报告名称与代码，
    避免硬编码 DMBIN_* 常量与厂商纸盒不一致。失败时退回“自动选择”。
    """
    result: List[Tuple[str, int]] = [("自动选择", 7)]  # DMBIN_AUTO
    if win32print is None:
        return result
    port = _printer_port(printer_name)
    if not port:
        return result
    try:
        names = win32print.DeviceCapabilities(printer_name, port, DC_BINNAMES)
        codes = win32print.DeviceCapabilities(printer_name, port, DC_BINS)
    except Exception:
        return result
    if not isinstance(names, tuple) or not isinstance(codes, tuple):
        return result
    for name, code in zip(names, codes):
        name = str(name).strip()
        if not name:
            continue
        if any(name == existing for existing, _ in result):
            continue
        result.append((name, int(code)))
    return result


def get_printable_insets_pt(printer_name: str) -> Optional[Tuple[float, float, float, float]]:
    """驱动报告的“可打印区域”相对纸张四边的内缩，单位点。

    通过 CreateDC + GetDeviceCaps 读取物理纸张与可打印像素：
      PHYSICALWIDTH/HEIGHT 是整张纸，HORZRES/VERTRES 是可打印像素，
      PHYSICALOFFSETX/Y 是可打印区左上角相对纸张左上角的偏移。
    驱动不提供或被系统拒绝时返回 None，界面按纸张边缘画虚线并提示。
    """
    if not printer_name:
        return None
    hdc = None
    try:
        hdc = _gdi32.CreateDCW("WINSPOOL", printer_name, None, None)
        if not hdc:
            return None
        horz = int(_gdi32.GetDeviceCaps(hdc, 8))              # HORZRES
        vert = int(_gdi32.GetDeviceCaps(hdc, 10))             # VERTRES
        dpi_x = int(_gdi32.GetDeviceCaps(hdc, 88)) or 300     # LOGPIXELSX
        dpi_y = int(_gdi32.GetDeviceCaps(hdc, 90)) or 300     # LOGPIXELSY
        phys_w = int(_gdi32.GetDeviceCaps(hdc, 110))          # PHYSICALWIDTH
        phys_h = int(_gdi32.GetDeviceCaps(hdc, 111))          # PHYSICALHEIGHT
        off_x = int(_gdi32.GetDeviceCaps(hdc, 112))           # PHYSICALOFFSETX
        off_y = int(_gdi32.GetDeviceCaps(hdc, 113))           # PHYSICALOFFSETY
        if horz <= 0 or vert <= 0 or phys_w <= 0 or phys_h <= 0:
            return None
        scale_x, scale_y = 72.0 / dpi_x, 72.0 / dpi_y
        left = max(0.0, off_x * scale_x)
        top = max(0.0, off_y * scale_y)
        right = max(0.0, (phys_w - off_x - horz) * scale_x)
        bottom = max(0.0, (phys_h - off_y - vert) * scale_y)
        return (left, top, right, bottom)
    except Exception:
        return None
    finally:
        if hdc:
            try:
                _gdi32.DeleteDC(hdc)
            except Exception:
                pass


def get_printer_status(printer_name: str) -> Dict[str, object]:
    """读取打印机状态，返回中文状态文本与关键标志。"""
    data: Dict[str, object] = {
        "text": "未知", "flags": 0, "offline": False, "jam": False,
        "paper_out": False, "paused": False, "printing": False, "jobs": 0,
        "default": False,
    }
    if win32print is None:
        data["text"] = "缺少 pywin32"
        return data
    handle = None
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2) or {}
    except Exception as exc:
        data["text"] = f"无法连接（{exc}）"
        data["offline"] = True
        return data
    finally:
        if handle:
            try:
                win32print.ClosePrinter(handle)
            except Exception:
                pass

    status = int(info.get("Status") or 0)
    attributes = int(info.get("Attributes") or 0)
    jobs = int(info.get("cJobs") or 0)
    data["flags"] = status
    data["jobs"] = jobs

    if attributes & getattr(win32print, "PRINTER_ATTRIBUTE_WORK_OFFLINE", 0x80):
        data["offline"] = True
    if status & getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x80):
        data["offline"] = True
    if status & getattr(win32print, "PRINTER_STATUS_PAPER_JAM", 0x8):
        data["jam"] = True
    if status & getattr(win32print, "PRINTER_STATUS_PAPER_OUT", 0x10):
        data["paper_out"] = True
    if status & getattr(win32print, "PRINTER_STATUS_PAUSED", 0x1):
        data["paused"] = True
    if status & getattr(win32print, "PRINTER_STATUS_PRINTING", 0x400):
        data["printing"] = True

    labels = []
    for flag_name, label in _STATUS_BITS:
        value = getattr(win32print, flag_name, 0)
        if value and status & value:
            labels.append(label)
    if data["offline"]:
        text = "脱机"
    elif data["jam"]:
        text = "卡纸"
    elif data["paper_out"]:
        text = "缺纸"
    elif data["paused"]:
        text = "已暂停"
    elif status & getattr(win32print, "PRINTER_STATUS_ERROR", 0x2):
        text = "错误"
    elif jobs > 0:
        text = f"忙碌（队列 {jobs} 个任务）"
    elif labels:
        text = "、".join(dict.fromkeys(labels))
    else:
        text = "就绪"
    data["text"] = text
    data["default"] = bool(printer_name == get_default_printer())
    return data


def check_printer_ready(printer_name: str) -> None:
    """打印前的轻量检查：能否打开打印机、是否处于脱机/缺纸状态。"""
    if win32print is None:
        raise PrintError("缺少 pywin32 库，请执行：pip install pywin32")
    status = get_printer_status(printer_name)
    if status["offline"]:
        raise PrinterOfflineError(
            f"打印机「{printer_name}」当前离线，请检查电源与 USB/网络连接后重试。"
        )
    if status["jam"]:
        raise PrintError(f"打印机「{printer_name}」报告卡纸，请清理后重试。")
    if status["paper_out"]:
        raise PrintError(f"打印机「{printer_name}」报告缺纸，请补充纸张后重试。")


def _try_cancel_job(printer_name: str, job_id: int) -> None:
    """尽力取消卡住的任务，失败不影响主流程。"""
    if win32print is None or job_id <= 0:
        return
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_CANCEL)
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        pass


def set_job_paused(printer_name: str, job_id: int, paused: bool) -> bool:
    """暂停/继续当前打印任务（队列控制）。失败返回 False，由调用方降级处理。"""
    if win32print is None or job_id <= 0:
        return False
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            control = win32print.JOB_CONTROL_PAUSE if paused else win32print.JOB_CONTROL_RESUME
            win32print.SetJob(handle, job_id, 0, None, control)
        finally:
            win32print.ClosePrinter(handle)
        return True
    except Exception:
        return False


def wait_for_print_job(
    printer_name: str,
    job_id: int,
    stop_event: threading.Event,
    log: Optional[Callable[[str], None]] = None,
    pause_event: Optional[threading.Event] = None,
) -> bool:
    """监听打印队列，直到本阶段任务完成。

    完成判定（任一满足）：任务已从队列消失 / PRINTED / COMPLETE /
    PagesPrinted >= TotalPages。
    脱机、缺纸、任务错误时抛出异常；用户中止返回 False。
    队列监听超时由调用方降级为固定延时。
    """
    if win32print is None:
        raise PrintError("缺少 pywin32 库，请执行：pip install pywin32")

    deadline = time.time() + JOB_TIMEOUT_SECONDS
    last_log = 0.0
    paused_hint_logged = False
    open_failures = 0
    get_job_failures = 0
    while True:
        if stop_event.is_set():
            return False
        if pause_event is not None and pause_event.is_set():
            # 用户暂停：不计入超时，也不重复打日志
            deadline += JOB_POLL_INTERVAL
            time.sleep(JOB_POLL_INTERVAL)
            continue
        if time.time() > deadline:
            raise PrintTimeoutError(
                "监听打印队列超时（任务可能卡在队列中）。"
                "将改用固定延时兜底继续流程；若经常超时，请检查打印机是否缺纸、卡纸，"
                "或直接在界面上改用「固定延时」模式。"
            )

        try:
            handle = win32print.OpenPrinter(printer_name)
        except Exception as exc:
            # 打印后台服务/安全软件会瞬时占用队列接口，短暂退避后重试，
            # 避免偶发 OpenPrinter 失败把整批判为“无法连接队列”。
            open_failures += 1
            if open_failures < OPEN_PRINTER_RETRY_ATTEMPTS and not stop_event.is_set():
                time.sleep(RETRY_BACKOFF_SECONDS * open_failures)
                continue
            raise PrintError(f"无法连接打印队列：{printer_name}\n{exc}") from exc
        open_failures = 0

        try:
            printer_info = win32print.GetPrinter(handle, 2) or {}
            attributes = int(printer_info.get("Attributes") or 0)
            printer_status = int(printer_info.get("Status") or 0)
            if attributes & getattr(win32print, "PRINTER_ATTRIBUTE_WORK_OFFLINE", 0x80) or \
                    printer_status & getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x80):
                _try_cancel_job(printer_name, job_id)
                raise PrinterOfflineError(
                    f"打印机「{printer_name}」已脱机，打印任务已取消。\n请检查电源与连接后重新打印。"
                )
            if (printer_status & getattr(win32print, "PRINTER_STATUS_PAUSED", 0x1)
                    and not paused_hint_logged and log):
                # 不改变流程，只解释“为什么一直没动静”，避免用户干等到监听超时
                paused_hint_logged = True
                log("打印机当前处于暂停状态：任务会一直停在队列中，"
                    "请在 Windows 队列里恢复后继续。")

            try:
                job = win32print.GetJob(handle, job_id, 2)
            except Exception:
                # 任务消失通常意味着已出纸完成；但瞬时的 RPC/权限故障（安全软件
                # 拦截等）同样会让 GetJob 抛异常。连续多次失败才判定“任务已出队列”，
                # 避免误判完成而提前提示翻面，导致整批正反错位。
                get_job_failures += 1
                if get_job_failures >= 3:
                    _logger.warning("GetJob 连续 %d 次失败，判定任务已出队列（视为完成）",
                                    get_job_failures)
                    return True
                time.sleep(JOB_POLL_INTERVAL)
                continue
            get_job_failures = 0

            flags = _job_status_flags(job)
            if flags & getattr(win32print, "JOB_STATUS_OFFLINE", 0):
                raise PrinterOfflineError(f"打印机「{printer_name}」脱机，请检查连接后重试。")
            if flags & (getattr(win32print, "JOB_STATUS_ERROR", 0)
                        | getattr(win32print, "JOB_STATUS_PAPEROUT", 0)):
                _try_cancel_job(printer_name, job_id)
                raise PrintError("打印任务出错（可能缺纸、卡纸或打印机报错），请检查设备后重试。")
            if flags & (getattr(win32print, "JOB_STATUS_PRINTED", 0x80)
                        | getattr(win32print, "JOB_STATUS_COMPLETE", 0x1000)):
                return True

            try:
                printed = int(job.get("PagesPrinted") or 0)
                total = int(job.get("TotalPages") or 0)
            except (TypeError, ValueError):
                printed, total = 0, 0
            if total > 0 and printed >= total:
                return True

            now = time.time()
            if log and now - last_log >= 5.0:
                last_log = now
                log(f"正在打印：已出纸 {printed}/{total} 页…" if total > 0 else "正在打印…")
        finally:
            win32print.ClosePrinter(handle)

        time.sleep(JOB_POLL_INTERVAL)


def open_directory(path: str) -> None:
    """用资源管理器打开目录（不存在则先创建）。"""
    try:
        os.makedirs(path, exist_ok=True)
        os.startfile(path)  # noqa: S606 - 仅 Windows 桌面工具
    except Exception as exc:
        raise PrintError(f"无法打开输出目录：{path}\n{exc}") from exc


def cancel_job(printer_name: str, job_id: int) -> None:
    """公开的“取消队列任务”入口（用于界面取消按钮）。"""
    _try_cancel_job(printer_name, job_id)


# ----------------------------------------------------------------------------
# DEVMODE 设置
# ----------------------------------------------------------------------------
@dataclass
class PrinterOptions:
    """传递给 DEVMODE 的打印选项。"""
    landscape: bool = False
    paper_name: str = "A4"
    paper_width_mm: float = 210.0
    paper_height_mm: float = 297.0
    source_code: int = 7
    color_mode: str = "color"      # color / gray / mono
    print_quality: str = "normal"  # draft / normal / high
    dpi: int = 300
    ink_saver: bool = False
    annots: bool = True


def devmode_field_updates(options: PrinterOptions) -> Dict[str, object]:
    """把打印选项转换为 DEVMODE 字段更新（纯函数，便于单测）。

    ★ 强制单面、单份、一页一版：
      - dmDuplex = 1（单面），保证手工翻面流程不被驱动自动双面破坏；
      - dmCopies = 1（份数由程序按页序控制，避免驱动重复出纸）；
      - dmDisplayFlags/dmNup = 2（DMNUP_ONEUP），排版由程序完成。
    """
    from pdf_layout import PAPER_CODES, CUSTOM_PAPER_NAME

    typical = options.paper_name in PAPER_CODES
    paper_code = PAPER_CODES.get(options.paper_name, DMPAPER_USER)
    updates: Dict[str, object] = {
        "dmOrientation": DMORIENT_LANDSCAPE if options.landscape else DMORIENT_PORTRAIT,
        "dmDuplex": DMDUP_SIMPLEX,
        "dmCopies": 1,
        "dmDefaultSource": int(options.source_code),
        "dmDisplayFlags": DMNUP_ONEUP,   # 与 dmNup 共用，禁止驱动再次多版排版
    }
    fields = (DM_ORIENTATION | DM_DUPLEX | DM_COPIES | DM_NUP | DM_DEFAULTSOURCE)

    updates["dmPaperSize"] = int(paper_code)
    fields |= DM_PAPERSIZE
    if options.paper_name == CUSTOM_PAPER_NAME or not typical:
        # 自定义纸张：宽高单位 0.1 毫米
        updates["dmPaperWidth"] = max(1, int(round(float(options.paper_width_mm) * 10)))
        updates["dmPaperLength"] = max(1, int(round(float(options.paper_height_mm) * 10)))
        fields |= DM_PAPERLENGTH | DM_PAPERWIDTH

    mode = "mono" if options.color_mode == "mono" else "color"
    updates["dmColor"] = DMCOLOR_MONOCHROME if mode == "mono" else DMCOLOR_COLOR
    fields |= DM_COLOR

    if options.print_quality == "draft":
        updates["dmPrintQuality"] = DMRES_DRAFT
        fields |= DM_PRINTQUALITY
    elif options.print_quality == "high":
        updates["dmPrintQuality"] = DMRES_HIGH
        fields |= DM_PRINTQUALITY
    # normal：不主动覆盖驱动默认质量

    updates["dmFields"] = fields
    return updates


def _create_printer_dc(printer_name: str, options: PrinterOptions):
    """创建打印机设备上下文（HDC），并通过 DEVMODE 应用打印选项。

    DEVMODE 获取/校验失败时退回驱动默认设置，不影响基本打印。
    返回 (HDC, 需要保持引用的 DEVMODE 缓冲区)。
    """
    devmode_buffer = None
    devmode_ptr = None
    printer_handle = wintypes.HANDLE()
    opened = False

    try:
        if _spool.OpenPrinterW(printer_name, ctypes.byref(printer_handle), None):
            opened = True
            size = _spool.DocumentPropertiesW(None, printer_handle, printer_name, None, None, 0)
            if size > 0:
                devmode_buffer = ctypes.create_string_buffer(size)
                devmode_ptr = ctypes.cast(devmode_buffer, ctypes.POINTER(_DEVMODEW))
                if _spool.DocumentPropertiesW(
                    None, printer_handle, printer_name, devmode_ptr, None, DM_OUT_BUFFER
                ) > 0:
                    devmode = devmode_ptr.contents
                    updates = devmode_field_updates(options)
                    for name, value in updates.items():
                        if name == "dmFields":
                            continue
                        setattr(devmode, name, value)
                    devmode.dmFields |= int(updates["dmFields"])
                    # 让驱动校验一次修改后的 DEVMODE（失败也不影响后续创建 DC）
                    _spool.DocumentPropertiesW(
                        None, printer_handle, printer_name,
                        devmode_ptr, devmode_ptr, DM_IN_BUFFER | DM_OUT_BUFFER,
                    )
                else:
                    devmode_ptr = None
                    devmode_buffer = None
    except Exception:
        devmode_ptr = None
        devmode_buffer = None
    finally:
        if opened:
            _spool.ClosePrinter(printer_handle)

    hdc = _gdi32.CreateDCW("WINSPOOL", printer_name, None, devmode_ptr)
    if not hdc:
        raise PrinterOfflineError(
            f"无法创建打印机上下文（打印机可能离线或驱动异常）：{printer_name}\n{_win_error_text()}"
        )
    return hdc, devmode_buffer


def validate_side_order(side_order: List[int], page_count: int) -> None:
    """提交打印任务前校验页序（纯函数，便于单测）。

    页码一旦越界，原实现会在 StartDoc 之后、循环中段才抛 IndexError，
    此时任务已提交、可能已经出纸，用户只能看到一个“未预期错误”。
    这里提前拦截，给出可读提示，且不会向队列提交任何任务。
    """
    if not side_order:
        raise PrintError("没有需要打印的页面。")
    if page_count <= 0:
        raise PrintError("PDF 没有可打印页面。")
    invalid = [page for page in side_order if page < 1 or page > page_count]
    if invalid:
        sample = "、".join(str(page) for page in invalid[:6])
        more = f" 等 {len(invalid)} 个" if len(invalid) > 6 else ""
        raise PrintError(
            f"页序中包含超出 PDF 页数（共 {page_count} 页）的页码：{sample}{more}。"
            "已中止本次打印，未向打印机提交任何任务。")


# ----------------------------------------------------------------------------
# GDI 打印
# ----------------------------------------------------------------------------
class GdiPrinter:
    """把指定页序的 PDF 输出面逐页送入 Windows 打印队列。"""

    def __init__(self, printer_name: str, options: PrinterOptions):
        self.printer_name = printer_name
        self.options = options
        self.last_job_id = 0
        self.pages_sent = 0          # 本次任务已成功结束（EndPage）的面数

    def print_pages(
        self,
        pdf_path: str,
        side_order: List[int],
        job_name: str,
        stop_event: threading.Event,
        on_page: Optional[Callable[[int, int, int], None]] = None,
        pause_event: Optional[threading.Event] = None,
    ) -> int:
        """按 side_order 顺序打印 PDF 的页面（1 基），返回打印任务 ID。

        提交阶段失败且“尚未成功出纸任何一面”时，最多再重试一次（退避后）。
        已经出纸的任务绝不重试，否则会造成重复页；这是对标 PDFtoPrinter /
        winprint 这类命令行打印工具的“干净失败再重提”做法。
        """
        from pdf_layout import open_pdf

        doc = open_pdf(pdf_path)
        try:
            # 先把页码越界挡在提交之前：原实现会在 StartDoc 之后才抛错，
            # 此时队列里可能已残留半个任务。
            validate_side_order(side_order, doc.page_count)
            attempt = 0
            while True:
                try:
                    return self._print_once(doc, side_order, job_name, stop_event,
                                            on_page, pause_event)
                except CancelledByUser:
                    raise
                except PrintError:
                    if attempt >= SUBMIT_RETRY_ATTEMPTS or self.pages_sent > 0:
                        raise
                    attempt += 1
                    time.sleep(RETRY_BACKOFF_SECONDS * attempt)
        finally:
            doc.close()

    def _print_once(
        self,
        doc,
        side_order: List[int],
        job_name: str,
        stop_event: threading.Event,
        on_page: Optional[Callable[[int, int, int], None]] = None,
        pause_event: Optional[threading.Event] = None,
    ) -> int:
        """单次打印尝试；失败时自行 AbortDoc / 取消残留任务后再抛出。"""
        hdc = None
        job_id = 0
        job_started = False
        self.pages_sent = 0
        try:
            hdc, _devmode_buffer = _create_printer_dc(self.printer_name, self.options)

            _gdi32.SetMapMode(hdc, MM_TEXT)
            _gdi32.SetStretchBltMode(hdc, STRETCH_HALFTONE)
            _gdi32.SetBrushOrgEx(hdc, 0, 0, None)

            printable_w = _gdi32.GetDeviceCaps(hdc, HORZRES)
            printable_h = _gdi32.GetDeviceCaps(hdc, VERTRES)
            if printable_w <= 0 or printable_h <= 0:
                raise PrintError("无法获取打印机可打印区域，请检查打印机驱动。")

            doc_info = _DOCINFOW()
            doc_info.cbSize = ctypes.sizeof(_DOCINFOW)
            doc_info.lpszDocName = job_name
            job_id = _gdi32.StartDocW(hdc, ctypes.byref(doc_info))
            if job_id <= 0:
                raise PrintError(f"创建打印任务失败：{_win_error_text()}")
            self.last_job_id = job_id
            job_started = True

            total = len(side_order)
            for index, rel in enumerate(side_order, start=1):
                if stop_event.is_set():
                    raise CancelledByUser("用户中止了打印。")
                while pause_event is not None and pause_event.is_set():
                    if stop_event.is_set():
                        raise CancelledByUser("用户中止了打印。")
                    time.sleep(0.2)

                page = doc[rel - 1]
                try:
                    raster = _rasterize(
                        page, self.options.dpi, self.options.color_mode,
                        ink_saver=getattr(self.options, "ink_saver", False),
                        annots=getattr(self.options, "annots", True),
                    )
                except Exception as exc:
                    raise PrintError(f"渲染 PDF 第 {rel} 页失败：{exc}") from exc
                dest = _fit_rect(raster.width, raster.height, printable_w, printable_h)

                bmi = _BITMAPINFOHEADER()
                bmi.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
                bmi.biWidth = raster.width
                bmi.biHeight = -raster.height   # 负值：自上而下，与 PyMuPDF 数据顺序一致
                bmi.biPlanes = 1
                bmi.biBitCount = 24
                bmi.biCompression = BI_RGB
                pixel_bits, dib_stride, _dib_keep = _prepare_dib(raster)
                bmi.biSizeImage = dib_stride * raster.height

                if _gdi32.StartPage(hdc) <= 0:
                    raise PrintError(f"开始打印页失败：{_win_error_text()}")
                result = _gdi32.StretchDIBits(
                    hdc, dest[0], dest[1], dest[2], dest[3],
                    0, 0, raster.width, raster.height,
                    pixel_bits, ctypes.byref(bmi), BI_RGB, SRCCOPY,
                )
                if result < 0:
                    raise PrintError(f"绘制页面失败：{_win_error_text()}")
                if _gdi32.EndPage(hdc) <= 0:
                    raise PrintError(f"结束打印页失败：{_win_error_text()}")
                self.pages_sent += 1

                if on_page:
                    on_page(index, total, rel)

            if _gdi32.EndDoc(hdc) <= 0:
                raise PrintError(f"提交打印任务失败：{_win_error_text()}")
            job_started = False
            return job_id
        except BaseException:
            if job_started and hdc:
                try:
                    _gdi32.AbortDoc(hdc)
                except Exception:
                    pass
            if job_id > 0:
                # 中途取消/出错时，尽力取消队列中可能残留的任务，避免继续出纸
                _try_cancel_job(self.printer_name, job_id)
            raise
        finally:
            if hdc:
                _gdi32.DeleteDC(hdc)
