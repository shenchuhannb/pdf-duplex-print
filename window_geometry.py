# -*- coding: utf-8 -*-
"""窗口几何：启动居中、位置记忆与可见性校验。

背景：
  * 之前的版本从不设置窗口位置，Windows 会把它放在屏幕左上角附近，
    看起来“不居中”；
  * 一旦开始记忆几何，就必须处理分辨率变化、外接显示器拔掉、配置里被写入
    屏幕外/负坐标（例如离屏截图脚本留下的 geometry）等情况，否则窗口会跑到
    看不见的地方。

处理规则（纯函数，便于单元测试）：
  1. 保存的几何若尺寸合法、且标题栏区域仍落在某块显示器工作区内，则恢复；
     同时把窗口夹回该工作区，避免贴边或越界。
  2. 否则（无记录 / 负坐标 / 超出分辨率 / 显示器被移除 / 尺寸过小）→ 在
     主显示器工作区内居中显示。
  3. 尺寸永远夹在 [min_size, 工作区] 之间，保证控件不被挤坏、窗口不出屏。
  4. Win32 枚举出来的工作区是**物理像素**；Tk 与 pywebview 的 geometry 是
     **逻辑像素**（96 DPI 基准）。`choose_logical_geometry` /
     `get_window_geometry` 负责按当前 DPI（例如 150% → 1.5）换算，
     避免恢复后窗口尺寸被放大 1.5 倍或缩回屏幕外。
"""

import ctypes
import math
import re
from ctypes import wintypes
from typing import Iterable, List, Optional, Sequence, Tuple

Rect = Tuple[int, int, int, int]          # (left, top, right, bottom)
Geometry = Tuple[int, int, int, int]      # (width, height, x, y)

DEFAULT_SIZE = (1280, 960)
# 最小宽度按用户实际窗口内容宽度（约 1024px）设定：1024×720 下所有控件必须完整可见
MIN_SIZE = (1024, 720)
# 判断“窗口还看得见”的最小标题栏区域（宽 x 高）
TITLEBAR_STRIP = (160, 40)


_GEOMETRY_RE = re.compile(r"^(\d+)x(\d+)([+-]\d+)([+-]\d+)$")


def parse_geometry(text: str) -> Optional[Geometry]:
    """解析 Tk 的 "WxH+X+Y" / "WxH-X-Y" 字符串。

    返回 None 表示格式非法；负偏移（例如离屏截图用的坐标）按原样解析，
    由选择逻辑判定是否可见，不在这里直接丢弃。
    """
    if not text or not isinstance(text, str):
        return None
    match = _GEOMETRY_RE.match(text.strip())
    if not match:
        return None
    return tuple(int(value) for value in match.groups())  # type: ignore[return-value]


def format_geometry(width: int, height: int, x: int, y: int) -> str:
    """转成 Tk 可识别的几何字符串（负坐标用 - 表示，正坐标带 +）。"""
    return f"{int(width)}x{int(height)}{int(x):+d}{int(y):+d}"


def _clamp_size(size: Sequence[int], work_area: Rect,
                min_size: Sequence[int] = MIN_SIZE) -> Tuple[int, int]:
    """把窗口尺寸夹到“最小尺寸 ~ 工作区尺寸”之间。"""
    area_w = max(200, work_area[2] - work_area[0])
    area_h = max(150, work_area[3] - work_area[1])
    min_w = min(min_size[0], area_w)
    min_h = min(min_size[1], area_h)
    width = max(min_w, min(int(size[0]), area_w))
    height = max(min_h, min(int(size[1]), area_h))
    return width, height


def centered_in(work_area: Rect, size: Sequence[int] = DEFAULT_SIZE,
                min_size: Sequence[int] = MIN_SIZE) -> Geometry:
    """在工作区内居中，返回 (宽, 高, x, y)。"""
    width, height = _clamp_size(size, work_area, min_size)
    x = work_area[0] + max(0, (work_area[2] - work_area[0] - width) // 2)
    y = work_area[1] + max(0, (work_area[3] - work_area[1] - height) // 2)
    return width, height, x, y


def _titlebar_visible(geometry: Geometry, work_area: Rect) -> bool:
    """窗口顶部一条（标题栏/拖动区）是否有足够部分落在工作区内。"""
    width, height, x, y = geometry
    strip_w = min(width, TITLEBAR_STRIP[0])
    strip_h = min(height, TITLEBAR_STRIP[1])
    overlap_w = min(x + strip_w, work_area[2]) - max(x, work_area[0])
    overlap_h = min(y + strip_h, work_area[3]) - max(y, work_area[1])
    return overlap_w >= strip_w * 0.9 and overlap_h >= strip_h * 0.9


def _clamp_into(geometry: Geometry, work_area: Rect,
                min_size: Sequence[int] = MIN_SIZE) -> Geometry:
    """保持尺寸，把左上角夹到工作区内（避免贴边/越界）。"""
    width, height, x, y = geometry
    width, height = _clamp_size((width, height), work_area, min_size)
    x = max(work_area[0], min(x, work_area[2] - width))
    y = max(work_area[1], min(y, work_area[3] - height))
    return width, height, x, y


def choose_geometry(
    saved: str,
    work_areas: Iterable[Rect],
    default_size: Sequence[int] = DEFAULT_SIZE,
    min_size: Sequence[int] = MIN_SIZE,
) -> Geometry:
    """根据保存的几何与显示器工作区，决定最终窗口位置。

    work_areas 的第一项视为主显示器工作区（Windows 枚举顺序的主屏通常在前，
    调用方也可以显式传入主屏优先的列表）。work_areas 与 saved 必须使用同一
    坐标系：物理工作区请先换算，或直接使用 `choose_logical_geometry`。
    """
    areas: List[Rect] = [tuple(int(v) for v in area) for area in work_areas if area]
    if not areas:
        # 极端情况：拿不到显示器信息，用 1024x768 兜底，仍然居中
        areas = [(0, 0, 1280, 800)]
    primary = areas[0]

    parsed = parse_geometry(saved)
    if parsed is not None:
        width, height, x, y = parsed
        if width >= max(min_size[0], 200) and height >= max(min_size[1], 150):
            geometry = (width, height, x, y)
            for area in areas:
                if _titlebar_visible(geometry, area):
                    return _clamp_into(geometry, area, min_size)
    # 没有可用记录：主屏工作区居中
    return centered_in(primary, default_size, min_size)


# ----------------------------------------------------------------------------
# DPI：Win32 物理像素 ↔ pywebview/Tk 逻辑像素（96 DPI 基准）
# ----------------------------------------------------------------------------
DEFAULT_DPI_SCALE = 1.0
LOGPIXELS_PER_INCH = 96

# GetWindowPlacement 的 showCmd：只有 SW_SHOWNORMAL 的 GetWindowRect 才可保存；
# 最大化/最小化必须取 rcNormalPosition（窗口还原后的位置），否则会把整块屏幕
# 的尺寸或 -32000 之类的离屏坐标写进配置。
SW_SHOWNORMAL = 1


def _scale_or_default(scale: float) -> float:
    try:
        value = float(scale)
    except (TypeError, ValueError):
        return DEFAULT_DPI_SCALE
    return value if value > 0 else DEFAULT_DPI_SCALE


def _round(value: float) -> int:
    """四舍五入（half-up）；不用内置 round 以免 .5 的银行家舍入造成抖动。"""
    return int(math.floor(float(value) + 0.5))


def physical_geometry_to_logical(geometry: Sequence[int], scale: float) -> Geometry:
    """物理像素几何 (宽, 高, x, y) → 逻辑像素几何（pywebview/Tk 口径）。"""
    factor = _scale_or_default(scale)
    width, height, x, y = (int(v) for v in geometry)
    return (_round(width / factor), _round(height / factor),
            _round(x / factor), _round(y / factor))


def logical_geometry_to_physical(geometry: Sequence[int], scale: float) -> Geometry:
    """逻辑像素几何 → 物理像素几何（用于校验/测试，窗口创建由 pywebview 换算）。"""
    factor = _scale_or_default(scale)
    width, height, x, y = (int(v) for v in geometry)
    return (_round(width * factor), _round(height * factor),
            _round(x * factor), _round(y * factor))


def rect_to_geometry(rect: Rect, scale: float) -> Geometry:
    """物理矩形 (left, top, right, bottom) → 逻辑几何 (宽, 高, x, y)。"""
    left, top, right, bottom = (int(v) for v in rect)
    return physical_geometry_to_logical(
        (right - left, bottom - top, left, top), scale)


def geometry_visible(geometry: Geometry, work_areas: Iterable[Rect]) -> bool:
    """几何的标题栏区域是否仍完整落在任一工作区内（同一坐标系）。"""
    return any(_titlebar_visible(geometry, area)
               for area in work_areas if area)


def logical_work_areas(work_areas: Iterable[Rect], scale: float) -> List[Rect]:
    """物理工作区列表 → 逻辑像素工作区列表（pywebview/Tk 口径）。"""
    factor = _scale_or_default(scale)
    return [(_round(int(area[0]) / factor), _round(int(area[1]) / factor),
             _round(int(area[2]) / factor), _round(int(area[3]) / factor))
            for area in work_areas if area]


def choose_logical_geometry(
    saved: str,
    work_areas: Iterable[Rect],
    scale: float = DEFAULT_DPI_SCALE,
    default_size: Sequence[int] = DEFAULT_SIZE,
    min_size: Sequence[int] = MIN_SIZE,
) -> Geometry:
    """按 DPI 把物理工作区换算成逻辑坐标后再决定窗口几何。

    saved 与返回值都是逻辑像素（Tk/pywebview 的 geometry 口径），
    work_areas 是 Win32 的物理像素工作区（`list_work_areas()` 的返回值）。
    保存的几何落在某块可见工作区内时恢复；否则主屏居中回退。
    """
    return choose_geometry(saved, logical_work_areas(work_areas, scale),
                           default_size, min_size)


# ----------------------------------------------------------------------------
# Win32：DPI 查询
# ----------------------------------------------------------------------------
def system_dpi_scale() -> float:
    """系统 DPI 缩放（1.0 = 96 DPI，1.5 = 150%）。

    必须在进程声明 DPI 感知之后调用；否则 Windows 会虚拟化成 96 DPI，
    150% 显示器也会返回 1.0。无法获取时回退 1.0。
    """
    try:
        user32 = ctypes.windll.user32
        get_dpi = getattr(user32, "GetDpiForSystem", None)
        if get_dpi is not None:
            get_dpi.argtypes = []
            get_dpi.restype = wintypes.UINT
            dpi = int(get_dpi())
            if dpi > 0:
                return max(DEFAULT_DPI_SCALE, dpi / LOGPIXELS_PER_INCH)
    except Exception:
        pass
    try:
        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32
        hdc = user32.GetDC(None)
        try:
            dpi = int(gdi32.GetDeviceCaps(ctypes.c_void_p(hdc), 88))  # LOGPIXELSX
        finally:
            user32.ReleaseDC(None, ctypes.c_void_p(hdc))
        if dpi > 0:
            return max(DEFAULT_DPI_SCALE, dpi / LOGPIXELS_PER_INCH)
    except Exception:
        pass
    return DEFAULT_DPI_SCALE


def window_dpi_scale(hwnd: int = 0) -> float:
    """指定窗口的 DPI 缩放（多屏/每窗口 DPI）；失败时退回系统 DPI。"""
    if hwnd:
        try:
            get_dpi = getattr(ctypes.windll.user32, "GetDpiForWindow", None)
            if get_dpi is not None:
                get_dpi.argtypes = [ctypes.c_void_p]
                get_dpi.restype = wintypes.UINT
                dpi = int(get_dpi(ctypes.c_void_p(int(hwnd))))
                if dpi > 0:
                    return max(DEFAULT_DPI_SCALE, dpi / LOGPIXELS_PER_INCH)
        except Exception:
            pass
    return system_dpi_scale()


# ----------------------------------------------------------------------------
# Win32：读取窗口几何（最大化/最小化取还原态）
# ----------------------------------------------------------------------------
class _WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", wintypes.UINT),
        ("flags", wintypes.UINT),
        ("showCmd", wintypes.UINT),
        ("ptMinPosition", wintypes.POINT),
        ("ptMaxPosition", wintypes.POINT),
        ("rcNormalPosition", wintypes.RECT),
    ]


def get_window_geometry(hwnd: int) -> Optional[Geometry]:
    """读取窗口当前几何并换算为逻辑像素；最大/最小化返回还原态。

    返回 (宽, 高, x, y)，与 Tk/pywebview 的 geometry 口径一致；
    句柄无效或查询失败时返回 None（调用方只记录日志，不影响关闭）。
    """
    if not hwnd:
        return None
    try:
        user32 = ctypes.windll.user32
        handle = ctypes.c_void_p(int(hwnd))
        rect: Optional[Tuple[int, int, int, int]] = None

        placement = _WINDOWPLACEMENT()
        placement.length = ctypes.sizeof(_WINDOWPLACEMENT)
        user32.GetWindowPlacement.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(_WINDOWPLACEMENT)]
        user32.GetWindowPlacement.restype = wintypes.BOOL
        placed = bool(user32.GetWindowPlacement(handle, ctypes.byref(placement)))
        if placed and int(placement.showCmd) != SW_SHOWNORMAL:
            normal = placement.rcNormalPosition
            if normal.right > normal.left and normal.bottom > normal.top:
                rect = (int(normal.left), int(normal.top),
                        int(normal.right), int(normal.bottom))

        if rect is None:
            box = wintypes.RECT()
            user32.GetWindowRect.argtypes = [
                ctypes.c_void_p, ctypes.POINTER(wintypes.RECT)]
            user32.GetWindowRect.restype = wintypes.BOOL
            if not user32.GetWindowRect(handle, ctypes.byref(box)):
                return None
            rect = (int(box.left), int(box.top),
                    int(box.right), int(box.bottom))

        if rect[2] <= rect[0] or rect[3] <= rect[1]:
            return None
        return rect_to_geometry(rect, window_dpi_scale(int(hwnd)))
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Win32：枚举显示器工作区
# ----------------------------------------------------------------------------
class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", wintypes.RECT),
        ("rcWork", wintypes.RECT),
        ("dwFlags", wintypes.DWORD),
    ]


def list_work_areas() -> List[Rect]:
    """枚举所有显示器的工作区（不含任务栏），主显示器排在第一位。

    失败时退回 Tk/Win32 的主屏尺寸减去任务栏高度，保证不会返回空列表。
    """
    areas: List[Rect] = []
    primary: Optional[Rect] = None
    try:
        user32 = ctypes.windll.user32
        monitor_enum_proc = ctypes.WINFUNCTYPE(
            ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p,
            ctypes.POINTER(wintypes.RECT), ctypes.c_void_p)

        def _callback(hmonitor, _hdc, _rect, _data):
            nonlocal primary
            info = _MONITORINFO()
            info.cbSize = ctypes.sizeof(_MONITORINFO)
            if user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
                work = info.rcWork
                rect = (int(work.left), int(work.top),
                        int(work.right), int(work.bottom))
                areas.append(rect)
                # MONITORINFOF_PRIMARY = 1
                if int(info.dwFlags) & 1:
                    primary = rect
            return 1

        user32.EnumDisplayMonitors(None, None, monitor_enum_proc(_callback), None)
    except Exception:
        areas = []
        primary = None
    if areas and primary is not None:
        # 主显示器必须排第一：居中与“工作区”都以主屏为准
        areas = [primary] + [area for area in areas if area != primary]
    if not areas:
        try:
            width = int(ctypes.windll.user32.GetSystemMetrics(0))
            height = int(ctypes.windll.user32.GetSystemMetrics(1))
            areas = [(0, 0, width, max(600, height - 48))]
        except Exception:
            areas = [(0, 0, 1280, 800)]
    return areas
