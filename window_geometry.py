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
  3. 尺寸永远夹在 [MIN_SIZE, 工作区] 之间，保证控件不被挤坏、窗口不出屏。
"""

import ctypes
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


def _clamp_size(size: Sequence[int], work_area: Rect) -> Tuple[int, int]:
    """把窗口尺寸夹到“最小尺寸 ~ 工作区尺寸”之间。"""
    area_w = max(200, work_area[2] - work_area[0])
    area_h = max(150, work_area[3] - work_area[1])
    min_w = min(MIN_SIZE[0], area_w)
    min_h = min(MIN_SIZE[1], area_h)
    width = max(min_w, min(int(size[0]), area_w))
    height = max(min_h, min(int(size[1]), area_h))
    return width, height


def centered_in(work_area: Rect, size: Sequence[int] = DEFAULT_SIZE) -> Geometry:
    """在工作区内居中，返回 (宽, 高, x, y)。"""
    width, height = _clamp_size(size, work_area)
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


def _clamp_into(geometry: Geometry, work_area: Rect) -> Geometry:
    """保持尺寸，把左上角夹到工作区内（避免贴边/越界）。"""
    width, height, x, y = geometry
    width, height = _clamp_size((width, height), work_area)
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
    调用方也可以显式传入主屏优先的列表）。
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
                    return _clamp_into(geometry, area)
    # 没有可用记录：主屏工作区居中
    return centered_in(primary, default_size)


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
