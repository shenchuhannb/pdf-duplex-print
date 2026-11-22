# -*- coding: utf-8 -*-
"""修复 Windows 下 Tk 窗口最大化/后台还原时的黑块（CodeArts 最终方案）。

两层机制，各管各的场景：

1. 类背景画刷（GCL_HBRBACKGROUND，主修复）
   Tk 顶层窗口类（TkTopLevel）的背景画刷为 NULL（实测 SetClassLongPtrW
   返回旧值 0）。最大化尺寸突变时，wrapper 客户区与内部 TkChild 之间
   的边框带（顶部约 31px、四周约 8px）由 DWM 用类画刷初始化——NULL
   画刷即呈现黑色（屏幕级采集实测暗像素峰值 19.7%）。把类画刷设为
   主题底色后，新暴露区域从源头就是底色。

   历史教训：此前在 WndProc 层面"事后刷"（WM_SIZE/WM_NCPAINT 整窗
   FillRect）全 部失败——DWM 在窗口过程收到消息之前就合成了黑帧；
   强行整窗刷会把已绘制的 UI 覆盖成白块；同步 RDW_FRAME 重绘也来不及。
   WS_EX_COMPOSITED 在 Tk 上会死锁或毁窗，同样不可用。

   注意：SetClassLongPtrW 影响本进程内所有 TkTopLevel 窗口（类是进程
   内注册的），效果正是所需：所有 toplevel 最大化/还原都不再闪黑。

2. WM_ERASEBKGND 子类化（辅修复）
   最小化或被遮挡后还原时，客户区尚未重绘的区域会短暂显示黑色。
   子类化 wrapper 拦截 WM_ERASEBKGND，用主题底色立即填充客户区并
   返回 1，控件随后照常绘制。

设置类背景画刷的失败会被静默忽略（仅失去主修复，不影响运行）。
"""

import ctypes
import sys
import time

GWL_WNDPROC = -4
GCL_HBRBACKGROUND = -10
WM_ERASEBKGND = 0x0014
WM_SIZE = 0x0005
WM_NCPAINT = 0x0085
SIZE_MINIMIZED = 1
DWMWA_TRANSITIONS_FORCEDISABLED = 3

# hwnd -> (WNDPROC 回调, 原窗口过程)；必须常驻引用防止回调被 GC。
# 高级设置等对话框每次打开是新 hwnd，旧条目仅几十字节，保留不清理。
_installed = {}

_class_brush_set = False


def install_erase_background(win, background: str) -> None:
    """为 Tk 顶层窗口安装黑块修复（幂等，非 Windows 或失败时静默跳过）。"""
    if sys.platform != "win32":
        return
    try:
        _install(win, background)
    except Exception:
        pass


def _install(win, background: str) -> None:
    global _class_brush_set
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                    ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    WNDPROC = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint,
        ctypes.c_size_t, ctypes.c_ssize_t)
    user32.SetWindowLongPtrW.restype = ctypes.c_void_p
    user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                         ctypes.c_void_p]
    user32.CallWindowProcW.restype = ctypes.c_ssize_t
    user32.CallWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                       ctypes.c_uint, ctypes.c_size_t,
                                       ctypes.c_ssize_t]
    user32.GetClientRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.ClientToScreen.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.GetWindowDC.restype = ctypes.c_void_p
    user32.GetWindowDC.argtypes = [ctypes.c_void_p]
    user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.FillRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                ctypes.c_void_p]
    user32.SetClassLongPtrW.restype = ctypes.c_ssize_t
    user32.SetClassLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                        ctypes.c_ssize_t]
    gdi32.CreateSolidBrush.restype = ctypes.c_void_p
    gdi32.CreateSolidBrush.argtypes = [ctypes.c_ulong]

    win.update_idletasks()
    hwnd = int(win.wm_frame(), 16)
    if not hwnd or hwnd in _installed:
        return
    rgb = background.lstrip("#")
    if len(rgb) != 6:
        return
    colorref = int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)  # COLORREF 为 BGR

    # 1) 主修复：类背景画刷。只在首个 toplevel 上设置一次（类共享）。
    if not _class_brush_set:
        brush = gdi32.CreateSolidBrush(colorref)
        if brush:
            old = user32.SetClassLongPtrW(
                ctypes.c_void_p(hwnd), GCL_HBRBACKGROUND, brush)
            if old != 0 or brush:
                # 旧值为 0（NULL）是预期情况；设置成功后所有 TkTopLevel
                # 的新暴露区域都由该画刷初始化。画刷交给系统管理，不删除。
                _class_brush_set = True

    # 2) 最大化/还原的系统过渡会先展示 DWM 尚未重绘的旧表面。
    try:
        dwmapi = ctypes.windll.dwmapi
        dwmapi.DwmSetWindowAttribute.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint]
        disabled = ctypes.c_int(1)
        dwmapi.DwmSetWindowAttribute(
            ctypes.c_void_p(hwnd), DWMWA_TRANSITIONS_FORCEDISABLED,
            ctypes.byref(disabled), ctypes.sizeof(disabled))
    except Exception:
        pass

    # 3) 辅修复：子类化 WM_ERASEBKGND，后台还原时客户区不再闪黑。
    old_proc = [None]
    size_time = [0.0]

    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    def _fill_frame_bands(h):
        """只填窗口矩形减去客户区的四条边框带（不碰客户区内容）。"""
        hdc = user32.GetWindowDC(h)
        if not hdc:
            return
        try:
            wr = RECT()
            user32.GetWindowRect(h, ctypes.byref(wr))
            ww, wh = wr.right - wr.left, wr.bottom - wr.top
            cr = RECT()
            user32.GetClientRect(h, ctypes.byref(cr))
            pt = POINT(cr.left, cr.top)
            user32.ClientToScreen(h, ctypes.byref(pt))
            cx, cy = pt.x - wr.left, pt.y - wr.top  # 客户区在窗口坐标的偏移
            cw, ch = cr.right - cr.left, cr.bottom - cr.top
            bands = (
                RECT(0, 0, ww, cy),                       # 顶部（标题栏）
                RECT(0, cy, cx, cy + ch),                 # 左
                RECT(cx + cw, cy, ww, cy + ch),           # 右
                RECT(0, cy + ch, ww, wh),                 # 底
            )
            for band in bands:
                if band.right > band.left and band.bottom > band.top:
                    user32.FillRect(ctypes.c_void_p(hdc), ctypes.byref(band),
                                    _eraser[0])
        finally:
            user32.ReleaseDC(h, ctypes.c_void_p(hdc))

    def _wnd_proc(h, msg, wparam, lparam):
        if msg == WM_ERASEBKGND and wparam:  # wparam 为目标 DC
            rect = RECT()
            user32.GetClientRect(h, ctypes.byref(rect))
            user32.FillRect(ctypes.c_void_p(wparam), ctypes.byref(rect),
                            _eraser[0])
            return 1
        if msg == WM_NCPAINT:
            # 尺寸变化后的短窗口内，框架重绘前先把四条边框带打底，
            # 新扩展的非客户区不再露出黑色旧表面；系统随后照常绘制
            # 标题栏。客户区不受影响（不会刷掉 UI）。拖动日志分隔条
            # 不经过 wrapper 的 WM_SIZE，节流窗口外零开销。
            if time.monotonic() - size_time[0] < 0.5:
                _fill_frame_bands(h)
            return user32.CallWindowProcW(old_proc[0], h, msg, wparam, lparam)
        if msg == WM_SIZE and wparam != SIZE_MINIMIZED:
            size_time[0] = time.monotonic()
        return user32.CallWindowProcW(old_proc[0], h, msg, wparam, lparam)

    eraser = gdi32.CreateSolidBrush(colorref)
    if not eraser:
        return
    _eraser = [eraser]
    callback = WNDPROC(_wnd_proc)
    old_proc[0] = user32.SetWindowLongPtrW(
        ctypes.c_void_p(hwnd), GWL_WNDPROC, ctypes.cast(callback, ctypes.c_void_p))
    if not old_proc[0]:
        return
    _installed[hwnd] = (callback, old_proc[0], _eraser[0])
