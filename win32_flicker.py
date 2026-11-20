# -*- coding: utf-8 -*-
"""修复 Windows 下 Tk 窗口从后台还原时闪黑色方块。

Tk 顶层窗口类（TkTopLevel）的背景画刷是 NULL：最小化或被遮挡后还原时，
Windows 会先把尚未重绘的区域呈现为黑色，Tk 再逐个控件绘制，首帧就能
看到几块黑色区域（实测还原后首帧近黑像素可达 88%）。

这里子类化顶层窗口的 WndProc，仅把 WM_ERASEBKGND 改为用主题底色立即
填充整个客户区并返回 1，其余消息全部转发原窗口过程。控件随后照常绘制，
未绘制的间隙显示界面底色而非黑色。设置类背景画刷（GCL_HBRBACKGROUND）
的替代方案实测会导致窗口持续全黑，不可用，因此采用子类化。
"""

import ctypes
import sys

GWL_WNDPROC = -4
WM_ERASEBKGND = 0x0014

# hwnd -> (WNDPROC 回调, 原窗口过程, 底色画刷)；必须常驻引用防止回调被 GC。
# 高级设置等对话框每次打开是新 hwnd，旧条目仅几百字节，保留不清理。
_installed = {}


def install_erase_background(win, background: str) -> None:
    """为 Tk 顶层窗口安装底色擦除（幂等，非 Windows 或失败时静默跳过）。"""
    if sys.platform != "win32":
        return
    try:
        _install(win, background)
    except Exception:
        pass


def _install(win, background: str) -> None:
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
    user32.FillRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                ctypes.c_void_p]
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
    brush = gdi32.CreateSolidBrush(colorref)
    if not brush:
        return
    old_proc = [None]

    def _wnd_proc(h, msg, wparam, lparam):
        if msg == WM_ERASEBKGND and wparam:  # wparam 为目标 DC
            rect = RECT()
            user32.GetClientRect(h, ctypes.byref(rect))
            user32.FillRect(ctypes.c_void_p(wparam), ctypes.byref(rect),
                            ctypes.c_void_p(brush))
            return 1
        return user32.CallWindowProcW(old_proc[0], h, msg, wparam, lparam)

    callback = WNDPROC(_wnd_proc)
    old_proc[0] = user32.SetWindowLongPtrW(
        ctypes.c_void_p(hwnd), GWL_WNDPROC, ctypes.cast(callback, ctypes.c_void_p))
    if not old_proc[0]:
        return
    _installed[hwnd] = (callback, old_proc[0], brush)
