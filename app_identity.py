# -*- coding: utf-8 -*-
"""进程/窗口身份：任务栏图标与分组。

Windows 任务栏按 AppUserModelID 分组并取该组的图标。用 pythonw.exe 启动时，
默认分组是 pythonw.exe，任务栏按钮就显示 python 图标；只调用 iconbitmap
不能改变任务栏按钮的图标来源。这里的做法是：

* 在创建 Tk 窗口前给进程设置显式 AppUserModelID（与 .lnk 里的
  System.AppUserModel.ID 保持一致），任务栏按此分组并读取窗口图标；
* 给窗口显式发送 WM_SETICON（ICON_BIG / ICON_SMALL / ICON_SMALL2），
  并同步设置窗口类的 GCLP_HICON / GCLP_HICONS，覆盖标题栏、Alt+Tab、
  任务栏缩略图等所有取图标路径；
* HICON 句柄由本模块持有，避免被 GC 提前销毁导致任务栏回退到宿主图标。
"""

import ctypes
import os
from ctypes import wintypes

# 与桌面快捷方式 .lnk 的 System.AppUserModel.ID 必须完全一致
APP_USER_MODEL_ID = "PDFTools.PDFDuplexPrint.1.0"

_ICON_HANDLES = []

IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
WM_SETICON = 0x0080
ICON_SMALL = 0
ICON_BIG = 1
ICON_SMALL2 = 2
GCLP_HICON = -14
GCLP_HICONS = -34
GA_ROOT = 2

_user32 = ctypes.windll.user32
_shell32 = ctypes.windll.shell32
_user32.LoadImageW.restype = wintypes.HANDLE
_user32.SendMessageW.restype = ctypes.c_ssize_t
_user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]


def set_process_app_user_model_id(aumid: str = APP_USER_MODEL_ID) -> bool:
    """设置当前进程的显式 AUMID（必须在创建窗口前调用）。"""
    try:
        _shell32.SetCurrentProcessExplicitAppUserModelID(ctypes.c_wchar_p(aumid))
        return True
    except Exception:
        return False


def get_process_app_user_model_id() -> str:
    """读回当前进程的显式 AUMID；未设置时返回空字符串。"""
    value = ctypes.c_wchar_p()
    try:
        result = _shell32.GetCurrentProcessExplicitAppUserModelID(ctypes.byref(value))
    except Exception:
        return ""
    try:
        if result != 0:
            return ""
        return value.value or ""
    finally:
        try:
            ctypes.windll.ole32.CoTaskMemFree(ctypes.cast(value, ctypes.c_void_p))
        except Exception:
            pass


def load_icon(ico_path: str, size: int):
    """从 .ico 里按目标尺寸加载 HICON（找不到精确尺寸时由系统缩放）。"""
    return _user32.LoadImageW(None, ctypes.c_wchar_p(ico_path), IMAGE_ICON,
                              int(size), int(size), LR_LOADFROMFILE)


def apply_window_icon(window, ico_path: str) -> dict:
    """给 Tk 窗口设置大小两种图标，返回各路径的实际 HICON 句柄。"""
    handles = {}
    if not ico_path or not os.path.isfile(ico_path):
        return handles
    try:
        window.iconbitmap(ico_path)      # 标题栏默认图标（Tk 路径）
    except Exception:
        pass
    try:
        hwnd = _user32.GetAncestor(int(window.winfo_id()), GA_ROOT)
    except Exception:
        hwnd = 0
    for name, size in (("small", 16), ("big", 32)):
        hicon = load_icon(ico_path, size)
        if not hicon:
            continue
        handles[name] = hicon
        _ICON_HANDLES.append(hicon)      # 保持引用，防止句柄被回收
        if hwnd:
            _user32.SendMessageW(hwnd, WM_SETICON,
                                 ICON_SMALL if size <= 16 else ICON_BIG, hicon)
            if size <= 16:
                _user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL2, hicon)
    if hwnd:
        try:
            _user32.SetClassLongPtrW(hwnd, GCLP_HICON, handles.get("big") or 0)
            _user32.SetClassLongPtrW(hwnd, GCLP_HICONS, handles.get("small") or 0)
        except Exception:
            pass
    return handles
