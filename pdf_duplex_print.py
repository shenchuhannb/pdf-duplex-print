# -*- coding: utf-8 -*-
"""PDF 批量双面打印工具 —— 程序入口。

本文件保持为唯一启动入口（桌面快捷方式指向它），内部逻辑拆分到：
  * core_utils.py    —— 异常、页码解析、拖拽数据解析等通用工具
  * pdf_layout.py    —— 页码/补页/双面页序/N-up/小册子/排版合成
  * printer_core.py  —— GDI 打印、DEVMODE、打印机状态与队列监听
  * settings_store.py—— 打印设置记忆
  * theme.py         —— 界面风格 A（浅色专业，唯一风格）
  * ui_main.py       —— 主界面与后台打印线程

命令行：
    python pdf_duplex_print.py               # 正常启动（风格 A）
"""

import argparse
import ctypes
import os
import sys
import tempfile
import time
import tkinter as tk
from tkinter import messagebox

from ui_main import APP_TITLE, DuplexPrintApp, _create_root  # noqa: F401

import app_identity as _identity
import printer_core as _printer_core
import pdf_layout as _pdf_layout
import theme as _theme
from log_setup import get_logger


def _check_dependencies() -> bool:
    """启动前检查 PyMuPDF / pywin32，缺失时给出中文安装提示。"""
    missing = []
    if getattr(_pdf_layout, "fitz", None) is None:
        missing.append("PyMuPDF（安装命令：pip install PyMuPDF）")
    if getattr(_printer_core, "win32print", None) is None:
        missing.append("pywin32（安装命令：pip install pywin32）")
    if not missing:
        return True

    message = "缺少运行依赖：\n\n" + "\n".join(missing)
    print(message)
    try:
        probe = tk.Tk()
        probe.withdraw()
        messagebox.showerror(APP_TITLE, message)
        probe.destroy()
    except Exception:
        pass
    return False


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(description="PDF 批量双面打印工具")
    parser.add_argument("--style", default="", choices=("", "A", "a"),
                        help="界面风格：仅保留 A（浅色专业）")
    # 隐藏诊断参数：构建完界面后立即退出，全程窗口透明不可见（供安装/升级自检）
    parser.add_argument("--self-check", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def _self_check(root, args, app) -> None:
    """隐藏窗口自检：验证入口、图标与界面可正常构建，全程不显示窗口。

    结果写入环境变量 PDF_DUPLEX_SELF_CHECK_FILE 指定的文件（可选），
    便于外部脚本确认启动成功，同时不产生任何可见窗口或残留进程。
    """
    root.withdraw()
    root.geometry("1280x980+40+40")
    root.attributes("-alpha", 0.0)   # 全透明：即使被映射也不可见
    root.update_idletasks()
    root.deiconify()
    for _ in range(3):
        root.update()
        time.sleep(0.05)
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
    lines = [
        f"ok=True",
        f"title={root.title()}",
        f"style={app.style_key}",
        f"aumid={_identity.get_process_app_user_model_id()}",
        f"icon_exists={os.path.isfile(icon_path)}",
        f"icon_size={os.path.getsize(icon_path) if os.path.isfile(icon_path) else 0}",
        f"temp_root={os.path.join(tempfile.gettempdir(), 'PDFDuplexPrint-temp')}",
        f"stale_temp_removed={len(getattr(app, 'stale_temp_removed', []))}",
        f"argv={' '.join(sys.argv[1:])}",
        f"cwd={os.getcwd()}",
        f"pid={os.getpid()}",
    ]
    target = os.environ.get("PDF_DUPLEX_SELF_CHECK_FILE", "")
    if target:
        try:
            with open(target, "w", encoding="utf-8") as handle:
                handle.write("\n".join(lines) + "\n")
        except OSError:
            pass
    else:
        print("\n".join(lines))
    # 可选保持存活若干秒（窗口保持全透明），供外部脚本核对进程命令行后关闭
    try:
        hold = float(os.environ.get("PDF_DUPLEX_SELF_CHECK_HOLD", "0") or 0)
    except ValueError:
        hold = 0.0
    end = time.time() + max(0.0, hold)
    while time.time() < end:
        try:
            root.update()
        except Exception:
            break
        time.sleep(0.05)
    # 先置退出态并空转一次，再销毁，避免 Tcl 残留 after 回调
    app.shutting_down = True
    app.preview_enabled = False
    try:
        root.update()
    except Exception:
        pass
    root.destroy()


def main(argv=None) -> None:
    # 日志落盘（%APPDATA%\PDFDuplexPrint\logs\app.log），便于打印失败后排查
    get_logger()

    # 任务栏图标/分组：必须在创建 Tk 窗口之前给进程设置显式 AppUserModelID，
    # 否则 pythonw.exe 启动时任务栏会沿用宿主进程的 python 图标。
    _identity.set_process_app_user_model_id()

    if not _check_dependencies():
        sys.exit(1)

    # 高分屏适配：per-monitor DPI aware，避免文字与图标模糊
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    args = _parse_args(argv)
    root = _create_root()
    # 首帧不闪白：Tk 根窗口创建即映射，若此时界面还没构建完，用户会先看到
    # 一块默认白底空窗。这里先 withdraw，并用最终主题的底色填充，等全部控件、
    # ttk 样式、几何与首帧布局完成后，再一次性 deiconify（不产生二段跳）。
    try:
        root.withdraw()
        root.configure(background=_theme.palette("A")["bg"])
    except tk.TclError:
        pass
    try:
        root.option_add("*Font", ("Microsoft YaHei UI", 9))
    except tk.TclError:
        pass
    app = DuplexPrintApp(root, style=args.style)
    if args.self_check or os.environ.get("PDF_DUPLEX_SELF_CHECK") == "1":
        _self_check(root, args, app)
        return
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    # 构建完毕再一次性显示：geometry 已在 __init__ 中由 _apply_window_geometry 定好
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.mainloop()


if __name__ == "__main__":
    main()
