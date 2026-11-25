# -*- coding: utf-8 -*-
"""pywebview 版入口：UI 用 WebView2（GPU 合成，resize 丝滑），后端 100% 复用。

路由方案：自建 bottle WSGI app 精确控制静态资源根目录，
绕开 pywebview 相对路径的 commonpath 解析（实测在子目录布局下 404）。
"""
import ctypes
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bottle
import webview

import window_geometry
from app_identity import set_process_app_user_model_id
from log_setup import get_logger
from webapp.bridge import AppBridge, WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE

MIN_SIZE = WINDOW_MIN_SIZE
DEFAULT_SIZE = WINDOW_DEFAULT_SIZE
WEBAPP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(WEBAPP_DIR, "static")
APP_ICON = os.path.join(os.path.dirname(WEBAPP_DIR), "app_icon.ico")

app = bottle.Bottle()


@app.route("/")
def index():
    return bottle.static_file("index.html", root=STATIC_DIR)


@app.route("/static/<file:path>")
def assets(file):
    return bottle.static_file(file, root=STATIC_DIR)


@app.route("/app-icon.ico")
def app_icon():
    return bottle.static_file("app_icon.ico", root=os.path.dirname(WEBAPP_DIR))


def _init_dpi_awareness() -> None:
    """在读取显示器工作区前声明系统级 DPI 感知。

    pywebview 在 webview.start() 创建 WinForms 窗口时才调用 SetProcessDPIAware；
    若启动前读取几何，进程仍处于 DPI 虚拟化状态，150% DPI 下 Win32 会返回
    96 DPI 的“虚拟”像素，恢复出来的窗口会被再放大 1.5 倍。这里提前做同样的
    系统级声明（幂等，与 pywebview 的 DPI 口径一致）。
    """
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        get_logger().exception("设置进程 DPI 感知失败，窗口几何将按 96 DPI 处理")


def _startup_geometry(saved: str):
    """解析保存的窗口几何，返回 pywebview 需要的逻辑像素 (宽, 高, x, y)。

    复用 window_geometry：物理工作区按 DPI 换算后校验可见性，记录落在可见
    工作区内则恢复并夹回边界；无记录 / 多屏断开 / 越界 / 尺寸过小时回退为
    主屏居中的默认尺寸。异常时返回 None，由 pywebview 使用默认居中。
    """
    try:
        work_areas = window_geometry.list_work_areas()
        scale = window_geometry.system_dpi_scale()
        return window_geometry.choose_logical_geometry(
            saved, work_areas, scale, DEFAULT_SIZE, MIN_SIZE)
    except Exception:
        get_logger().exception("解析窗口几何失败，使用默认尺寸")
        return None


def main() -> int:
    # 必须在创建 WinForms/WebView2 窗口前设置，任务栏才不会并入 pythonw.exe。
    set_process_app_user_model_id()
    _init_dpi_awareness()
    if os.environ.get("WV_SELFTEST") == "1":
        os.environ["APPDATA"] = tempfile.mkdtemp(prefix="pdf-duplex-selftest-")
    bridge = AppBridge()
    title = "PDF 批量双面打印工具"
    if os.environ.get("WV_SELFTEST") == "1":
        title += " [自检]"
    width, height = DEFAULT_SIZE
    x = y = None
    geometry = _startup_geometry(str(bridge.vars.get("window_geometry") or ""))
    if geometry is not None:
        width, height, x, y = geometry
    window = webview.create_window(
        title,
        app,                      # WSGI app：pywebview 自动起本地服务
        js_api=bridge,
        width=width,
        height=height,
        x=x,
        y=y,
        min_size=MIN_SIZE,
    )
    bridge.attach(window)
    # WebView2 (EdgeChromium)：Windows 默认引擎，GPU 合成渲染
    # 同时在 WinForms 初始化阶段设置 Form.Icon；bridge 再补齐窗口类/小图标。
    webview.start(gui="edgechromium", debug=os.environ.get("WV_DEBUG") == "1",
                  icon=APP_ICON)
    return 0


if __name__ == "__main__":
    sys.exit(main())
