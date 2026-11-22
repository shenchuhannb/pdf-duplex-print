# -*- coding: utf-8 -*-
"""pywebview 版入口：UI 用 WebView2（GPU 合成，resize 丝滑），后端 100% 复用。

路由方案：自建 bottle WSGI app 精确控制静态资源根目录，
绕开 pywebview 相对路径的 commonpath 解析（实测在子目录布局下 404）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bottle
import webview

from webapp.bridge import AppBridge

MIN_SIZE = (1024, 680)
DEFAULT_SIZE = (1280, 860)
WEBAPP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(WEBAPP_DIR, "static")

app = bottle.Bottle()


@app.route("/")
def index():
    return bottle.static_file("index.html", root=STATIC_DIR)


@app.route("/static/<file:path>")
def assets(file):
    return bottle.static_file(file, root=STATIC_DIR)


def main() -> int:
    bridge = AppBridge()
    window = webview.create_window(
        "PDF 批量双面打印工具",
        app,                      # WSGI app：pywebview 自动起本地服务
        js_api=bridge,
        width=DEFAULT_SIZE[0],
        height=DEFAULT_SIZE[1],
        min_size=MIN_SIZE,
    )
    bridge.attach(window)
    # WebView2 (EdgeChromium)：Windows 默认引擎，GPU 合成渲染
    webview.start(gui="edgechromium", debug=os.environ.get("WV_DEBUG") == "1")
    return 0


if __name__ == "__main__":
    sys.exit(main())
