# -*- coding: utf-8 -*-
"""测量 WebView 原生窗口、页面、桥接和可交互首屏的实际启动耗时。"""
import os
import sys
import threading
import time

START = time.perf_counter()


def log(message):
    print(f"{time.perf_counter() - START:.3f}s {message}", flush=True)


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import webview  # noqa: E402
from app_identity import set_process_app_user_model_id  # noqa: E402
from webapp.main import APP_ICON, DEFAULT_SIZE, MIN_SIZE, app  # noqa: E402
from webapp.bridge import AppBridge  # noqa: E402

log("imports ready")
set_process_app_user_model_id()
bridge = AppBridge()
window = webview.create_window(
    "PDF 批量双面打印工具 [启动测速]", app, js_api=bridge,
    width=DEFAULT_SIZE[0], height=DEFAULT_SIZE[1], min_size=MIN_SIZE)


def interactive_probe():
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        try:
            ready = window.evaluate_js(
                "!!(document.getElementById('btn-add') && "
                "document.getElementById('btn-add').onclick)")
            if ready:
                log("interactive")
                window.destroy()
                return
        except Exception:
            pass
        time.sleep(0.05)
    log("interactive timeout")
    window.destroy()


for name in ("shown", "loaded", "_pywebviewready"):
    event = getattr(window.events, name)
    event += (lambda event_name: lambda *args: log(event_name))(name)

bridge.attach(window)
window.events._pywebviewready += lambda *_args: threading.Thread(
    target=interactive_probe, daemon=True).start()
log("webview.start")
webview.start(gui="edgechromium", icon=APP_ICON)
log("closed")
