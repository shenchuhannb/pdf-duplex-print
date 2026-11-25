# -*- coding: utf-8 -*-
"""从真实 Explorer 窗口鼠标拖放 PDF 到 WebView 文件列表。"""
import os
import subprocess
import sys
import tempfile
import time

import fitz
from pywinauto import Desktop, mouse

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TITLE = "PDF 批量双面打印工具"
OUT = os.path.join(ROOT, "docs", "screenshots", "webapp_real_drag.png")


def wait_new_window(before, title=None, class_name=None, timeout=15):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        windows = Desktop(backend="uia").windows(title=title, class_name=class_name)
        for window in windows:
            if window.element_info.handle not in before:
                return window
        time.sleep(0.1)
    raise RuntimeError(f"等待窗口超时: title={title!r} class={class_name!r}")


def handles(title=None, class_name=None):
    return {w.element_info.handle for w in Desktop(backend="uia").windows(
        title=title, class_name=class_name)}


def main():
    temp_dir = tempfile.mkdtemp(prefix="pdf-real-drag-")
    appdata = tempfile.mkdtemp(prefix="pdf-real-drag-settings-")
    pdf_path = os.path.join(temp_dir, "drag-sample.pdf")
    doc = fitz.open()
    page = doc.new_page(width=400, height=300)
    page.insert_text((40, 80), "REAL_MOUSE_DRAG_SUCCESS", fontsize=18)
    doc.save(pdf_path)
    doc.close()

    env = dict(os.environ, APPDATA=appdata)
    before_app = handles(title=TITLE)
    app_process = subprocess.Popen(
        [os.path.join(ROOT, ".venv", "Scripts", "pythonw.exe"), "webapp/main.py"],
        cwd=ROOT, env=env)
    app = wait_new_window(before_app, title=TITLE)
    app.set_focus()
    time.sleep(1.5)

    before_explorer = handles(class_name="CabinetWClass")
    explorer_process = subprocess.Popen(["explorer.exe", temp_dir])
    explorer = wait_new_window(before_explorer, class_name="CabinetWClass")
    explorer.set_focus()
    # 两个窗口必须不重叠，否则目标坐标会落回 Explorer，而不是应用列表。
    import win32api
    import win32gui
    screen_w, screen_h = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
    app_w = min(1536, int(screen_w * 0.60))
    app_h = min(1020, int(screen_h * 0.85))
    win32gui.MoveWindow(app.element_info.handle, 0, 0, app_w, app_h, True)
    explorer_x = app_w + 20
    win32gui.MoveWindow(explorer.element_info.handle, explorer_x, 40,
                        screen_w - explorer_x - 20, min(720, screen_h - 100), True)
    explorer.set_focus()
    time.sleep(1.5)
    print("Explorer title:", explorer.window_text(), flush=True)

    candidates = explorer.descendants(title="drag-sample.pdf")
    if not candidates:
        candidates = [w for w in explorer.descendants()
                      if "drag-sample" in (w.window_text() or "")]
    print("File candidates:", [(w.element_info.control_type, w.window_text(),
                               str(w.rectangle())) for w in candidates], flush=True)
    if not candidates:
        raise RuntimeError("资源管理器中未找到测试文件")
    list_items = [w for w in candidates if w.element_info.control_type == "ListItem"]
    item = (list_items or candidates)[0]
    source = item.rectangle().mid_point()
    target_rect = app.rectangle()
    target = (int(target_rect.left + target_rect.width() * 0.23),
              int(target_rect.top + target_rect.height() * 0.43))

    print("Drag:", (source.x, source.y), "->", target, flush=True)
    mouse.move(coords=(source.x, source.y))
    time.sleep(0.2)
    mouse.press(button="left", coords=(source.x, source.y))
    time.sleep(0.4)
    steps = 30
    for index in range(1, steps + 1):
        x = int(source.x + (target[0] - source.x) * index / steps)
        y = int(source.y + (target[1] - source.y) * index / steps)
        mouse.move(coords=(x, y))
        time.sleep(0.03)
    time.sleep(0.4)
    mouse.release(button="left", coords=target)
    time.sleep(3.0)

    app.set_focus()
    app.capture_as_image().save(OUT)
    # WebView2 的 UIA 文本树在部分运行时为空白，不能把“读不到文本”误判为拖放失败。
    # 后端最近记录由 add_files 成功后落盘，用它确认真实路径确实进入应用。
    import json
    settings_file = os.path.join(appdata, "PDFDuplexPrint", "settings.json")
    recent = []
    if os.path.isfile(settings_file):
        with open(settings_file, "r", encoding="utf-8") as handle:
            recent = json.load(handle).get("recent_files", [])
    success = pdf_path in recent
    print("DRAG_BACKEND_SUCCESS:", success, flush=True)
    print("SCREENSHOT:", OUT, flush=True)
    print("RECENT_FILES:", recent, flush=True)

    try:
        explorer.close()
    except Exception:
        pass
    try:
        app.close()
    except Exception:
        pass
    app_process.wait(timeout=8)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
