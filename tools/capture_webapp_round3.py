# -*- coding: utf-8 -*-
r"""第三轮反馈的真实窗口截图：1024 逻辑宽度空状态 + 横版 2/4 合 1 预览。

运行：.venv\Scripts\python.exe tools\capture_webapp_round3.py

约束：
  * 全部操作在应用进程内通过 evaluate_js 完成，不调用任何真实打印。
  * 使用隔离的 APPDATA，不污染用户配置，也不会留下“最近文件”。
  * 截图用 Win32 PrintWindow(PW_RENDERFULLCONTENT)；WebView2 偶尔返回空白时，
    回退到前台 BitBlt（窗口已在桌面可见，属真实窗口截图）。
"""
import ctypes
import json
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

APPDATA = tempfile.mkdtemp(prefix="pdf-round3-shot-")
os.environ["APPDATA"] = APPDATA

import fitz  # noqa: E402
import webview  # noqa: E402
from PIL import Image  # noqa: E402

import bottle  # noqa: E402
import win32con  # noqa: E402
import win32gui  # noqa: E402
import win32ui  # noqa: E402

from webapp.bridge import AppBridge  # noqa: E402

WEBAPP_DIR = os.path.join(ROOT, "webapp")
STATIC_DIR = os.path.join(WEBAPP_DIR, "static")
SHOTS = os.path.join(ROOT, "docs", "screenshots")
SAMPLE_DIR = tempfile.mkdtemp(prefix="pdf-round3-src-")

app = bottle.Bottle()


@app.route("/")
def index():
    return bottle.static_file("index.html", root=STATIC_DIR)


@app.route("/static/<file:path>")
def assets(file):
    return bottle.static_file(file, root=STATIC_DIR)


@app.route("/app-icon.ico")
def app_icon():
    return bottle.static_file("app_icon.ico", root=ROOT)


def make_sample(path, pages, title):
    doc = fitz.open()
    font_file = r"C:\Windows\Fonts\msyh.ttc"
    font_name = "helv"
    for page_no in range(1, pages + 1):
        page = doc.new_page(width=595, height=842)
        try:
            if os.path.isfile(font_file):
                page.insert_font(fontname="msyh", fontfile=font_file)
                font_name = "msyh"
        except Exception:
            font_name = "helv"
        page.draw_rect(fitz.Rect(56, 56, 539, 786), color=(0.68, 0.72, 0.78), width=1.2)
        page.insert_text((80, 130), title, fontsize=24, fontname=font_name)
        page.insert_text((80, 200), f"第 {page_no} 页 / 共 {pages} 页", fontsize=34,
                         fontname=font_name)
        page.insert_text((80, 260), "横版 2 合 1 / 4 合 1 预览验证页", fontsize=14,
                         fontname=font_name, color=(0.33, 0.36, 0.42))
        for row in range(6):
            y = 330 + row * 46
            page.draw_line((80, y), (515, y), color=(0.82, 0.85, 0.89), width=0.8)
        page.insert_text((80, 700), "PDF-DUPLEX-ROUND3", fontsize=12,
                         fontname=font_name, color=(0.45, 0.48, 0.52))
    doc.save(path)
    doc.close()
    return path


def grab(hwnd, out_path):
    """PrintWindow 抓真实窗口；空白时回退 BitBlt（窗口已在桌面）。"""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)

    from ctypes import wintypes
    user32 = ctypes.windll.user32
    user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
    user32.PrintWindow.restype = wintypes.BOOL
    ok = user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
    bmpinfo = bitmap.GetInfo()
    raw = bitmap.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                           raw, "raw", "BGRX", 0, 1)
    content = img.convert("L").getextrema()[1] - img.convert("L").getextrema()[0]
    if not ok or content < 12:
        # 回退：把窗口置前，再从屏幕 DC 拷贝（仍是真实窗口内容）。
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        screen_dc = win32gui.GetWindowDC(hwnd)
        src = win32ui.CreateDCFromHandle(screen_dc)
        save_dc.BitBlt((0, 0), (width, height), src, (0, 0), win32con.SRCCOPY)
        raw = bitmap.GetBitmapBits(True)
        img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                               raw, "raw", "BGRX", 0, 1)
        src.DeleteDC()
        win32gui.ReleaseDC(hwnd, screen_dc)
        print("CAPTURE_FALLBACK_BITBLT:", out_path, flush=True)

    img.save(out_path)
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    win32gui.DeleteObject(bitmap.GetHandle())
    print("SHOT:", out_path, img.size, flush=True)


def main() -> int:
    pdf_a = make_sample(os.path.join(SAMPLE_DIR, "横版示例-A.pdf"), 6, "季度报告 A")
    pdf_b = make_sample(os.path.join(SAMPLE_DIR, "横版示例-B.pdf"), 4, "附录说明 B")
    bridge = AppBridge()
    window = webview.create_window("PDF 批量双面打印工具", app, js_api=bridge,
                                   width=1024, height=680, min_size=(1024, 680))
    bridge.attach(window)

    state = {"iid": None}

    def js(code):
        return window.evaluate_js(code)

    def wait_js(code, timeout=20.0):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                if js(code):
                    return True
            except Exception:
                pass
            time.sleep(0.15)
        return False

    def shot(name):
        time.sleep(0.6)
        hwnd = int(window.native.Handle.ToInt64())
        grab(hwnd, os.path.join(SHOTS, f"round3_{name}.png"))

    def wait_preview(timeout=45.0):
        """等到预览图真正回填、状态行不再是“正在加载”，避免截到 loading 覆盖层。"""
        return wait_js(
            "(()=>{const img=document.querySelector('#preview-stage img');"
            "const st=document.getElementById('preview-status').textContent;"
            "return !!img && img.complete && img.naturalWidth>0"
            " && !st.includes('正在') && !document.querySelector("
            "'#preview-stage .preview-empty')})()", timeout=timeout)

    def worker():
        try:
            ok = wait_js("!!(window.pywebview && window.pywebview.api && "
                         "document.getElementById('btn-add'))", timeout=30)
            print("BRIDGE_READY:", ok, flush=True)
            window.resize(1024, 680)
            time.sleep(1.2)
            js("window.__pushLog(new Date().toTimeString().slice(0,8), '就绪，可以添加文件。')")
            empty_metrics = js(
                "JSON.stringify((()=>{const b=document.getElementById('btn-empty-add');"
                "const r=b.getBoundingClientRect();"
                "const list=document.getElementById('file-list').getBoundingClientRect();"
                "const bar=document.getElementById('selection-bar');"
                "const box=document.getElementById('empty-state').getBoundingClientRect();"
                "const visible=Math.max(0,Math.min(r.bottom,list.bottom)"
                "-Math.max(r.top,list.top));"
                "const hit=document.elementFromPoint(r.left+r.width/2,r.top+r.height/2);"
                "return {vw:window.innerWidth,vh:window.innerHeight,"
                "button:[r.left,r.top,r.right,r.bottom],list:[list.top,list.bottom],"
                "emptyBox:[box.top,box.bottom],visiblePx:visible,"
                "visibleRatio:r.height>0?visible/r.height:0,"
                "barHidden:bar.hidden,centerHit:hit?hit.tagName+(hit.id?'#'+hit.id:''):null"
                ",hitIsButton:hit===b||b.contains(hit),"
                "emptyScrolled:document.getElementById('empty-state').scrollHeight"
                ">document.getElementById('empty-state').clientHeight+1}})())")
            print("EMPTY_BUTTON:", empty_metrics, flush=True)
            shot("empty_1024")

            js(f"void doAdd({json.dumps([pdf_a, pdf_b])})")
            ready = wait_js("state.files.length === 2 && state.selected.size === 2 && "
                            "!!document.querySelector('#preview-stage img')", timeout=40)
            print("FILES_READY:", ready, flush=True)
            time.sleep(0.8)
            shot("added_1024")

            js("document.getElementById('preview-orientation').value = 'landscape';"
               "document.getElementById('preview-nup').value = '2';"
               "document.querySelector('input[name=preview-view][value=single]').click();"
               "onSettingChange()")
            landscape2 = wait_preview()
            print("LANDSCAPE_NUP2:", landscape2, flush=True)
            js("selectPreviewFile(state.files[1])")   # 切到最后一个文件（4 页）
            wait_preview()
            time.sleep(0.8)
            shot("landscape_nup2_1024")

            js("document.getElementById('preview-nup').value = '4'; onSettingChange()")
            landscape4 = wait_preview()
            print("LANDSCAPE_NUP4:", landscape4, flush=True)
            time.sleep(0.8)
            shot("landscape_nup4_1024")
            print("SHOT_METRICS:", json.dumps({
                "files": js("state.files.length"),
                "selected": js("state.selected.size"),
                "status": js("document.getElementById('preview-status').textContent"),
                "paper": js("document.getElementById('preview-status').textContent"),
                "printerWidth": js("Math.round(document.getElementById("
                                   "'printer-select').getBoundingClientRect().width)"),
                "selRow": js("(()=>{const bs=[...document.querySelectorAll("
                             "'#selection-bar .selection-buttons .btn')]"
                             ".map(b=>b.getBoundingClientRect());"
                             "return bs.map(r=>[r.left,r.top,r.right,r.bottom])})()"),
            }, ensure_ascii=False), flush=True)
        except Exception as exc:  # noqa: BLE001
            import traceback
            print("SHOT_ERROR:\n" + traceback.format_exc(), flush=True)
        finally:
            try:
                window.destroy()
            except Exception:
                pass

    webview.start(gui="edgechromium", func=worker)
    return 0


if __name__ == "__main__":
    sys.exit(main())