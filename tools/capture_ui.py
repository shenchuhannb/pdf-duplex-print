# -*- coding: utf-8 -*-
"""离屏生成风格 A 的最终截图：docs/ui_screenshot.png（主界面）与 docs/ui_advanced.png（高级设置）。

运行：python tools\\capture_ui.py

★ 硬性约束（后续修改必须保留）：
  1. 所有截图必须离屏完成，绝对不能在用户桌面弹出任何窗口（包括一闪而过的
     Tk 窗口 / messagebox / 翻面提示 / 文件对话框）。
  2. 禁止使用 PIL.ImageGrab 抓屏——它依赖屏幕可见像素，会强制窗口上屏。
     本脚本让窗口保持正常坐标但整体 alpha=0（完全透明），DWM 仍会合成内容，
     再用 Win32 PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT) 抓内存位图。
     单纯把窗口移到屏幕外时 DWM 不合成 ttk 子控件，截图会缺失页签/边框。
  3. 本脚本不触发任何真实打印/交互；高级设置与翻面确认框同样以透明窗口
     离屏渲染，翻面框抓帧后自动发出 Esc（等价于取消），不会继续第二阶段。
"""

import ctypes
import os
import shutil
import sys
import tempfile
import time
import tkinter as tk
import tkinter.font as tkfont

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fitz
from PIL import Image

import win32gui
import win32ui

TEMP_ROOT = tempfile.mkdtemp(prefix="pdf_duplex_capture_")
os.environ["APPDATA"] = TEMP_ROOT  # 隔离真实配置

from ui_main import DuplexPrintApp, _create_root  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")

_print_window = ctypes.windll.user32.PrintWindow
_print_window.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
_print_window.restype = ctypes.c_int


def _build_pdf(path: str, pages: int, title: str) -> str:
    doc = fitz.open()
    font_name = "helv"
    font_file = r"C:\Windows\Fonts\msyh.ttc"
    for index in range(pages):
        page = doc.new_page(width=595, height=842)
        try:
            if os.path.isfile(font_file):
                page.insert_font(fontname="msyh", fontfile=font_file)
                font_name = "msyh"
        except Exception:
            font_name = "helv"
        page.insert_text((56, 90), title, fontsize=22, fontname=font_name)
        page.insert_text((56, 150), f"第 {index + 1} 页", fontsize=36, fontname=font_name)
        page.draw_rect(fitz.Rect(56, 200, 539, 720), color=(0.75, 0.78, 0.82), width=1)
        page.insert_text((80, 250), "示例正文：用于演示打印预览的成纸效果。",
                         fontsize=12, fontname=font_name,
                         color=(0.35, 0.38, 0.42))
    doc.save(path)
    doc.close()
    return path


def _root_hwnd(window) -> int:
    """取顶层窗口句柄（winfo_id 返回的是子窗口）。"""
    hwnd = int(window.winfo_id())
    return win32gui.GetAncestor(hwnd, 0x0002)  # GA_ROOT


def _grab(hwnd, out_path: str, flag: int, checker=None) -> bool:
    """用 PrintWindow 把窗口内容抓到内存位图并保存 PNG。返回是否抓到内容。"""
    checker = checker or _is_contentful
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = max(1, right - left), max(1, bottom - top)
    window_dc = win32gui.GetWindowDC(hwnd)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    memory_dc.SelectObject(bitmap)
    bmp_path = out_path + f".{flag}.bmp"
    try:
        ok = _print_window(hwnd, memory_dc.GetSafeHdc(), flag)
        bitmap.SaveBitmapFile(memory_dc, bmp_path)
        image = _crop_client(hwnd, Image.open(bmp_path).convert("RGB"))
        if not checker(image):
            return False
        image.save(out_path, optimize=True)
        return True
    finally:
        try:
            os.remove(bmp_path)
        except OSError:
            pass
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, window_dc)


def _crop_client(hwnd, image: Image.Image) -> Image.Image:
    """裁掉标题栏/边框，只保留客户区（让截图尺寸等于设置的窗口内容尺寸）。"""
    left, top, _, _ = win32gui.GetWindowRect(hwnd)
    client_w, client_h = win32gui.GetClientRect(hwnd)[2:]
    origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
    offset_x, offset_y = origin_x - left, origin_y - top
    box = (offset_x, offset_y, offset_x + client_w, offset_y + client_h)
    if box[2] <= image.size[0] and box[3] <= image.size[1] and client_w > 0:
        return image.crop(box)
    return image


def _is_contentful(image: Image.Image) -> bool:
    """校验截图非空：尺寸合理、有明显明暗层次、颜色数量足够。"""
    if image.size[0] < 600 or image.size[1] < 500:
        return False
    gray = image.convert("L")
    low, high = gray.getextrema()
    if high - low < 30:
        return False
    colors = image.getcolors(maxcolors=400_000)
    return bool(colors) and len(colors) > 200


def _is_contentful_dialog(image: Image.Image) -> bool:
    """对话框较小，放宽尺寸阈值，但仍要求有明暗层次与足够颜色。"""
    if image.size[0] < 240 or image.size[1] < 120:
        return False
    gray = image.convert("L")
    low, high = gray.getextrema()
    if high - low < 30:
        return False
    colors = image.getcolors(maxcolors=200_000)
    return bool(colors) and len(colors) > 60


def _capture_window(hwnd, out_path: str, checker=None) -> Image.Image:
    """优先 PW_RENDERFULLCONTENT（DWM 合成内容），失败时退回 WM_PRINT。"""
    for flag in (2, 0):
        if _grab(hwnd, out_path, flag, checker):
            image = Image.open(out_path)
            assert (checker or _is_contentful)(image), "截图内容疑似空白"
            return image
    raise RuntimeError("PrintWindow 未能取到窗口内容")


def _pump(root, seconds: float) -> None:
    """在离屏状态下推进 Tk 事件循环（窗口 alpha=0，用户不可见）。"""
    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update()
        time.sleep(0.05)


def _wait_preview_ready(root, app, timeout: float = 15.0) -> bool:
    """等待后台成纸预览完成（离屏时同样需要驱动队列轮询）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if app.preview_pngs and app.preview_metrics_data:
            return True
        time.sleep(0.1)
    return False


def _add_demo_logs(app) -> None:
    """填 6 条示例日志，证明常驻日志区在最小窗口下仍能显示 5~6 行。"""
    for line in (
        "已添加 2 个文件：产品手册.pdf、会议纪要.pdf。",
        "打印流程：整批翻面 · 打完全部再取纸；正面 1,3,5 → 翻面一次 → 反面 6,4,2。",
        "临时合并文件已清理，输出目录：打印输出。",
        "成纸预览已更新（当前设置：A4 · 1 合 1 · 收缩到可打印区）。",
        "提示：可多选后批量设置范围，也可双击单行“本次范围”修改。",
        "已为 2 个文件设置范围 1-2、份数 2。",
    ):
        app._append_log(line)


def _log_metrics(root, app) -> str:
    """打印日志区/上方区域的实测像素，供汇报引用。"""
    line_space = tkfont.Font(root=root, font=app.log_text.cget("font")).metrics("linespace")
    lines = app.log_text.winfo_height() / float(max(1, line_space))
    return ("日志区 %dpx / 正文 %dpx / 行高 %dpx → %.1f 行；"
            "文件列表 %dpx / 预览画布 %dpx / 工作区 %dpx"
            % (app.log_drawer.winfo_height(), app.log_text.winfo_height(), line_space, lines,
               app.tree.winfo_height(), app.preview_canvas.winfo_height(),
               app.workspace.winfo_height()))


def _make_app(pdfs, scheme: str = "", size=(1280, 960), scaling: float = 0.0):
    """创建完全透明的窗口并加载示例文件（全程不上屏）。

    scaling：模拟高 DPI 的 tk scaling（96dpi=1.333，150%=2.0）。
    """
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    root = _create_root()
    root.withdraw()                      # 先隐藏，构建期间绝不上屏
    # 截图统一按 100% 逻辑尺寸渲染：显式设定 tk scaling，并按同一比例换算
    # 像素间距，避免本机 150% DPI 把 1100×720 窗口挤到不可用。
    scale = scaling or 1.3333
    root.tk.call("tk", "scaling", scale)
    ui_scale = scale / 1.3333
    # Tk 命名字体是进程级状态：按当前档位显式重设，避免上一张图（150%）
    # 留下的 13pt 字号污染后面 100% 的截图（Treeview 行高、日志行高会失真）。
    for name, font_size in (("TkDefaultFont", 9), ("TkTextFont", 9), ("TkMenuFont", 9),
                            ("TkHeadingFont", 9), ("TkCaptionFont", 9),
                            ("TkTooltipFont", 8), ("TkSmallCaptionFont", 8),
                            ("TkIconFont", 9), ("TkFixedFont", 9)):
        try:
            tkfont.nametofont(name, root=root).configure(
                size=max(6, int(round(font_size * ui_scale))))
        except tk.TclError:
            pass
    app = DuplexPrintApp(root, style="A", layout_scheme=scheme,
                         ui_scale_override=ui_scale)
    if pdfs:
        app._add_paths(pdfs)
    root.geometry(f"{size[0]}x{size[1]}+40+40")
    root.attributes("-alpha", 0.0)        # 完全透明
    root.update_idletasks()
    root.deiconify()                     # 映射窗口才能被 PrintWindow 渲染
    root.update()
    return root, app


def _teardown(root, app) -> None:
    """清理：置为退出态并取消所有 after 轮询，避免 Tcl 残留回调警告。"""
    app.shutting_down = True
    app.preview_enabled = False
    for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
        poll_id = getattr(app, attr, None)
        if poll_id is not None:
            try:
                root.after_cancel(poll_id)
            except Exception:
                pass
            setattr(app, attr, None)
    if app.preview_temp_dir:
        shutil.rmtree(app.preview_temp_dir, ignore_errors=True)
    try:
        root.update()
    except Exception:
        pass
    root.destroy()


def capture_main(pdfs, out_path: str, scheme: str = "", size=(1280, 960),
                 scaling: float = 0.0, wait_preview: bool = True):
    """主界面截图：文件列表 + 成纸预览 + 主操作（含纸张轮廓/可打印区/实际百分比）。"""
    root, app = _make_app(pdfs, scheme=scheme, size=size, scaling=scaling)
    # 展示核心能力：手动双面 + 流程 1（整批翻面 · 打完全部再取纸）
    app.vars["manual_duplex"].set(True)
    app.duplex_mode_var.set("duplex")
    app.vars["print_flow"].set("batch_all")
    app.vars["copies"].set("1")
    app.vars["copy_exec_on"].set(False)
    app._sync_widgets()
    _pump(root, 0.4)
    if wait_preview:
        if not _wait_preview_ready(root, app):
            raise RuntimeError("成纸预览未在限定时间内完成，终止截图")
        app._show_preview_page(0)
        _pump(root, 0.4)
    if app.files:
        _add_demo_logs(app)
    _pump(root, 0.2)
    print("   " + _log_metrics(root, app))

    if float(root.attributes("-alpha")) > 0.0:
        raise RuntimeError("窗口未处于透明状态，已中止截图")
    image = _capture_window(_root_hwnd(root), out_path)
    _teardown(root, app)
    return image.size


def capture_flow_options(pdfs, docs_root: str) -> None:
    """三种打印流程各自选中时的设置条裁图（flow_option_1/2/3.png）。

    copies=1 + 手动双面：三个选项全部可用，便于逐字核对文案与排版。
    """
    for index, flow in enumerate(("batch_all", "batch_take", "per_file"), start=1):
        out_path = os.path.join(docs_root, f"flow_option_{index}.png")
        root, app = _make_app(list(pdfs), size=(1280, 960))
        app.vars["manual_duplex"].set(True)
        app.duplex_mode_var.set("duplex")
        app.vars["copies"].set("1")
        app.vars["copy_exec_on"].set(False)
        app.vars["print_flow"].set(flow)
        app._sync_widgets()
        _pump(root, 0.4)
        strip = app.printer_combo.master.master
        hwnd = _root_hwnd(root)
        full_path = out_path + ".full.png"
        full = _capture_window(hwnd, full_path)
        origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
        x = strip.winfo_rootx() - origin_x
        y = strip.winfo_rooty() - origin_y
        margin = 8
        box = (max(0, x - margin), max(0, y - margin),
               min(full.width, x + strip.winfo_width() + margin),
               min(full.height, y + strip.winfo_height() + margin))
        crop = full.crop(box)
        crop.save(out_path, optimize=True)
        try:
            os.remove(full_path)
        except OSError:
            pass
        print(f"流程 {index}（{flow}）设置条 ->", out_path, crop.size)
        _teardown(root, app)


def capture_advanced(out_path: str, pdfs=()):
    """高级设置对话框截图：以透明 Toplevel 离屏渲染（不弹真实窗口）。"""
    root, app = _make_app(list(pdfs))
    if app.files:
        app.tree.selection_set(app.files[0]["iid"])
        app._sync_widgets()
    # visible=False：构建完成但不上屏；随后设透明再由本脚本 deiconify 抓帧
    app.open_advanced_settings(visible=False)
    win = app.advanced_window
    # 先隐藏再设置透明与几何（避免映射瞬间上屏，也避免几何请求被撤销）
    win.withdraw()
    win.attributes("-alpha", 0.0)
    win.geometry("1200x1000+90+30")
    win.deiconify()
    root.update()
    _pump(root, 0.6)
    if float(win.attributes("-alpha")) > 0.0:
        raise RuntimeError("高级设置窗口未处于透明状态，已中止截图")
    image = _capture_window(_root_hwnd(win), out_path)
    app._close_advanced_settings()
    _teardown(root, app)
    return image.size


def capture_per_file(out_path: str, pdfs):
    """逐文件打印范围截图：按上方范围 / 单独设置 / 非法范围（红色标行）。"""
    root, app = _make_app(list(pdfs))
    app.vars["per_file"].set(True)
    app.vars["range_mode"].set("all")
    app.vars["copies"].set("2")
    if len(app.files) >= 3:
        app.files[1]["range_text"] = "1-2"
        app.files[1]["copies"] = "3"
        app.files[2]["range_text"] = "abc"
    app.tree.selection_set(app.files[0]["iid"])
    app._sync_widgets()
    _pump(root, 0.4)
    if not _wait_preview_ready(root, app):
        raise RuntimeError("成纸预览未在限定时间内完成，终止截图")
    app._show_preview_page(0)
    _pump(root, 0.4)
    if float(root.attributes("-alpha")) > 0.0:
        raise RuntimeError("窗口未处于透明状态，已中止截图")
    image = _capture_window(_root_hwnd(root), out_path)
    _teardown(root, app)
    return image.size


def capture_selection_bar(out_path: str, pdfs):
    """逐文件多种选择截图：多选 2 个文件，操作条显示“已选 2 个文件”。"""
    root, app = _make_app(list(pdfs))
    app.vars["per_file"].set(True)
    app.vars["range_mode"].set("all")
    app.vars["copies"].set("2")
    if len(app.files) >= 3:
        app.files[1]["range_text"] = "2-4"
        app.files[1]["copies"] = "2"
        app.files[2]["range_text"] = "abc"
    selected = [record["iid"] for record in app.files[:2]]
    app.tree.selection_set(selected)
    app._on_tree_select()
    app._sync_widgets()
    _pump(root, 0.4)
    if not _wait_preview_ready(root, app):
        raise RuntimeError("成纸预览未在限定时间内完成，终止截图")
    app._show_preview_page(0)
    _add_demo_logs(app)
    _pump(root, 0.3)
    print("   " + _log_metrics(root, app), "| 操作条：", app.sel_info_var.get())
    if float(root.attributes("-alpha")) > 0.0:
        raise RuntimeError("窗口未处于透明状态，已中止截图")
    image = _capture_window(_root_hwnd(root), out_path)
    _teardown(root, app)
    return image.size


def capture_flip_modal(out_path: str, pdfs):
    """翻面确认模态框截图：透明离屏渲染，抓帧后发 Esc 取消（不进入第二阶段）。"""
    root, app = _make_app(list(pdfs))
    app.offscreen_dialog_alpha = 0.0
    detail = [
        "阶段：整批阶段一（全部正面）已完成",
        "文件：《产品手册》、《会议纪要》",
        "张数：4 张纸（8 个输出面）",
        "翻面方向：整叠纸【整体翻面】，保持上下顺序不变（通常为长边翻转）",
        "正面页序：\n《产品手册》：1, 3\n《会议纪要》：1, 3, 5, 7",
        "反面页序：8, 6, 4, 2（整批全局逆序）",
    ]
    sample = "示例：正面 1,3,5 … → 整叠翻面 → 反面 6,4,2 …（每文件已补空白页，不跨文件错位）"
    state = {"captured": False}

    def grab_and_cancel() -> None:
        dialogs = [w for w in root.winfo_children()
                   if isinstance(w, tk.Toplevel) and w.winfo_viewable()]
        if not dialogs:
            return
        win = dialogs[-1]
        root.update_idletasks()
        _capture_window(_root_hwnd(win), out_path, checker=_is_contentful_dialog)
        state["captured"] = True
        win.event_generate("<Escape>")

    # 700ms 后抓帧并取消；该回调在 wait_window 的主循环里执行
    root.after(700, grab_and_cancel)
    confirmed = app._flip_modal("请翻面后继续（整批只翻 1 次）", detail, sample)
    if confirmed:
        raise RuntimeError("翻面框不应被确认（脚本只发送 Esc）")
    if not state["captured"]:
        raise RuntimeError("未抓到翻面确认框")
    _teardown(root, app)
    return Image.open(out_path).size


def _capture_modal(out_path: str, pdfs, trigger):
    """通用离屏模态框截图：trigger 触发 wait_window，700ms 后抓帧并发 Esc。

    trigger 的返回值只用于日志；Esc 一律等价取消，不会继续任何打印阶段。
    """
    root, app = _make_app(list(pdfs))
    app.offscreen_dialog_alpha = 0.0
    state = {"captured": False}

    def grab_and_cancel() -> None:
        dialogs = [w for w in root.winfo_children()
                   if isinstance(w, tk.Toplevel) and w.winfo_viewable()]
        if not dialogs:
            return
        win = dialogs[-1]
        root.update_idletasks()
        _capture_window(_root_hwnd(win), out_path, checker=_is_contentful_dialog)
        state["captured"] = True
        win.event_generate("<Escape>")

    root.after(700, grab_and_cancel)
    result = trigger(app)
    if not state["captured"]:
        raise RuntimeError(f"未抓到对话框：{out_path}")
    _teardown(root, app)
    return Image.open(out_path).size, result


def capture_batch_take(out_path: str, pdfs):
    """整批一次翻面 + 每文件暂停：文件之间“请取出这 N 张纸”确认框。"""
    def trigger(app):
        app._ui_call = lambda fn: fn()
        return app._file_done_modal(
            {"name": "会议纪要.pdf"}, 2, 3, 4, "产品手册.pdf",
            mode="batch", group_done=2, group_total=3)

    return _capture_modal(out_path, pdfs, trigger)


def capture_perfile_flip(out_path: str, pdfs):
    """逐文件翻面：每个文件第一面打完后“请翻面”确认框。"""
    def trigger(app):
        app._ui_call = lambda fn: fn()
        return app._ask_flip_file({"name": "产品手册.pdf"}, 4, [4, 2], 1, 3)

    return _capture_modal(out_path, pdfs, trigger)


def capture_perfile_take(out_path: str, pdfs):
    """逐文件翻面：每个文件第二面打完后“请取出该文件”确认框。"""
    def trigger(app):
        return app._file_done_modal(
            {"name": "产品手册.pdf"}, 1, 3, 4, "会议纪要.pdf",
            mode="by_file", final=False)

    return _capture_modal(out_path, pdfs, trigger)


def capture_copies_spin(out_path: str, pdfs=(), scaling: float = 2.0,
                        size=(1280, 960)):
    """裁出“份数”Spinbox 控件区域（150% DPI 最苛刻），并校验 4 位数字放得下。

    断言：font.measure("9999") ≤ 输入框可用文本宽度；
          控件高度 ≥ 字体行高 + 上下内边距。
    """
    import re
    root, app = _make_app(list(pdfs), size=size, scaling=scaling)
    app.vars["copies"].set("9999")
    _pump(root, 0.4)
    spin = app.copies_spin
    style_font = root.tk.call("ttk::style", "lookup", "Copies.TSpinbox", "-font")
    font = tkfont.Font(root=root, font=style_font or "TkDefaultFont")
    padding = root.tk.call("ttk::style", "lookup", "Copies.TSpinbox", "-padding")
    nums = [int(v) for v in re.findall(r"\d+", str(padding))]
    left_pad, top_pad = (nums[0] if nums else 0), (nums[-1] if nums else 0)
    # ttk Spinbox 的 width 只表示文本字段，上下箭头在字段右侧单独占位；
    # 箭头宽度不随 DPI 放大（实测 100%=18px、150%=18px），不能用 _px() 换算。
    arrow = 18
    available = spin.winfo_width() - 2 * left_pad - arrow
    need = font.measure("9999")
    if available < need:
        raise RuntimeError(f"份数框放不下 4 位数字：可用 {available} < {need}")
    if spin.winfo_height() < font.metrics("linespace") + 2 * top_pad:
        raise RuntimeError("份数框高度小于行高 + 上下内边距")
    print(f"   份数框 {spin.winfo_width()}×{spin.winfo_height()}，padding={padding}，"
          f"行高 {font.metrics('linespace')}，measure('9999')={need}，可用文本宽 {available}")
    if float(root.attributes("-alpha")) > 0.0:
        raise RuntimeError("窗口未处于透明状态，已中止截图")
    hwnd = _root_hwnd(root)
    full_path = out_path + ".full.png"
    full = _capture_window(hwnd, full_path)
    origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
    x = spin.winfo_rootx() - origin_x
    y = spin.winfo_rooty() - origin_y
    margin = 6
    box = (max(0, x - margin), max(0, y - margin),
           x + spin.winfo_width() + margin, y + spin.winfo_height() + margin)
    crop = full.crop(box)
    crop.save(out_path, optimize=True)
    try:
        os.remove(full_path)
    except OSError:
        pass
    _teardown(root, app)
    return crop.size


def _composite_before_after(before: str, after: str, out_path: str) -> None:
    """左右拼接“打磨前 / 打磨后”，顶部标注，便于一眼对比。"""
    from PIL import ImageDraw, ImageFont

    left = Image.open(before).convert("RGB")
    right = Image.open(after).convert("RGB")
    gap, header = 16, 36
    width = left.width + gap + right.width
    height = max(left.height, right.height) + header
    canvas = Image.new("RGB", (width, height), "#eef1f5")
    canvas.paste(left, (0, header))
    canvas.paste(right, (left.width + gap, header))
    draw = ImageDraw.Draw(canvas)
    font_path = r"C:\Windows\Fonts\msyh.ttc"
    try:
        font = ImageFont.truetype(font_path, 16)
    except Exception:
        font = ImageFont.load_default()
    for text, x in (("打磨前（上一轮）", 8), ("打磨后（本轮）", left.width + gap + 8)):
        draw.text((x, 8), text, fill="#3a4a5c", font=font)
    canvas.save(out_path, optimize=True)


def main() -> None:
    os.makedirs(DOCS, exist_ok=True)
    folder = tempfile.mkdtemp(prefix="pdfs_", dir=TEMP_ROOT)
    pdf3 = _build_pdf(os.path.join(folder, "产品手册.pdf"), 3, "产品手册")
    pdf4 = _build_pdf(os.path.join(folder, "会议纪要.pdf"), 4, "会议纪要")
    pdf2 = _build_pdf(os.path.join(folder, "报价单.pdf"), 2, "报价单")

    out_main = os.path.join(DOCS, "ui_screenshot.png")
    print("主界面 ->", out_main, capture_main([pdf3, pdf4], out_main))

    out_adv = os.path.join(DOCS, "ui_advanced.png")
    print("高级设置 ->", out_adv, capture_advanced(out_adv, [pdf3]))

    out_scope = os.path.join(DOCS, "per_file_scope.png")
    print("逐文件范围 ->", out_scope,
          capture_per_file(out_scope, [pdf3, pdf2, pdf4]))

    out_selection = os.path.join(DOCS, "per_file_selection.png")
    print("逐文件多选操作条 ->", out_selection,
          capture_selection_bar(out_selection, [pdf3, pdf2, pdf4]))

    out_flip = os.path.join(DOCS, "flip_modal.png")
    print("翻面确认框 ->", out_flip, capture_flip_modal(out_flip, [pdf3, pdf4]))

    # 本轮：两种模式的停顿点弹窗 + 份数控件裁图
    out_batch_take = os.path.join(DOCS, "flip_modal_batch.png")
    print("整批文件间取纸 ->", out_batch_take,
          capture_batch_take(out_batch_take, [pdf3, pdf4, pdf2]))
    out_pf_flip = os.path.join(DOCS, "flip_modal_perfile_flip.png")
    print("逐文件翻面确认 ->", out_pf_flip,
          capture_perfile_flip(out_pf_flip, [pdf3, pdf4, pdf2]))
    out_pf_take = os.path.join(DOCS, "flip_modal_perfile_take.png")
    print("逐文件取纸确认 ->", out_pf_take,
          capture_perfile_take(out_pf_take, [pdf3, pdf4, pdf2]))
    # 三种打印流程的设置条裁图（逐字核对文案）
    capture_flow_options([pdf3, pdf4, pdf2], DOCS)

    out_spin = os.path.join(DOCS, "fix_copies_spinbox.png")
    print("份数控件裁图（150% DPI）->", out_spin,
          capture_copies_spin(out_spin))

    # 本轮打磨截图（离屏 PrintWindow，不上屏）
    # 1024×720：用户实际窗口内容宽度下限，逐控件无横向裁剪
    out_1024 = os.path.join(DOCS, "ui_1024.png")
    print("1024 主界面 ->", out_1024,
          capture_main([pdf3, pdf4], out_1024, size=(1024, 720)))
    polish_targets = (
        ("polish_1280", [pdf3, pdf4], (1280, 960), 0.0, True),
        ("polish_1100", [pdf3, pdf4], (1100, 720), 0.0, True),
        # 150% 模拟：tk scaling=2.0，物理像素 1920×1440 = 逻辑 1280×960
        ("ui_150pct", [pdf3, pdf4], (1920, 1440), 2.0, True),
        # 空状态：无文件时的引导画面
        ("polish_empty", [], (1280, 960), 0.0, False),
    )
    for name, pdfs, size, scaling, wait_preview in polish_targets:
        out = os.path.join(DOCS, name + ".png")
        print(f"打磨图 {name} ->", out,
              capture_main(pdfs, out, size=size, scaling=scaling,
                           wait_preview=wait_preview))

    before = os.path.join(DOCS, "layout_2_1280.png")
    after = os.path.join(DOCS, "polish_1280.png")
    if os.path.isfile(before):
        composite = os.path.join(DOCS, "polish_before_after.png")
        _composite_before_after(before, after, composite)
        print("前后对照 ->", composite)


if __name__ == "__main__":
    main()
