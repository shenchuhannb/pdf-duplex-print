# -*- coding: utf-8 -*-
"""主界面（tkinter/ttk）。

界面分为五个页签：文件、打印设置、版面、预览、日志；底部固定执行区。
打印任务全部在后台线程执行，通过队列回调更新界面，保证不冻结。
"""

import ctypes
import logging
import os
import queue
import shutil
import subprocess
import tempfile
import threading
import time
import traceback
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk
from typing import Dict, List, Optional, Tuple

from core_utils import (
    CancelledByUser, PdfError, PrintError, PrinterOfflineError, PrintTimeoutError,
    SettingsError, ToolError, format_page_list, is_pdf_path, is_word_path, is_wps_path,
    parse_page_range, split_drop_data,
)
from image_convert import (
    ImageConvertError, image_to_pdf, is_image_path, supported_extensions_text,
)
from word_convert import (
    PRINT_TEMP_PREFIX, BatchCancelled, WordConvertError, available_converters,
    cleanup_stale_temp_dirs, convert_batch, convert_file, create_temp_dir,
    font_warning_text, pdf_font_warning_text,
)
from pdf_layout import (
    CUSTOM_PAPER_NAME, NUP_CHOICES, PAPER_SIZES_MM, LayoutOptions, build_layout_plan,
    compose_layout_pdf, compute_duplex_orders, compute_merged_duplex_orders,
    compute_merged_per_copy_orders, compute_single_sided_orders, describe_plan,
    get_pdf_page_count, merge_layout_pdfs, merged_side_ranges, merged_sheet_count,
    plan_batch_summary, preview_metrics, resolve_range_pages,
)
from printer_core import (
    GdiPrinter, PrinterOptions, cancel_job, check_printer_ready, get_default_printer,
    get_printable_insets_pt, get_printer_sources, get_printer_status, list_printers,
    open_directory, set_job_paused, wait_for_print_job,
)
from settings_store import (
    acrobat_defaults, add_recent, default_settings, load_settings, save_settings,
    settings_path,
)
from window_geometry import (
    DEFAULT_SIZE as WINDOW_DEFAULT_GEOMETRY_SIZE, MIN_SIZE as WINDOW_MIN_SIZE,
    choose_geometry, format_geometry, list_work_areas,
)
import theme as _theme
from app_identity import apply_window_icon

_logger = logging.getLogger("pdf_duplex")

try:
    import tkinterdnd2 as _tkdnd
except ImportError:  # 可选依赖
    _tkdnd = None

APP_TITLE = "PDF 批量双面打印工具"
PREVIEW_MAX_THUMBS = 40   # 预览缩略图最多渲染页数，避免超大文档卡顿
PREVIEW_LONG_EDGE = 880   # 成纸预览渲染像素：按最长边渲染，Canvas 再按控件大小缩放
LIVE_REFRESH_DEBOUNCE_MS = 100  # 设置改动后的防抖时间：连续输入只在停顿后刷新一次
# 间距规范：4px 细间距 / 8px 常规间距 / 12px 卡片间距；控件内边距统一 12×10
GAP_S = 4                        # 同一组控件内部（如按钮与文字）的细间距
GAP = 8                          # 8px 基准网格：分区间距/内边距的倍数
GAP_CARD = 12                    # 卡片与卡片之间的间距
CARD_PAD = (10, 6)               # 卡片内边距：左右 10 / 上下 6（720 高度预算）
# 两套布局定义（运行路径只用默认一种；另一套可显式传入 layout_scheme 构建/截图）
LAYOUT_SCHEMES = ("stacked", "split")
DEFAULT_LAYOUT_SCHEME = "stacked"
# 窗口默认/最小尺寸（与 window_geometry 保持一致）
WINDOW_DEFAULT_SIZE = WINDOW_DEFAULT_GEOMETRY_SIZE

# 界面显示用中文标签（变量内部仍保存英文枚举值，便于配置兼容与测试）
SCALE_LABELS = {
    "shrink": "适合可打印区",
    "fit": "拉伸铺满纸张",
    "actual": "实际大小 100%",
    "custom": "自定义百分比",
    "paper": "按纸张缩放",
}
SCALE_LABEL_TO_VALUE = {label: value for value, label in SCALE_LABELS.items()}
RANGE_LABELS = {
    "all": "全部页面",
    "current": "当前页",
    "custom": "指定范围",
    "odd": "仅奇数页",
    "even": "仅偶数页",
    "recent": "上次范围",
}
RANGE_LABEL_TO_VALUE = {label: value for value, label in RANGE_LABELS.items()}
# 高级设置下拉：界面只显示中文，内部仍保存英文枚举值（配置与测试兼容）
PAGE_ORDER_LABELS = {"row": "从左到右", "column": "从上到下"}
COLOR_LABELS = {"color": "彩色", "gray": "灰度", "mono": "黑白"}
QUALITY_LABELS = {"draft": "草稿", "normal": "标准", "high": "高质量"}
WAIT_LABELS = {"queue": "监听打印队列（推荐）", "delay": "固定延时（秒）"}


def _fmt_number(value) -> str:
    """数值显示：整数去掉 .0。"""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    return str(int(number)) if abs(number - int(number)) < 1e-9 else f"{number:g}"


def _create_root() -> tk.Tk:
    """创建主窗口：有 tkinterdnd2 时用其根窗口以支持拖拽，否则退回普通 Tk。

    Windows 上偶发 Tcl/Tk 库文件瞬时读不到（杀毒软件或索引服务占用），
    重试几次即可，避免启动直接失败。
    """
    last_error = None
    for _ in range(3):
        if _tkdnd is not None:
            try:
                return _tkdnd.TkinterDnD.Tk()
            except Exception:
                pass
        try:
            return tk.Tk()
        except tk.TclError as exc:
            last_error = exc
            time.sleep(0.3)
    raise last_error


class DuplexPrintApp:
    def __init__(self, root: tk.Tk, style: str = "", layout_scheme: str = "",
                 ui_scale_override: float = 0.0):
        self.root = root
        # 布局方案：stacked=上下分区（默认，推荐）/ split=左右分栏（备选定义）
        self.layout_scheme = (
            layout_scheme if layout_scheme in LAYOUT_SCHEMES else DEFAULT_LAYOUT_SCHEME)
        self.log_drawer_open = True      # 日志常驻展开，不再折叠
        self.settings = load_settings()
        self.style_key = _theme.resolve_style(style or str(self.settings.get("ui_style") or ""))
        self.colors = _theme.apply_theme(root, self.style_key)
        # 高 DPI 系数：100% 时为 1.0，125%/150% 时按比例放大固定像素间距与行高；
        # 字号用 pt（Tk 会随 tk scaling 自动放大），避免 150% 下文字被裁切。
        # ui_scale_override 仅供离屏截图/测试固定 DPI 使用，不影响真实运行。
        if ui_scale_override:
            self.ui_scale = float(ui_scale_override)
        else:
            try:
                self.ui_scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
            except tk.TclError:
                self.ui_scale = 1.0
        self.recent_files: List[str] = list(self.settings.get("recent_files") or [])
        self.recent_ranges: List[str] = list(self.settings.get("recent_ranges") or [])
        self.files: List[dict] = []          # {"path","pages","iid","status","range_text","copies"}
        self.ui_queue: "queue.Queue[tuple]" = queue.Queue()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.worker: Optional[threading.Thread] = None
        # 队列/状态轮询的 after id：退出时显式取消，避免窗口销毁后 Tcl 报
        # “invalid command name ..._poll_ui_queue”的残留回调警告。
        self.ui_poll_id = None
        self.printer_poll_id = None
        self.shutting_down = False   # 退出中：停止一切 after 轮询
        self._live_refresh_id = None  # 设置联动防抖的 after id
        # PDF 页数缓存：path -> (mtime_ns, size, pages)；避免每次刷新都重开文件
        self.pdf_info_cache: Dict[str, Tuple[int, int, int]] = {}
        self.paused = False
        self.current_job_id = 0
        self.print_active = False
        self.phase_var = tk.StringVar(value="")
        self.metric_vars = {
            "files": tk.StringVar(value="0"),
            "pages": tk.StringVar(value="0"),
            "sheets": tk.StringVar(value="0"),
            "flips": tk.StringVar(value="1"),
        }

        # 预览（后台线程）
        self.preview_enabled = True
        self.preview_generation = 0
        self.preview_thread: Optional[threading.Thread] = None
        self.preview_after_id = None
        self.preview_temp_dir = ""
        # Word → 临时 PDF 转换：串行线程 + 每个源文件的临时目录（退出时清理）
        self.word_thread: Optional[threading.Thread] = None
        self.word_temp_dirs: Dict[str, str] = {}   # 源文件路径 -> 临时目录
        self.preview_pages = 0
        self.preview_page_index = 0
        self.preview_pngs: Dict[int, bytes] = {}
        self.preview_photo = None
        self.preview_page_var = tk.StringVar(value="— / —")
        self.paper_sources: List[Tuple[str, int]] = [("自动选择", 7)]

        self._init_vars()
        self._build_ui()
        # 启动清理上次运行残留的临时目录（进程已退出/崩溃留下的）
        self.stale_temp_removed: List[str] = []
        try:
            self.stale_temp_removed = cleanup_stale_temp_dirs()
        except Exception:
            self.stale_temp_removed = []
        if self.stale_temp_removed:
            self._append_log(
                f"已清理 {len(self.stale_temp_removed)} 个上次残留的临时目录。")
        # 启动时恢复/居中窗口（校验保存的几何是否仍在可见工作区内）
        self._apply_window_geometry()
        self._bind_shortcuts()
        self._setup_drag_drop()
        self._refresh_printers()
        self._populate_recent_menu()
        self._sync_widgets()
        self._update_preview()
        self.ui_poll_id = self.root.after(100, self._poll_ui_queue)
        if self.preview_enabled:
            self.printer_poll_id = self.root.after(1500, self._poll_printer_status)

    # ------------------------------------------------------------------ 变量
    def _init_vars(self) -> None:
        s = self.settings

        def sv(key: str) -> tk.StringVar:
            return tk.StringVar(value=str(s.get(key, "")))

        def bv(key: str) -> tk.BooleanVar:
            return tk.BooleanVar(value=bool(s.get(key)))

        self.vars: Dict[str, tk.Variable] = {
            "printer": tk.StringVar(value=str(s.get("printer") or "")),
            "range_mode": sv("range_mode"),
            "range_text": sv("range_text"),
            "per_file": bv("per_file"),
            "copies": sv("copies"),
            "collate": sv("collate"),
            "manual_duplex": bv("manual_duplex"),
            "color_mode": sv("color_mode"),
            "ink_saver": bv("ink_saver"),
            "print_quality": sv("print_quality"),
            "dpi": sv("dpi"),
            "annots": bv("annots"),
            "scale_mode": sv("scale_mode"),
            "scale_percent": sv("scale_percent"),
            "nup": sv("nup"),
            "nup_border": bv("nup_border"),
            "page_order": sv("page_order"),
            "booklet": bv("booklet"),
            "paper_name": sv("paper_name"),
            "paper_width_mm": sv("paper_width_mm"),
            "paper_height_mm": sv("paper_height_mm"),
            "source_name": sv("source_name"),
            "orientation": sv("orientation"),
            "auto_rotate": bv("auto_rotate"),
            "margin_left_mm": sv("margin_left_mm"),
            "margin_top_mm": sv("margin_top_mm"),
            "margin_right_mm": sv("margin_right_mm"),
            "margin_bottom_mm": sv("margin_bottom_mm"),
            "crop_left_mm": sv("crop_left_mm"),
            "crop_top_mm": sv("crop_top_mm"),
            "crop_right_mm": sv("crop_right_mm"),
            "crop_bottom_mm": sv("crop_bottom_mm"),
            "reverse": bv("reverse"),
            "skip_blank": bv("skip_blank"),
            "wait_mode": sv("wait_mode"),
            "delay_seconds": sv("delay_seconds"),
            "auto_open": bv("auto_open"),
            "keep_output": bv("keep_output"),
            "output_dir": sv("output_dir"),
            "print_flow": sv("print_flow"),
            "copy_exec_on": bv("copy_exec_on"),
        }
        # 兼容旧测试/外部代码的常用别名
        self.range_mode = self.vars["range_mode"]
        self.copies_var = self.vars["copies"]
        self.dpi_var = self.vars["dpi"]
        self.printer_var = self.vars["printer"]
        self.wait_mode = self.vars["wait_mode"]
        self.delay_var = self.vars["delay_seconds"]
        self.range_mode_var = self.vars["range_mode"]
        self.collate_var = self.vars["collate"]
        self.per_file_var = self.vars["per_file"]
        self.nup_var = self.vars["nup"]
        self.booklet_var = self.vars["booklet"]
        self.paper_var = self.vars["paper_name"]
        self.source_var = self.vars["source_name"]
        self.orientation_var = self.vars["orientation"]
        self.scale_mode_var = self.vars["scale_mode"]
        self.scale_percent_var = self.vars["scale_percent"]
        self.about_note = tk.StringVar(
            value="手动双面为本工具核心能力：先正序打印奇数面，翻面后逆序打印偶数面。"
        )
        # 输出方式：single=单面（Acrobat 默认口径）/ duplex=手动双面
        self.duplex_mode_var = tk.StringVar(
            value="duplex" if bool(s.get("manual_duplex")) else "single")
        # 预览视图：single=单页 / spread=双页并排
        self.preview_view_var = tk.StringVar(value=str(s.get("preview_view") or "single"))
        # 窗口几何与高级设置对话框
        self.advanced_window = None
        self.printable_insets = None
        self.printable_insets_printer = ""
        self.preview_metrics_data: Dict[str, object] = {}
        # 当前预览对应的版式签名：相同则只重绘 Canvas，不重新渲染 PDF
        self._preview_signature = None
        # 成纸预览的位图缓存：key=(面序号, 目标宽, 目标高)
        self._paper_photo_cache: Dict[Tuple[int, int, int], object] = {}

    # ------------------------------------------------------------------ 界面
    def _px(self, value: float) -> int:
        """按当前 DPI 换算固定像素值：100% 原值，125%/150% 等比放大。"""
        return int(round(float(value) * self.ui_scale))

    def _card_padding(self) -> Tuple[int, int]:
        return (self._px(CARD_PAD[0]), self._px(CARD_PAD[1]))

    def _build_ui(self) -> None:
        """界面骨架：顶部信息条 → 工作区 → 日志抽屉 → 底部设置/操作条。

        采用 8px 基准网格；主次动线为“文件列表 + 成纸预览 + 开始打印”，
        低频项（页边距/N-up/纸张/质量…）全部收进“高级设置…”。
        """
        self.root.title(APP_TITLE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        # Tk 的 iconbitmap 只覆盖窗口标题栏的一部分路径；任务栏按钮还需 WM_SETICON
        # 设置 ICON_BIG/ICON_SMALL（配合进程 AUMID 才能稳定显示本项目图标）。
        if os.path.isfile(icon_path):
            self.icon_handles = apply_window_icon(self.root, icon_path)
        # 行高随 DPI 放大（theme 里给的是 100% 基准值 28px）
        try:
            ttk.Style(self.root).configure("Treeview", rowheight=self._px(_theme.ROW_HEIGHT))
        except tk.TclError:
            pass

        # 内边距从 16 收到 12：1024 宽窗口下给内容多留 8px，避免右侧控件贴边
        container = tk.Frame(self.root, background=self.colors["bg"],
                             padx=self._px(12), pady=self._px(10))
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        self._build_header(container)

        self.workspace = ttk.Frame(container, style="Page.TFrame")
        self.workspace.grid(row=1, column=0, sticky="nsew", pady=(self._px(GAP), 0))
        self.workspace.columnconfigure(0, weight=1)
        self.workspace.rowconfigure(0, weight=1)
        if self.layout_scheme == "split":
            self._build_workspace_split()
        else:
            self._build_workspace_stacked()

        if self.layout_scheme == "stacked":
            self._build_control_bar(container)
        else:
            self._build_action_row(self.action_slot, vertical=True)

        # 日志固定在最底部：分隔条(row4) + 常驻日志区(row5)
        self._build_log_drawer(container)

        # 全量实时绑定：所有可调项改动后都触发同样的联动（内部做 100ms 防抖）
        for _name, var in self.vars.items():
            var.trace_add("write", lambda *_a: self._on_setting_change())
        # 窗口拉伸后重新排版预览（防抖合并连续 Configure 事件）
        self.root.bind("<Configure>", self._on_root_configure)
        # Tab 顺序与视觉顺序一致（默认创建顺序是“文件区→预览→设置→操作”，
        # 与用户看到的自上而下顺序不同，因此显式指定焦点链）
        self._setup_tab_order()

    # ------------------------------------------------------- 方案 2：上下分区
    def _build_workspace_stacked(self) -> None:
        """上区 = 文件列表（左）+ 成纸预览（右）；下区 = 紧凑设置条 + 主操作。"""
        ws = self.workspace
        # 文件列表 : 预览 = 5 : 6：预览列略宽于列表，纸张观感更接近实际打印效果。
        # 最小宽按 1024 窗口预算：列表 5/11≈449px、预览 6/11≈539px，两列都能放下。
        ws.columnconfigure(0, weight=5, minsize=self._px(410), uniform="main")
        ws.columnconfigure(1, weight=6, minsize=self._px(430), uniform="main")
        ws.rowconfigure(0, weight=1)

        files_card = ttk.Frame(ws, style="Card.TFrame", padding=self._card_padding())
        files_card.grid(row=0, column=0, sticky="nsew", padx=(0, self._px(GAP_CARD)))
        self._build_file_panel(files_card)

        preview_card = ttk.Frame(ws, style="Card.TFrame", padding=self._card_padding())
        preview_card.grid(row=0, column=1, sticky="nsew")
        self._build_preview_panel(preview_card)

    def _build_control_bar(self, parent: tk.Frame) -> None:
        """底部设置/操作条：两行设置 + 固定错误提示行 + 主操作行。"""
        c = self.colors
        bar = tk.Frame(parent, background=c["surface"],
                       padx=self._px(12), pady=self._px(8),
                       highlightthickness=1, highlightbackground=c["border"])
        bar.grid(row=2, column=0, sticky="ew", pady=(self._px(GAP_CARD), 0))
        bar.columnconfigure(0, weight=1)

        settings = ttk.Frame(bar, style="Bar.TFrame")
        settings.grid(row=0, column=0, sticky="ew")
        self._build_settings_strip(settings)

        ttk.Separator(bar, orient="horizontal", style="Line.TSeparator").grid(
            row=1, column=0, sticky="ew", pady=(self._px(GAP_S), self._px(GAP_S)))

        action = ttk.Frame(bar, style="Bar.TFrame")
        action.grid(row=2, column=0, sticky="ew")
        self._build_action_row(action, vertical=False)

    # ------------------------------------------------------- 方案 1：左右分栏
    def _build_workspace_split(self) -> None:
        """左列 = 文件列表（上）+ 成纸预览（下）；右列 = 设置表单 + 主操作。"""
        ws = self.workspace
        ws.columnconfigure(0, weight=7, minsize=self._px(470), uniform="main")
        ws.columnconfigure(1, weight=5, minsize=self._px(400), uniform="main")
        ws.rowconfigure(0, weight=1)

        left = ttk.Frame(ws, style="Page.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, self._px(GAP_CARD)))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=3, minsize=self._px(200))
        left.rowconfigure(1, weight=4, minsize=self._px(240))

        files_card = ttk.Frame(left, style="Card.TFrame", padding=self._card_padding())
        files_card.grid(row=0, column=0, sticky="nsew", pady=(0, self._px(GAP_CARD)))
        self._build_file_panel(files_card)

        preview_card = ttk.Frame(left, style="Card.TFrame", padding=self._card_padding())
        preview_card.grid(row=1, column=0, sticky="nsew")
        self._build_preview_panel(preview_card)

        right = ttk.Frame(ws, style="Card.TFrame",
                          padding=(self._px(16), self._px(12)))
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        form = ttk.Frame(right)
        form.grid(row=0, column=0, sticky="new")
        self.settings_panel = form
        self._build_settings_form(form)

        ttk.Separator(right, orient="horizontal", style="Line.TSeparator").grid(
            row=2, column=0, sticky="ew", pady=(self._px(GAP), self._px(GAP)))
        self.action_slot = ttk.Frame(right)
        self.action_slot.grid(row=3, column=0, sticky="ew")

    def _on_root_configure(self, event) -> None:
        if getattr(event, "widget", None) is not self.root:
            return
        self._layout_header_compact()
        self._sync_preview_info_visibility()
        self._resize_log_for_window()
        self._schedule_live_refresh()

    def _build_header(self, parent: tk.Frame) -> None:
        """顶部标题/状态条：品牌标识 + 标题 + 关键指标 + 打印机状态。

        指标徽标使用固定宽度容器（宽度不随数字长度变化），数字与标签都按
        relx=0.5 放在卡片正中，使用同一中轴；数字用等宽数字字体 Segoe UI，
        位数变化（1→123）只在卡片内居中，不推动相邻徽标、也不会左右漂移。
        """
        c = self.colors
        header = tk.Frame(parent, background=c["header_bg"], height=self._px(72),
                          highlightthickness=1, highlightbackground=c["border"])
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(2, weight=1)
        self.header_frame = header
        self._header_full_height = self._px(72)
        self._header_compact_height = self._px(58)

        mark = tk.Frame(header, background=c["accent"], width=self._px(36),
                        height=self._px(36))
        mark.grid(row=0, column=0, rowspan=2, padx=(self._px(12), self._px(10)),
                  pady=self._px(12))
        mark.pack_propagate(False)
        tk.Label(mark, text="PDF", background=c["accent"], foreground=c["accent_text"],
                 font=(_theme.FONT_FAMILY, 10, "bold")).place(relx=0.5, rely=0.5, anchor="center")

        self.header_title_label = tk.Label(
            header, text="PDF 批量双面打印", background=c["header_bg"],
            foreground=c["text"], font=_theme.FONT_TITLE, anchor="w")
        self.header_title_label.grid(row=0, column=1, sticky="sw", pady=(self._px(14), 0))
        # 主流程步骤图：① 选文件 → ② 转换 → ③ 打印。
        # 旧版这里是一行静态卖点文字，用户无法判断“现在该做什么”；
        # 改为三步骤图，当前步骤加粗高亮、已完成转绿，转换失败转红。
        steps = tk.Frame(header, background=c["header_bg"])
        steps.grid(row=1, column=1, sticky="nw", pady=(self._px(2), self._px(10)))
        self.header_steps_frame = steps
        self.step_labels: List[tk.Label] = []
        self.step_captions = ("① 选文件", "② 转换", "③ 打印")
        self.step_states: Tuple[str, str, str] = ("active", "pending", "pending")
        for index, caption in enumerate(self.step_captions):
            if index:
                tk.Label(steps, text="›", background=c["header_bg"],
                         foreground=c["border"], font=_theme.FONT_SMALL).pack(
                    side="left", padx=self._px(6))
            label = tk.Label(steps, text=caption, background=c["header_bg"],
                             foreground=c["muted"], font=_theme.FONT_SMALL, anchor="w")
            label.pack(side="left")
            self.step_labels.append(label)
        # 兼容旧方法名：紧凑模式显隐同一个控件
        self.header_subtitle_label = steps

        metrics = tk.Frame(header, background=c["header_bg"])
        metrics.grid(row=0, column=2, rowspan=2, sticky="e", padx=(self._px(8), self._px(10)))
        self.metric_chips: Dict[str, dict] = {}
        metric_defs = (("files", "文件"), ("pages", "选取页"), ("sheets", "预计张数"),
                       ("flips", "翻面次数"))
        # 徽标尺寸按实际字体行高计算：150% 下数字行高会明显变大，
        # 固定 48px 高 + 固定 y 会让数字与下方说明文字重叠。
        chip_top = self._px(3)
        for index, (key, caption) in enumerate(metric_defs):
            # 固定宽度 88px：数字位数变化（1→123）只在本框内居中，不推动相邻徽标
            chip = tk.Frame(metrics, background=c["surface_alt"], width=self._px(88),
                            highlightthickness=1, highlightbackground=c["border"])
            chip.grid(row=0, column=index, padx=(0, self._px(GAP)))
            chip.grid_propagate(False)
            value_label = tk.Label(
                chip, textvariable=self.metric_vars[key], background=c["surface_alt"],
                foreground=c["accent"], font=_theme.FONT_NUM, anchor="center",
                justify="center")
            value_label.place(relx=0.5, y=chip_top, anchor="n")
            caption_label = tk.Label(
                chip, text=caption, background=c["surface_alt"], foreground=c["muted"],
                font=_theme.FONT_SMALL, anchor="center", justify="center")
            # 说明文字紧贴数字控件的实际请求高度摆放，避免 150% 下数字行高
            # 比 font.metrics 估算值更大而造成 3~4px 重叠。
            caption_top = chip_top + value_label.winfo_reqheight() + self._px(1)
            caption_label.place(relx=0.5, y=caption_top, anchor="n")
            chip.configure(height=caption_top + caption_label.winfo_reqheight()
                           + self._px(2))
            self.metric_chips[key] = {"frame": chip, "value": value_label,
                                      "caption": caption_label}

        status_chip = tk.Frame(header, background=c["header_bg"], width=self._px(160))
        status_chip.grid(row=0, column=3, rowspan=2, sticky="e",
                         padx=(self._px(2), self._px(12)))
        status_chip.grid_propagate(False)
        self.printer_status_var = tk.StringVar(value="状态：未检测")
        self.printer_status_label = tk.Label(
            status_chip, textvariable=self.printer_status_var, background=c["header_bg"],
            foreground=c["accent"], font=(_theme.FONT_FAMILY, 9), anchor="e")
        self.printer_status_label.pack(anchor="e")
        tk.Label(status_chip, text="手动双面打印", background=c["header_bg"],
                 foreground=c["muted"], font=_theme.FONT_SMALL, anchor="e").pack(
            anchor="e", pady=(self._px(GAP_S), 0))

    def _layout_header_compact(self) -> None:
        """窗口较矮时收起标题副标题，把高度让给预览区（1024×720 仍然完整可见）。"""
        if not hasattr(self, "header_frame"):
            return
        try:
            compact = self.root.winfo_height() < self._px(820)
        except tk.TclError:
            return
        if compact:
            # 紧凑高度（1024×720 / 高 DPI）下没有第二行空间：隐藏完整步骤图，
            # 改由文件卡右上角的“第 N 步”提示与阶段徽标承担当前阶段提示。
            self.header_subtitle_label.grid_remove()
            self.header_frame.configure(height=self._header_compact_height)
        else:
            self.header_subtitle_label.grid()
            self.header_subtitle_label.grid_configure(
                pady=(self._px(2), self._px(10)))
            self.header_frame.configure(height=self._header_full_height)

    def _update_flow_steps(self) -> None:
        """更新顶部三步骤图（选文件 → 转换 → 打印）的当前状态。

        旧问题：副标题是固定卖点文案，用户看不出“文件是否已就绪、下一步做什么”。
        新效果：当前步骤蓝色加粗，已完成转绿，转换失败行转红；转换中仍停留在第 2 步。
        """
        if not getattr(self, "step_labels", None):
            return
        c = self.colors
        files = getattr(self, "files", [])
        pending = [f for f in files
                   if f.get("kind") in ("word", "wps", "image")
                   and not f.get("pdf_path") and f.get("status") != "转换失败"]
        failed = any(str(f.get("status") or "") in ("转换失败", "无法读取") for f in files)
        status_text = str(self.status_var.get()) if hasattr(self, "status_var") else ""
        if not files:
            states = ["active", "pending", "pending"]
        elif pending:
            states = ["done", "active", "pending"]
        elif failed:
            states = ["done", "danger", "active"]
        elif getattr(self, "print_active", False):
            states = ["done", "done", "active"]
        elif status_text.startswith("整批打印完成"):
            states = ["done", "done", "done"]
        else:
            states = ["done", "done", "active"]
        self.step_states = tuple(states)
        palette = {"done": c["success"], "active": c["accent"], "danger": c["danger_text"]}
        for label, state in zip(self.step_labels, states):
            label.configure(
                foreground=palette.get(state, c["muted"]),
                font=(_theme.FONT_FAMILY, 8, "bold") if state == "active"
                else _theme.FONT_SMALL)

    # ------------------------------------------------------------- 文件列表
    def _build_file_panel(self, frame: ttk.Frame) -> None:
        """文件列表（方案 1/2 共用）：标题行 + 范围提示 + 操作条 + 表格 + 选中操作条 + 最近使用。"""
        self.tab_files = frame
        frame.rowconfigure(3, weight=1)
        frame.columnconfigure(0, weight=1)

        title = ttk.Frame(frame)
        title.grid(row=0, column=0, sticky="ew")
        title.columnconfigure(1, weight=1)
        ttk.Label(title, text="文件列表", style="Subtitle.TLabel").grid(row=0, column=0, sticky="w")
        self.file_count_var = tk.StringVar(value="0 个文件")
        ttk.Label(title, textvariable=self.file_count_var, style="Muted.TLabel").grid(
            row=0, column=1, sticky="w", padx=(self._px(GAP), 0))
        self.drop_hint_var = tk.StringVar(value="可拖入 PDF / Word / 图片")
        self.drop_hint_label = ttk.Label(title, textvariable=self.drop_hint_var,
                                         style="Muted.TLabel", anchor="e")
        self.drop_hint_label.grid(row=0, column=2, sticky="e")

        # 范围提示：独立一行（per_file 开启时才显示），口语化说明“按上方范围/＊单设”
        self.scope_hint_var = tk.StringVar(value="")
        self.scope_hint_label = ttk.Label(
            frame, textvariable=self.scope_hint_var, style="Muted.TLabel",
            anchor="w", wraplength=self._px(400), justify="left")
        self.scope_hint_label.grid(row=1, column=0, sticky="w", pady=(self._px(2), 0))
        self.scope_hint_label.grid_remove()

        # 操作条：单行排列（1024 宽下列表卡约 443px，8 个按钮一行放不下，
        # 低频率的上移/下移/去重/打开目录收进“更多”菜单，主按钮保持可见）
        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="ew",
                     pady=(self._px(GAP), self._px(GAP)))
        actions.columnconfigure(3, weight=1)
        self.add_file_btn = ttk.Button(actions, text="＋ 添加文件", style="Soft.TButton",
                                       command=self.add_files)
        self.add_dir_btn = ttk.Button(actions, text="添加文件夹", style="Ghost.TButton",
                                      command=self.add_folder)
        self.remove_btn = ttk.Button(actions, text="移除", command=self.remove_selected)
        self.clear_btn = ttk.Button(actions, text="清空", style="Danger.TButton",
                                    command=self.clear_files)
        self.add_file_btn.grid(row=0, column=0, sticky="w", padx=(0, self._px(GAP_S)))
        self.add_dir_btn.grid(row=0, column=1, sticky="w", padx=(0, self._px(GAP_S)))
        self.remove_btn.grid(row=0, column=2, sticky="w")
        self.more_btn = ttk.Menubutton(actions, text="更多 ▾", style="Ghost.TButton")
        more_menu = tk.Menu(self.more_btn, tearoff=0)
        more_menu.add_command(label="上移", command=lambda: self.move_selected(-1))
        more_menu.add_command(label="下移", command=lambda: self.move_selected(1))
        more_menu.add_command(label="去重", command=self.dedupe_files)
        more_menu.add_command(label="打开目录", command=self.show_in_explorer)
        self.more_btn.configure(menu=more_menu)
        self.more_menu = more_menu
        self.more_btn.grid(row=0, column=4, sticky="w", padx=(0, self._px(GAP_S)))
        # 危险操作靠右、与主操作分离（“清空”不再和“移除”相邻）
        self.clear_btn.grid(row=0, column=5, sticky="e")

        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=3, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("name", "pages", "scope", "plan", "status"),
            show="headings",
            height=8,
            selectmode="extended",
        )
        # 基础宽度只是初值；每次刷新/窗口变化后由 _fit_tree_columns() 按内容实测：
        # 文件名列吃剩余宽度，其余列按内容自适应，保证各列宽之和 ≤ 可视宽度。
        headings = (("name", "文件名", 130, 90, "w"), ("pages", "总页", 40, 36, "center"),
                    ("scope", "本次范围", 88, 60, "center"),
                    ("plan", "输出与补页", 94, 72, "center"),
                    ("status", "状态", 84, 64, "center"))
        for key, text, width, minwidth, anchor in headings:
            # 表头与单元格使用同一对齐基准，避免“标题居中、内容靠左”的错位感
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=self._px(width), minwidth=self._px(minwidth),
                             anchor=anchor, stretch=(key == "name"))
        self._init_tree_tags()
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", style="Thin.Vertical.TScrollbar",
                               command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree_scroll = scroll
        # 空列表引导：旧版无文件时表格是一片空白，新用户不知道可以拖入/点按钮。
        # 与 Treeview 共用同一个 grid 单元格，空列表时隐藏表格、显示引导；
        # 有文件时切换回来，既不占额外高度，也不产生控件重叠。
        c = self.colors
        self.empty_state = tk.Frame(tree_frame, background=c["surface"])
        self.empty_state.grid(row=0, column=0, sticky="nsew")
        box = tk.Frame(self.empty_state, background=c["surface"])
        box.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(box, text="还没有待打印文件",
                 background=c["surface"], foreground=c["text"],
                 font=_theme.FONT_SUBTITLE).pack()
        tk.Label(box,
                 text="把 PDF / Word / 图片拖到这里，或点上方“＋ 添加文件”",
                 background=c["surface"], foreground=c["muted"],
                 font=_theme.FONT_BODY, wraplength=self._px(320),
                 justify="center").pack(pady=(self._px(4), 0))
        tk.Label(box, text="Word / WPS / 图片会先自动转换为 PDF 再打印",
                 background=c["surface"], foreground=c["muted"],
                 font=_theme.FONT_SMALL).pack(pady=(self._px(2), 0))
        self._update_empty_state()
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._show_context_menu)
        # 键盘：列表有焦点时 Ctrl+A 全选、Esc 取消选中（不改变双击快捷编辑）
        self.tree.bind("<Control-a>", self._select_all_on_tree)
        self.tree.bind("<Control-A>", self._select_all_on_tree)
        self.tree.bind("<Escape>", self._clear_tree_selection)
        # 文件名被省略时，悬停 0.6s 显示完整文件名与路径（离屏窗口不弹出）
        self._tree_tip = None
        self._tree_tip_after = None
        self._tree_tip_key = None
        self.tree.bind("<Motion>", self._on_tree_motion, add="+")
        self.tree.bind("<Leave>", lambda _e: self._hide_tree_tooltip(), add="+")
        self.tree.bind("<ButtonPress-1>", lambda _e: self._hide_tree_tooltip(), add="+")
        # 只在可视宽度真正变化时重算列宽，避免“改列宽→触发 Configure→再改”的循环
        self.tree.bind("<Configure>", self._on_tree_configure)

        self._build_context_menu()

        # 选中操作条：信息一行 + 按钮一行（1024 宽下按钮必须完整可见，不能挤成半截）
        sel_bar = ttk.Frame(frame)
        sel_bar.grid(row=4, column=0, sticky="ew", pady=(self._px(GAP), 0))
        sel_bar.columnconfigure(0, weight=1)
        self.sel_info_var = tk.StringVar(value="未选择文件")
        self.sel_info_label = ttk.Label(sel_bar, textvariable=self.sel_info_var,
                                        style="Muted.TLabel", anchor="w")
        self.sel_info_label.grid(row=0, column=0, sticky="ew")
        self.sel_range_btn = ttk.Button(
            sel_bar, text="设置范围…", command=lambda: self.edit_selected_scope("range"))
        self.sel_copies_btn = ttk.Button(
            sel_bar, text="设置份数…", command=lambda: self.edit_selected_scope("copies"))
        self.sel_restore_btn = ttk.Button(
            sel_bar, text="恢复按上方范围", style="Ghost.TButton",
            command=self.restore_global_for_selected)
        self.sel_apply_all_btn = ttk.Button(
            sel_bar, text="应用到全部文件", style="Ghost.TButton",
            command=self.apply_selected_scope_to_all)
        self.sel_buttons = [self.sel_range_btn, self.sel_copies_btn,
                            self.sel_restore_btn, self.sel_apply_all_btn]
        for index, button in enumerate(self.sel_buttons):
            button.grid(row=1, column=index, sticky="w",
                        padx=(0, self._px(GAP_S)), pady=(self._px(GAP_S), 0))
        # 125%/150% 下四个按钮的单行请求宽度会超过文件卡内宽，改为 2×2 排列，
        # 避免最左侧按钮被压成半截（不依赖固定像素宽度估算）。
        self.sel_bar = sel_bar
        sel_bar.bind("<Configure>", self._on_sel_bar_configure)
        self._update_selection_bar()

        # 最近使用：单行紧凑布局（下拉吸收剩余宽度，不挤压右侧按钮）
        recent = ttk.Frame(frame)
        recent.grid(row=5, column=0, sticky="ew", pady=(self._px(GAP), 0))
        recent.columnconfigure(1, weight=1)
        ttk.Label(recent, text="最近：", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.recent_var = tk.StringVar()
        self.recent_combo = ttk.Combobox(recent, textvariable=self.recent_var, state="readonly",
                                         width=12)
        self.recent_combo.grid(row=0, column=1, sticky="ew", padx=(0, self._px(GAP_S)))
        ttk.Button(recent, text="添加", command=self._add_recent_file).grid(
            row=0, column=2, padx=(0, self._px(GAP_S)))
        ttk.Button(recent, text="清除记录", style="Ghost.TButton",
                   command=self._clear_recent).grid(row=0, column=3)

    def _on_sel_bar_configure(self, event=None) -> None:
        """选中操作条宽度变化：单行放不下时改为 2×2，保证按钮文字完整。"""
        width = int(event.width) if event is not None else self.sel_bar.winfo_width()
        if width <= 1:
            return
        gap = self._px(GAP_S)
        required = sum(button.winfo_reqwidth() for button in self.sel_buttons) \
            + gap * max(0, len(self.sel_buttons) - 1)
        per_row = len(self.sel_buttons) if required <= width else 2
        key = (per_row, width)
        if getattr(self, "_sel_layout_key", None) == key:
            return
        self._sel_layout_key = key
        for index, button in enumerate(self.sel_buttons):
            button.grid_configure(row=1 + index // per_row, column=index % per_row,
                                  sticky="w", padx=(0, gap),
                                  pady=(self._px(GAP_S), 0))

    def _on_tree_configure(self, event=None) -> None:
        width = event.width if event is not None else self.tree.winfo_width()
        if abs(width - getattr(self, "_last_tree_width", -1)) < 2:
            return
        self._last_tree_width = width
        self._fit_tree_columns()

    def _fit_tree_columns(self) -> None:
        """按实际内容自适应列宽，并保证所有列宽之和 ≤ 可视宽度（不横向截断）。

        文件名列吃剩余宽度；范围/输出/状态列用实测文字宽度（含表头），
        设置上下限，避免长文案把其它列挤没。
        """
        if not hasattr(self, "tree"):
            return
        try:
            available = self.tree.winfo_width()
        except tk.TclError:
            return
        if available <= 1:
            return
        try:
            font = tkfont.Font(root=self.root, font=ttk.Style(self.root).lookup(
                "Treeview", "font") or "TkDefaultFont")
        except tk.TclError:
            return
        pad = self._px(12)
        columns = list(self.tree["displaycolumns"])
        if not columns or columns == ["#all"]:
            columns = list(self.tree["columns"])
        headings = {"name": "文件名", "pages": "总页", "scope": "本次范围",
                    "plan": "输出与补页", "status": "状态"}
        limits = {"name": (self._px(90), 10_000), "pages": (self._px(34), self._px(56)),
                  "scope": (self._px(58), self._px(104)),
                  "plan": (self._px(70), self._px(132)),
                  "status": (self._px(62), self._px(104))}
        desired: Dict[str, int] = {}
        for key in columns:
            if key == "name":
                continue
            longest = font.measure(headings[key])
            for iid in self.tree.get_children(""):
                try:
                    index = ("name", "pages", "scope", "plan", "status").index(key)
                    text = str(self.tree.item(iid, "values")[index])
                except (tk.TclError, IndexError, ValueError):
                    continue
                longest = max(longest, font.measure(text))
            low, high = limits[key]
            desired[key] = max(low, min(high, longest + pad))
        name_low = limits["name"][0]
        # 极端窄窗口：先把非文件名列压回各自下限，再让文件名列吃剩余宽度；
        # 无论哪种情况，列宽总和都严格 ≤ 可视宽度。
        excess = sum(desired.values()) + name_low - available
        if excess > 0:
            for key in ("status", "plan", "scope", "pages"):
                if key not in desired or excess <= 0:
                    continue
                low = limits[key][0]
                cut = min(excess, desired[key] - low)
                desired[key] -= cut
                excess -= cut
        fixed_total = sum(desired.values())
        name_width = max(name_low, available - fixed_total)
        self.tree.column("name", width=name_width)
        for key, width in desired.items():
            self.tree.column(key, width=width)
        # 文件名超长时主动加省略号（Treeview 原生是硬裁切，会显示成“半截字”）
        for record in self.files:
            iid = record.get("iid")
            if not iid:
                continue
            try:
                values = list(self.tree.item(iid, "values"))
            except tk.TclError:
                continue
            if not values:
                continue
            values[0] = self._elide_filename(
                os.path.basename(record["path"]), font, max(10, name_width - pad))
            self.tree.item(iid, values=values)

    @staticmethod
    def _elide_filename(name: str, font, max_px: int) -> str:
        """文件名超长时保留扩展名再省略（旧版会把 .pdf 一起截掉）。"""
        if font.measure(name) <= max_px:
            return name
        stem, ext = os.path.splitext(name)
        if ext:
            tail = "…" + ext
            if font.measure(tail) <= max_px:
                return DuplexPrintApp._elide_text(
                    stem, font, max_px - font.measure(tail)) + tail
        return DuplexPrintApp._elide_text(name, font, max_px)

    @staticmethod
    def _elide_text(text: str, font, max_px: int) -> str:
        """按像素宽度截断并加省略号（省略号前保留尽可能完整的词）。"""
        if font.measure(text) <= max_px:
            return text
        ellipsis = "…"
        low, high = 0, len(text)
        while low < high:
            middle = (low + high + 1) // 2
            if font.measure(text[:middle] + ellipsis) <= max_px:
                low = middle
            else:
                high = middle - 1
        return text[:low] + ellipsis

    def _init_tree_tags(self) -> None:
        """表格标签：斑马纹（奇偶行）+ 正在打印行（浅蓝底，与选择行区分）。"""
        c = self.colors
        self.tree.tag_configure("stripe", background=c["tree_stripe"])
        self.tree.tag_configure("active", background=c["active_row"],
                                foreground=c["accent"])
        # 逐文件范围非法：只标红该行，不影响其它文件与整批统计
        self.tree.tag_configure("invalid", background=c["danger_soft"],
                                foreground=c["danger_text"])
        # 未单独设置：整行浅灰，表示“按上方范围打印”（ttk 不支持单元格级颜色）
        self.tree.tag_configure("scope_auto", foreground=c["muted"])
        # 转换失败/无法读取：整行红色前景，长列表里也能一眼定位问题文件
        self.tree.tag_configure("failed", foreground=c["danger_text"])

    def _status_display(self, status: str) -> str:
        """状态列的“色点/徽标”文本：● 可打印、◌ 待转换、◆ 已转换、✕ 失败。"""
        status = str(status or "")
        mapping = {"待打印": "● 待打印", "待转换": "◌ 待转换", "转换中…": "◌ 转换中…",
                   "Word（已转换为 PDF）": "◆ 已转换", "WPS（已转换为 PDF）": "◆ 已转换",
                   "图片（已转换为 PDF）": "◆ 已转换", "转换失败": "✕ 转换失败",
                   "无法读取": "✕ 无法读取", "已完成": "✓ 已完成", "已跳过": "— 已跳过"}
        return mapping.get(status, status)

    @staticmethod
    def _is_active_status(status: str) -> bool:
        """是否为“正在打印/排版”的活动状态（用于行高亮；待打印不算）。"""
        text = str(status or "")
        if "待打印" in text:
            return False
        return any(word in text for word in (
            "排版中", "单面打印", "整批·正面", "整批·反面", "第1面", "第2面", "·正面", "·反面"))

    def _apply_row_tags(self) -> None:
        """按显示顺序重刷斑马纹与“正在打印”高亮。"""
        if not hasattr(self, "tree"):
            return
        for position, iid in enumerate(self.tree.get_children("")):
            try:
                values = self.tree.item(iid, "values")
                status = str(values[4]) if len(values) > 4 else ""
                tags = ["stripe"] if position % 2 else []
                record = next((f for f in self.files if f.get("iid") == iid), None)
                # 逐文件范围开启时：未单独设置的行整行浅灰（表示按上方范围打印）
                if (record is not None and bool(self.vars["per_file"].get())
                        and not self._record_has_custom_scope(record)):
                    tags.append("scope_auto")
                if record is not None and self._record_scope_invalid(record):
                    tags.append("invalid")
                clean = status
                for glyph in ("● ", "◌ ", "◆ ", "✕ ", "✓ "):
                    clean = clean.replace(glyph, "")
                if "转换失败" in clean or "无法读取" in clean:
                    tags.append("failed")
                if self._is_active_status(clean):
                    tags.append("active")
                self.tree.item(iid, tags=tuple(tags))
            except tk.TclError:
                return

    # 兼容旧命名
    def _build_file_tab(self) -> None:
        self._build_file_panel(self.tab_files)

    def _on_tree_double_click(self, event) -> None:
        """双击“本次范围”列直接编辑；双击其它位置仍打开预览。"""
        column = self.tree.identify_column(event.x)
        if column == "#3" and bool(self.vars["per_file"].get()):
            row = self.tree.identify_row(event.y)
            if row:
                self.tree.selection_set(row)
                self.edit_selected_scope()
            return
        self.open_selected_preview()

    def _on_tree_motion(self, event) -> None:
        """悬停被省略的文件名单元格：延迟弹出完整文件名 + 路径。"""
        if not getattr(self, "files", None):
            self._hide_tree_tooltip()
            return
        row = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not row or column != "#1":
            self._hide_tree_tooltip()
            return
        record = next((f for f in self.files if f.get("iid") == row), None)
        if record is None:
            self._hide_tree_tooltip()
            return
        full = os.path.basename(record["path"])
        try:
            shown = str(self.tree.item(row, "values")[0])
        except (tk.TclError, IndexError):
            self._hide_tree_tooltip()
            return
        if shown == full:
            self._hide_tree_tooltip()
            return
        key = (row, full)
        if getattr(self, "_tree_tip_key", None) == key:
            return
        self._hide_tree_tooltip()
        self._tree_tip_key = key
        x, y = event.x_root, event.y_root
        self._tree_tip_after = self.root.after(
            600, lambda: self._show_tree_tooltip(full, record["path"], x, y))

    def _show_tree_tooltip(self, name: str, path: str, x: int, y: int) -> None:
        self._tree_tip_after = None
        if getattr(self, "shutting_down", False):
            return
        try:
            alpha = float(self.root.attributes("-alpha"))
        except (tk.TclError, TypeError, ValueError):
            alpha = 1.0
        # 离屏截图 / 自动化窗口 alpha=0：绝不创建顶层窗口，避免意外上屏。
        if alpha < 0.95:
            return
        c = self.colors
        tip = tk.Toplevel(self.root)
        tip.overrideredirect(True)
        tip.configure(background=c["border"])
        tk.Label(tip, text=f"{name}\n{path}", background=c["surface_alt"],
                 foreground=c["text"], font=_theme.FONT_SMALL, justify="left",
                 anchor="w", wraplength=self._px(420), padx=self._px(8),
                 pady=self._px(4)).pack(padx=1, pady=1)
        tip.update_idletasks()
        tip.geometry(f"+{int(x) + self._px(12)}+{int(y) + self._px(16)}")
        self._tree_tip = tip

    def _hide_tree_tooltip(self) -> None:
        self._tree_tip_key = None
        after_id = getattr(self, "_tree_tip_after", None)
        if after_id is not None:
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
            self._tree_tip_after = None
        tip = getattr(self, "_tree_tip", None)
        if tip is not None:
            try:
                tip.destroy()
            except tk.TclError:
                pass
            self._tree_tip = None

    def _build_context_menu(self) -> None:
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="打开预览", command=self.open_selected_preview)
        self.context_menu.add_command(label="在资源管理器显示", command=self.show_in_explorer)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="上移", command=lambda: self.move_selected(-1))
        self.context_menu.add_command(label="下移", command=lambda: self.move_selected(1))
        self.context_menu.add_separator()
        # 与选中操作条保持一致：右键先选中整行，再按同一入口批量设置
        self.context_menu.add_command(label="设置范围…",
                                      command=lambda: self.edit_selected_scope("range"))
        self.context_menu.add_command(label="设置份数…",
                                      command=lambda: self.edit_selected_scope("copies"))
        self.context_menu.add_command(label="恢复按上方范围",
                                      command=self.restore_global_for_selected)
        self.context_menu.add_command(label="应用到全部文件",
                                      command=self.apply_selected_scope_to_all)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="去重", command=self.dedupe_files)
        self.context_menu.add_command(label="移除", command=self.remove_selected)

    def _show_context_menu(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    # ------------------------------------------------- 主界面设置（两种排布）
    def _init_setting_widgets(self) -> None:
        """创建两种排布共用的设置控件（只建变量与共享容器，不负责摆放）。"""
        self.range_label_var = tk.StringVar(
            value=RANGE_LABELS.get(str(self.vars["range_mode"].get()), "全部页面"))
        self.range_error_var = tk.StringVar()
        self.flow_note_var = tk.StringVar()
        self.advanced_summary_var = tk.StringVar(value="高级设置：全部为默认值")
        self.manual_duplex_var = self.vars["manual_duplex"]

    def _make_printer_widgets(self, parent, compact: bool) -> dict:
        combo = ttk.Combobox(parent, textvariable=self.printer_var, state="readonly",
                             width=16 if compact else 20)
        refresh = ttk.Button(parent, text="刷新", width=4, command=self._refresh_printers)
        return {"combo": combo, "refresh": refresh}

    def _apply_grid_minwidths(self, frame: ttk.Frame, columns, *, margin: int = 8) -> None:
        """按实际字体/控件请求宽度给栅格列设最小宽度（额外留 margin 像素）。

        用于替代 pack：列宽由每列最宽的子控件决定，任何缩放档位都不会
        让同一行的控件互相压住；额外余量避免“刚好卡边”。
        """
        gap = self._px(margin)
        # 新建的 Frame 在第一次 idle 之前 winfo_reqwidth() 还是 1，
        # 必须先让 Tk 完成几何计算，否则嵌套框（如输出方式单选组）
        # 会被当成 1px 宽，列最小宽度就失去意义。
        try:
            frame.update_idletasks()
        except tk.TclError:
            pass
        for column in columns:
            request = 0
            for child in frame.winfo_children():
                info = child.grid_info()
                try:
                    if int(info.get("column", "-1")) == column:
                        request = max(request, int(child.winfo_reqwidth()))
                except (TypeError, ValueError):
                    continue
            frame.grid_columnconfigure(column, minsize=request + gap)

    def _build_settings_strip(self, box: ttk.Frame) -> None:
        """默认设置条：显式多段栅格，禁止 pack 挤压。

        4 段布局（1024×720 @150% 仍不重叠、不裁字）：
          1. 打印机 / 份数 / 输出方式
          2. 打印流程三段单选 + 高级设置（窄窗口自动拆成两行）
          3. 页面范围 / 按份执行 / 范围错误 / 流程说明
        每列用 grid_columnconfigure(minsize=实测请求宽度+8px)，不依赖估算。
        """
        self._init_setting_widgets()
        self.settings_box = box
        box.columnconfigure(0, weight=0)
        box.columnconfigure(1, weight=1)

        line1 = ttk.Frame(box)
        line1.grid(row=0, column=0, columnspan=2, sticky="ew")
        printer = self._make_printer_widgets(line1, compact=True)
        self.printer_combo = printer["combo"]
        self.refresh_btn = printer["refresh"]
        ttk.Label(line1, text="打印机", width=5, anchor="e").grid(row=0, column=0, sticky="e")
        self.printer_combo.grid(row=0, column=1, sticky="w")
        self.refresh_btn.grid(row=0, column=2)

        ttk.Label(line1, text="份数", width=3, anchor="e").grid(row=0, column=3, sticky="e")
        # width=5：至少完整显示 4 位数字（99 份 + 右侧箭头留白）；
        # 专用 Copies.TSpinbox 加大内边距，避免数字贴住/被上下箭头压住。
        self.copies_spin = ttk.Spinbox(line1, from_=1, to=99, width=5,
                                       style="Copies.TSpinbox",
                                       textvariable=self.vars["copies"],
                                       command=self._on_setting_change)
        self.copies_spin.grid(row=0, column=4)
        self.copies_spin.bind("<KeyRelease>", lambda _e: self._on_setting_change())

        ttk.Label(line1, text="输出方式", width=8, anchor="e").grid(row=0, column=5, sticky="e")
        out = ttk.Frame(line1)
        out.grid(row=0, column=6, sticky="w")
        self.single_radio = ttk.Radiobutton(
            out, text="单面", value="single", variable=self.duplex_mode_var,
            command=self._on_duplex_mode_change)
        self.single_radio.grid(row=0, column=0, sticky="w")
        self.manual_duplex_check = ttk.Radiobutton(
            out, text="手动双面", value="duplex", variable=self.duplex_mode_var,
            command=self._on_duplex_mode_change)
        self.manual_duplex_check.grid(row=0, column=1, sticky="w", padx=(self._px(GAP), 0))
        self._apply_grid_minwidths(line1, range(7))

        # 第二行：打印流程三选一（替代旧的“批次翻面 + 每文件暂停”组合）
        flow_row = ttk.Frame(box)
        flow_row.grid(row=1, column=0, sticky="w", pady=(self._px(1), 0))
        self.settings_flow_row = flow_row
        ttk.Label(flow_row, text="打印流程", width=8, anchor="e").grid(
            row=0, column=0, sticky="e")
        self.flow_batch_all_radio = ttk.Radiobutton(
            flow_row, text="整批翻面 · 打完全部再取纸",
            value="batch_all", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_batch_all_radio.grid(row=0, column=1, sticky="w")
        self.flow_batch_take_radio = ttk.Radiobutton(
            flow_row, text="整批翻面 · 逐文件取纸",
            value="batch_take", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_batch_take_radio.grid(row=0, column=2, sticky="w")
        self.flow_per_file_radio = ttk.Radiobutton(
            flow_row, text="逐文件翻面 · 逐文件取纸",
            value="per_file", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_per_file_radio.grid(row=0, column=3, sticky="w")
        self._apply_grid_minwidths(flow_row, range(4))

        # 第三行：页面范围 + 按份执行 + 范围/流程提示（错误列弹性但不压别人）
        line3 = ttk.Frame(box)
        line3.grid(row=2, column=0, sticky="ew", pady=(self._px(1), 0))
        ttk.Label(line3, text="页面范围", width=8, anchor="e").grid(
            row=0, column=0, sticky="e")
        self.range_mode_combo = ttk.Combobox(
            line3, textvariable=self.range_label_var, state="readonly", width=8,
            values=tuple(RANGE_LABELS.values()))
        self.range_mode_combo.grid(row=0, column=1, sticky="w")
        self.range_mode_combo.bind("<<ComboboxSelected>>", self._on_range_mode_selected)
        self.range_entry = ttk.Entry(line3, textvariable=self.vars["range_text"], width=9)
        self.range_entry.grid(row=0, column=2, sticky="w")
        self.range_entry.bind("<KeyRelease>", lambda _e: self._on_setting_change())
        self.copy_exec_check = ttk.Checkbutton(
            line3, text="按份执行", variable=self.vars["copy_exec_on"],
            command=self._on_setting_change)
        self.copy_exec_check.grid(row=0, column=3, sticky="w")
        self.range_error_label = ttk.Label(line3, textvariable=self.range_error_var,
                                           foreground=self.colors["danger"], anchor="w")
        self.range_error_label.grid(row=0, column=4, sticky="ew", padx=(self._px(GAP), 0))
        # 流程说明与范围错误同一行（都在右侧信息区），不再单独占一行：
        # 720 高度下多一行会把文件列表压到 5 行以下。
        note_text = "份数>1 需按份执行才可按文件取纸"
        self.flow_note_label = ttk.Label(line3, text=note_text, style="Muted.TLabel",
                                         anchor="w")
        self.flow_note_label.grid(row=0, column=5, sticky="w", padx=(self._px(GAP), 0))
        # 先用完整文案量出列宽，再切回空变量；列最小宽度会保留说明文字的空间。
        self._apply_grid_minwidths(line3, (0, 1, 2, 3, 5))
        self.flow_note_label.configure(textvariable=self.flow_note_var,
                                       text="")
        line3.columnconfigure(4, weight=1)

        # 第四段：高级设置摘要与操作按钮。默认与打印流程同行右对齐；
        # 窗口较窄（如 125%/150% 下的 1024）自动落到独立一行，
        # 避免挤占流程单选的宽度，也避免在 720 高度下白白多占一行。
        line4 = ttk.Frame(box)
        line4.grid(row=1, column=1, sticky="e", padx=(self._px(GAP), 0),
                   pady=(self._px(1), 0))
        self.settings_advanced_row = line4
        line4.columnconfigure(0, weight=1)
        self.advanced_summary_label = ttk.Label(
            line4, textvariable=self.advanced_summary_var, style="Muted.TLabel",
            anchor="e")
        self.advanced_summary_label.grid(row=0, column=1, sticky="e")
        self.advanced_btn = ttk.Button(line4, text="高级设置…", command=self.open_advanced_settings)
        self.advanced_btn.grid(row=0, column=2, sticky="e")
        self.restore_btn = ttk.Button(
            line4, text="恢复默认", style="Ghost.TButton", command=self.restore_defaults)
        self.restore_btn.grid(row=0, column=3, sticky="e")
        self._apply_grid_minwidths(line4, (1, 2, 3))

        # 流程说明列固定在范围行右侧；这里只切换文案，不增删行高。
        box.bind("<Configure>", self._reflow_settings_strip, add="+")
        self._strip_layout_key = None
        # 构建期间 reqwidth 可能尚未稳定：等布局空闲后再判断一次同行/拆行。
        box.after_idle(self._reflow_settings_strip)
        # “恢复 Acrobat 等效默认（单面）”放在高级设置底部：它是低频的基线重置，
        # 且 Acrobat 基线本身是单面输出，放在设置条里容易被误读成本工具默认。
        # 首次启动与“恢复默认”的默认输出方式始终是手动双面。

    def _reflow_settings_strip(self, event=None) -> None:
        """按实际可用宽度决定“流程 + 高级设置”同行还是拆行。"""
        box = getattr(self, "settings_box", None)
        flow_row = getattr(self, "settings_flow_row", None)
        line4 = getattr(self, "settings_advanced_row", None)
        if box is None or flow_row is None or line4 is None:
            return
        width = int(event.width) if event is not None else box.winfo_width()
        if width <= 1:
            return
        gap = self._px(GAP * 2)
        # 请求宽度会随字体/文案变化（首次构建时可能还是小值），因此缓存键只记录
        # “是否同行”，不记录宽度——否则同一宽度下 reqwidth 变大后不会重新判断，
        # 右侧摘要会被挤压截断（125%/150% 的 1024 窗口就踩过这个坑）。
        required = flow_row.winfo_reqwidth() + line4.winfo_reqwidth() + gap
        side_by_side = required + self._px(8) <= width
        if side_by_side == self._strip_layout_key:
            return
        self._strip_layout_key = side_by_side
        if side_by_side:
            flow_row.grid_configure(row=1, column=0, columnspan=1, sticky="w")
            line4.grid_configure(row=1, column=1, columnspan=1, sticky="e",
                                 padx=(self._px(GAP), 0))
        else:
            flow_row.grid_configure(row=1, column=0, columnspan=2, sticky="w")
            line4.grid_configure(row=3, column=0, columnspan=2, sticky="ew",
                                 padx=0)

    def _build_settings_form(self, box: ttk.Frame) -> None:
        """方案 1：右侧竖排表单，标签统一 8 字符宽并右对齐。"""
        self._init_setting_widgets()
        box.columnconfigure(1, weight=1)

        def label(row: int, text: str) -> None:
            ttk.Label(box, text=text, width=8, anchor="e").grid(
                row=row, column=0, sticky="e", padx=(0, self._px(GAP)), pady=(0, self._px(GAP)))

        label(0, "打印机")
        printer_row = ttk.Frame(box)
        printer_row.grid(row=0, column=1, sticky="ew", pady=(0, self._px(GAP)))
        printer_row.columnconfigure(0, weight=1)
        printer = self._make_printer_widgets(printer_row, compact=False)
        self.printer_combo = printer["combo"]
        self.refresh_btn = printer["refresh"]
        self.printer_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.refresh_btn.grid(row=0, column=1)

        label(1, "份数")
        self.copies_spin = ttk.Spinbox(box, from_=1, to=99, width=5,
                                       style="Copies.TSpinbox",
                                       textvariable=self.vars["copies"],
                                       command=self._on_setting_change)
        self.copies_spin.grid(row=1, column=1, sticky="w", pady=(0, self._px(GAP)))
        self.copies_spin.bind("<KeyRelease>", lambda _e: self._on_setting_change())

        label(2, "页面范围")
        range_row = ttk.Frame(box)
        range_row.grid(row=2, column=1, sticky="w", pady=(0, self._px(GAP)))
        self.range_mode_combo = ttk.Combobox(
            range_row, textvariable=self.range_label_var, state="readonly", width=9,
            values=tuple(RANGE_LABELS.values()))
        self.range_mode_combo.grid(row=0, column=0, padx=(0, 6))
        self.range_mode_combo.bind("<<ComboboxSelected>>", self._on_range_mode_selected)
        self.range_entry = ttk.Entry(range_row, textvariable=self.vars["range_text"], width=10)
        self.range_entry.grid(row=0, column=1, sticky="w")
        self.range_entry.bind("<KeyRelease>", lambda _e: self._on_setting_change())

        self.range_error_label = ttk.Label(box, textvariable=self.range_error_var,
                                           foreground=self.colors["danger"], anchor="w")
        self.range_error_label.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, self._px(GAP)))

        label(4, "输出方式")
        out = ttk.Frame(box)
        out.grid(row=4, column=1, sticky="w", pady=(0, self._px(GAP)))
        self.single_radio = ttk.Radiobutton(
            out, text="单面", value="single", variable=self.duplex_mode_var,
            command=self._on_duplex_mode_change)
        self.single_radio.pack(side="left")
        self.manual_duplex_check = ttk.Radiobutton(
            out, text="手动双面", value="duplex", variable=self.duplex_mode_var,
            command=self._on_duplex_mode_change)
        self.manual_duplex_check.pack(side="left", padx=(self._px(GAP), 0))

        label(5, "打印流程")
        flip = ttk.Frame(box)
        flip.grid(row=5, column=1, sticky="w", pady=(0, self._px(GAP)))
        self.flow_batch_all_radio = ttk.Radiobutton(
            flip, text="整批翻面 · 打完全部再取纸",
            value="batch_all", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_batch_all_radio.grid(row=0, column=0, sticky="w")
        self.flow_batch_take_radio = ttk.Radiobutton(
            flip, text="整批翻面 · 逐文件取纸",
            value="batch_take", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_batch_take_radio.grid(row=1, column=0, sticky="w")
        self.flow_per_file_radio = ttk.Radiobutton(
            flip, text="逐文件翻面 · 逐文件取纸",
            value="per_file", variable=self.vars["print_flow"],
            command=self._on_setting_change)
        self.flow_per_file_radio.grid(row=2, column=0, sticky="w")
        self.copy_exec_check = ttk.Checkbutton(
            flip, text="按份执行", variable=self.vars["copy_exec_on"],
            command=self._on_setting_change)
        self.copy_exec_check.grid(row=3, column=0, sticky="w", pady=(self._px(GAP_S), 0))

        buttons = ttk.Frame(box)
        buttons.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(self._px(GAP), 0))
        self.advanced_btn = ttk.Button(buttons, text="高级设置…", command=self.open_advanced_settings)
        self.advanced_btn.pack(side="left")
        self.restore_btn = ttk.Button(
            buttons, text="恢复默认", style="Ghost.TButton", command=self.restore_defaults)
        self.restore_btn.pack(side="right")
        self.restore_acrobat_btn = ttk.Button(
            buttons, text="恢复 Acrobat 默认（单面）", style="Ghost.TButton",
            command=self.restore_acrobat_defaults)
        self.restore_acrobat_btn.pack(side="right", padx=(0, 6))

        ttk.Label(box, textvariable=self.advanced_summary_var, style="Muted.TLabel",
                  wraplength=380, justify="left").grid(
            row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))

    # 兼容旧命名（外部脚本可能调用）
    def _build_quick_settings(self, box: ttk.Frame) -> None:
        self._build_settings_form(box)

    def _on_duplex_mode_change(self) -> None:
        """单面/手动双面单选与 BooleanVar 的双向同步。"""
        duplex = str(self.duplex_mode_var.get()) == "duplex"
        self.vars["manual_duplex"].set(duplex)
        self._sync_widget_states()
        self._schedule_live_refresh()

    def _on_range_mode_selected(self, _event=None) -> None:
        """范围下拉（中文显示）→ 内部枚举值。"""
        label = str(self.range_label_var.get())
        value = RANGE_LABEL_TO_VALUE.get(label)
        if value:
            self.vars["range_mode"].set(value)
        self._sync_widget_states()
        self._schedule_live_refresh()

    def _on_scale_mode_selected(self, _event=None) -> None:
        """缩放下拉（中文显示）→ 内部枚举值，改动即时重绘。"""
        label = str(self.scale_label_var.get())
        value = SCALE_LABEL_TO_VALUE.get(label)
        if value:
            self.vars["scale_mode"].set(value)
        self._sync_widget_states()
        self._schedule_live_refresh()

    def _on_preview_frame_configure(self, event=None) -> None:
        width = event.width if event is not None else self.preview_frame.winfo_width()
        self._layout_preview_tool(width)

    def _layout_preview_tool(self, width: int) -> None:
        """极窄预览列（<510px）才把翻页/缩放拆成两行。

        1024 宽主界面预览列内宽约 525px，缩短下拉与文案后单行即可放下，
        这样工具条高度固定，缩放徽标/翻页控件不会随窗口尺寸跳动。
        """
        if not hasattr(self, "page_group"):
            return
        if width < self._px(510):
            self.tool_sep.grid_remove()
            self.zoom_group.grid_configure(row=1, column=0, sticky="w",
                                           pady=(self._px(GAP_S), 0))
        else:
            self.zoom_group.grid_configure(row=0, column=2, sticky="w", pady=0)
            self.tool_sep.grid()
        # 状态条宽度不足时隐藏快捷键提示（信息随窗口宽度增减，不推挤其它段）
        if hasattr(self, "shortcut_hint_label"):
            try:
                strip_width = self.preview_status_strip.winfo_width()
                if strip_width and strip_width < self._px(560):
                    self.shortcut_hint_label.pack_forget()
                elif not self.shortcut_hint_label.winfo_manager():
                    self.shortcut_hint_label.pack(side="right", padx=self._px(8))
            except tk.TclError:
                pass

    def _on_preview_view_change(self) -> None:
        """单页/双页并排切换：立即重绘（不重新打开 PDF）。"""
        self._sync_widget_states()
        self._draw_paper_preview()

    def _on_preview_canvas_configure(self, event=None) -> None:
        """画布尺寸变化：先按 4:3 上限收敛宽度，再重绘纸张。

        宽度上限 = max(420px, 高度 × 4/3)，并受预览卡片内宽限制；sticky 不含
        ew，因此收敛后的画布在卡片内水平居中，1024 宽下不再是大扁条。
        收敛会再触发一次 Configure，此时宽度已匹配，只重绘不递归。
        """
        if not self.shutting_down:
            try:
                width = event.width if event is not None else self.preview_canvas.winfo_width()
                height = event.height if event is not None else self.preview_canvas.winfo_height()
                frame_w = self.preview_frame.winfo_width()
                if height > 1 and frame_w > 1:
                    desired = min(frame_w, max(self._px(420), int(height * 4 / 3)))
                    if abs(desired - width) >= self._px(3):
                        self.preview_canvas.configure(width=desired)
                        return
            except tk.TclError:
                pass
        self._draw_paper_preview()

    def _on_preview_page_jump(self) -> None:
        """按输入框页码跳转（越界夹取）。"""
        try:
            target = int(float(str(self.preview_page_entry_var.get()).strip()))
        except ValueError:
            target = self.preview_page_index + 1
        self._show_preview_page(max(0, min(target - 1, max(0, self.preview_pages - 1))))


    # --------------------------------------------------------- 高级设置内容
    def _choice_combo(self, parent, var: tk.StringVar, labels: Dict[str, str],
                      width: int = 8) -> ttk.Combobox:
        """中文显示 / 英文存值的只读下拉（保证界面不出现英文枚举）。"""
        display = tk.StringVar(value=labels.get(var.get(), str(var.get())))
        combo = ttk.Combobox(parent, state="readonly", width=width,
                             textvariable=display, values=tuple(labels.values()))
        if not hasattr(self, "choice_combos"):
            self.choice_combos = []
        self.choice_combos.append((var, display, labels))

        def on_select(_event=None) -> None:
            chosen = display.get()
            for value, label in labels.items():
                if label == chosen:
                    var.set(value)
                    break
            self._on_setting_change()

        combo.bind("<<ComboboxSelected>>", on_select)
        return combo

    def _refresh_choice_combos(self) -> None:
        for var, display, labels in getattr(self, "choice_combos", []):
            display.set(labels.get(var.get(), str(var.get())))

    def _build_advanced_content(self, frame: ttk.Frame) -> None:
        """高级设置：版式 / 纸张 / 图像质量 / 输出与清理 四组 + “其它”。

        每个保留项都对应真实输出（审计表见本轮汇报）；
        界面全部显示中文，枚举值仍以英文保存，旧配置保持兼容。
        """
        frame.columnconfigure(0, weight=1, uniform="adv")
        frame.columnconfigure(1, weight=1, uniform="adv")

        # ---- 同源版式预览：直接复用主预览的几何数据（preview_metrics） ----
        preview_box = ttk.LabelFrame(frame, text=" 版式预览（与主界面同一数据源） ",
                                     padding=(8, 6))
        preview_box.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.adv_preview_canvas = tk.Canvas(
            preview_box, height=self._px(118), background=self.colors["surface_alt"],
            highlightthickness=1, highlightbackground=self.colors["border"])
        self.adv_preview_canvas.pack(fill="x")
        self.adv_preview_caption_var = tk.StringVar(
            value="添加文件后，这里显示与主预览相同的纸张、槽位与缩放结果。")
        ttk.Label(preview_box, textvariable=self.adv_preview_caption_var,
                  foreground=self.colors["muted"]).pack(anchor="w", pady=(4, 0))
        self.adv_preview_canvas.bind(
            "<Configure>", lambda _e: self._draw_advanced_preview())

        # ---- 组 1：版式 ----
        layout_box = ttk.LabelFrame(frame, text=" 版式 ", padding=(8, 6))
        layout_box.grid(row=1, column=0, sticky="nsew", pady=(0, 8), padx=(0, 6))
        ttk.Label(layout_box, text="每张页数：").grid(row=0, column=0, sticky="w")
        self.nup_combo = ttk.Combobox(
            layout_box, textvariable=self.vars["nup"], state="readonly", width=6,
            values=tuple(str(n) for n in NUP_CHOICES))
        self.nup_combo.grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.nup_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_setting_change())
        ttk.Label(layout_box, text="页序：").grid(row=0, column=2, sticky="w")
        self.order_combo = self._choice_combo(
            layout_box, self.vars["page_order"], PAGE_ORDER_LABELS, 9)
        self.order_combo.grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(layout_box, text="显示页边框", variable=self.vars["nup_border"],
                        command=self._on_setting_change).grid(
            row=1, column=1, sticky="w", pady=(6, 0))
        self.booklet_check = ttk.Checkbutton(
            layout_box, text="小册子排版", variable=self.vars["booklet"],
            command=self._sync_widgets)
        self.booklet_check.grid(row=1, column=2, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(layout_box, text="页边距（毫米）", style="Muted.TLabel").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.margin_entries = {}
        margin_fields = (("margin_left_mm", "左"), ("margin_top_mm", "上"),
                         ("margin_right_mm", "右"), ("margin_bottom_mm", "下"))
        for index, (key, label) in enumerate(margin_fields):
            row, col = 3 + index // 2, (index % 2) * 3
            ttk.Label(layout_box, text=f"{label}：").grid(row=row, column=col, sticky="w")
            entry = ttk.Entry(layout_box, textvariable=self.vars[key], width=6)
            entry.grid(row=row, column=col + 1, sticky="w", padx=(0, 8), pady=2)
            entry.bind("<KeyRelease>", lambda _e: self._on_setting_change())
            self.margin_entries[key] = entry

        ttk.Label(layout_box, text="内容裁切（毫米）", style="Muted.TLabel").grid(
            row=5, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.crop_entries = {}
        crop_fields = (("crop_left_mm", "左"), ("crop_top_mm", "上"),
                       ("crop_right_mm", "右"), ("crop_bottom_mm", "下"))
        for index, (key, label) in enumerate(crop_fields):
            row, col = 6 + index // 2, (index % 2) * 3
            ttk.Label(layout_box, text=f"{label}：").grid(row=row, column=col, sticky="w")
            entry = ttk.Entry(layout_box, textvariable=self.vars[key], width=6)
            entry.grid(row=row, column=col + 1, sticky="w", padx=(0, 8), pady=2)
            entry.bind("<KeyRelease>", lambda _e: self._on_setting_change())
            self.crop_entries[key] = entry
        ttk.Label(layout_box,
                  text="边距只做版面留白；裁切会真正去掉页面四周内容（如扫描阴影）。\n"
                       "小册子固定每张 2 页，并与手动双面配合：先正序打正面、整叠翻面后逆序打背面。",
                  foreground=self.colors["muted"], justify="left",
                  wraplength=480).grid(row=8, column=0, columnspan=4, sticky="w",
                                       pady=(8, 0))

        # ---- 组 2：纸张 ----
        paper_box = ttk.LabelFrame(frame, text=" 纸张 ", padding=(8, 6))
        paper_box.grid(row=1, column=1, sticky="nsew", pady=(0, 8), padx=(6, 0))
        paper_box.columnconfigure(3, weight=1)
        ttk.Label(paper_box, text="尺寸：").grid(row=0, column=0, sticky="w")
        self.paper_combo = ttk.Combobox(
            paper_box, textvariable=self.vars["paper_name"], state="readonly", width=11,
            values=tuple(PAPER_SIZES_MM.keys()) + (CUSTOM_PAPER_NAME,))
        self.paper_combo.grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.paper_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_widgets())

        custom_row = ttk.Frame(paper_box)
        custom_row.grid(row=0, column=2, columnspan=2, sticky="w")
        ttk.Label(custom_row, text="自定义 宽").pack(side="left")
        self.custom_paper_w_entry = ttk.Entry(
            custom_row, textvariable=self.vars["paper_width_mm"], width=6)
        self.custom_paper_w_entry.pack(side="left", padx=(3, 0))
        ttk.Label(custom_row, text="高").pack(side="left", padx=(6, 0))
        self.custom_paper_h_entry = ttk.Entry(
            custom_row, textvariable=self.vars["paper_height_mm"], width=6)
        self.custom_paper_h_entry.pack(side="left", padx=(3, 0))
        ttk.Label(custom_row, text="毫米").pack(side="left", padx=(3, 0))

        ttk.Label(paper_box, text="纸张来源：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.source_combo = ttk.Combobox(
            paper_box, textvariable=self.vars["source_name"], state="readonly", width=26)
        self.source_combo.grid(row=1, column=1, columnspan=3, sticky="w", pady=(6, 0))

        direction_row = ttk.Frame(paper_box)
        direction_row.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Label(direction_row, text="方向：").pack(side="left")
        for value, label in (("auto", "自动"), ("portrait", "纵向"), ("landscape", "横向")):
            ttk.Radiobutton(direction_row, text=label, value=value,
                            variable=self.vars["orientation"],
                            command=self._on_setting_change).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(paper_box, text="自动旋转（横竖页自动转 90° 适配槽位）",
                        variable=self.vars["auto_rotate"],
                        command=self._on_setting_change).grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Label(paper_box,
                  text="纸盒名称由打印机驱动提供，驱动不支持时自动回退默认纸盒。\n"
                       "纸张尺寸/来源/方向通过打印机驱动生效。",
                  foreground=self.colors["muted"], justify="left",
                  wraplength=480).grid(row=4, column=0, columnspan=4, sticky="w",
                                       pady=(8, 0))

        # ---- 组 3：图像质量 ----
        quality_box = ttk.LabelFrame(frame, text=" 图像质量 ", padding=(8, 6))
        quality_box.grid(row=2, column=0, sticky="nsew", pady=(0, 8), padx=(0, 6))
        ttk.Label(quality_box, text="颜色：").grid(row=0, column=0, sticky="w")
        self.color_combo = self._choice_combo(
            quality_box, self.vars["color_mode"], COLOR_LABELS, 8)
        self.color_combo.grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Checkbutton(quality_box, text="省墨模式", variable=self.vars["ink_saver"],
                        command=self._on_setting_change).grid(row=0, column=2, sticky="w")
        ttk.Label(quality_box, text="质量：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.quality_combo = self._choice_combo(
            quality_box, self.vars["print_quality"], QUALITY_LABELS, 8)
        self.quality_combo.grid(row=1, column=1, sticky="w", pady=(6, 0))
        ttk.Label(quality_box, text="分辨率：").grid(row=1, column=2, sticky="w",
                                                   pady=(6, 0), padx=(10, 0))
        ttk.Combobox(quality_box, textvariable=self.vars["dpi"], state="readonly", width=6,
                     values=("150", "200", "300", "600")).grid(
            row=1, column=3, sticky="w", pady=(6, 0))
        ttk.Checkbutton(quality_box, text="打印批注/表单/戳记（默认不打印）",
                        variable=self.vars["annots"], command=self._on_setting_change).grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Label(quality_box,
                  text="颜色/质量/省墨/分辨率通过驱动设置生效；批注按页面内容渲染。",
                  foreground=self.colors["muted"], justify="left", wraplength=480).grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(8, 0))

        # ---- 组 4：输出与清理 ----
        process_box = ttk.LabelFrame(frame, text=" 输出与清理 ", padding=(8, 6))
        process_box.grid(row=2, column=1, sticky="nsew", pady=(0, 8), padx=(6, 0))
        process_box.columnconfigure(1, weight=1)
        ttk.Checkbutton(process_box, text="逆序打印",
                        variable=self.vars["reverse"], command=self._on_setting_change).grid(
            row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Checkbutton(process_box, text="忽略空白页",
                        variable=self.vars["skip_blank"], command=self._on_setting_change).grid(
            row=0, column=1, sticky="w")
        ttk.Label(process_box, text="分页：", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        collate_row = ttk.Frame(process_box)
        collate_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Radiobutton(collate_row, text="自动分页（整份连续）", value="collate",
                        variable=self.vars["collate"],
                        command=self._on_setting_change).pack(side="left")
        ttk.Radiobutton(collate_row, text="逐页重复（不分页）", value="uncollated",
                        variable=self.vars["collate"],
                        command=self._on_setting_change).pack(side="left", padx=(10, 0))
        self.per_file_check = ttk.Checkbutton(
            process_box, text="逐文件单设范围/份数",
            variable=self.vars["per_file"], command=self._on_setting_change)
        self.per_file_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Checkbutton(process_box, text="完成后打开输出目录",
                        variable=self.vars["auto_open"], command=self._on_setting_change).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Checkbutton(process_box, text="保留排版 PDF",
                        variable=self.vars["keep_output"],
                        command=self._on_setting_change).grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(process_box, text="输出目录：").grid(row=6, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(process_box, textvariable=self.vars["output_dir"]).grid(
            row=6, column=1, sticky="ew", pady=(6, 0))
        ttk.Button(process_box, text="浏览…", width=6,
                   command=self._browse_output_dir).grid(row=6, column=2, sticky="w",
                                                         pady=(6, 0))
        ttk.Label(process_box,
                  text="Word/WPS/图片的临时 PDF 在打印完成或退出时自动清理；\n"
                       "下次启动还会清理上次崩溃留下的残留目录。",
                  foreground=self.colors["muted"], justify="left", wraplength=480).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # ---- 其它（低频项折叠，避免主区域拥挤） ----
        self.other_toggle_btn = ttk.Button(
            frame, text="其它（完成判断与页面设置） ▾", style="Ghost.TButton",
            command=self._toggle_advanced_other)
        self.other_toggle_btn.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 4))
        self.other_frame = ttk.LabelFrame(frame, text=" 其它 ", padding=(8, 6))
        self.other_frame.grid(row=4, column=0, columnspan=2, sticky="ew")
        self.other_frame.grid_remove()
        ttk.Label(self.other_frame, text="完成判断：", style="Muted.TLabel").pack(side="left")
        ttk.Radiobutton(self.other_frame, text=WAIT_LABELS["queue"], value="queue",
                        variable=self.vars["wait_mode"],
                        command=self._sync_widgets).pack(side="left", padx=(4, 10))
        ttk.Radiobutton(self.other_frame, text="固定延时", value="delay",
                        variable=self.vars["wait_mode"],
                        command=self._sync_widgets).pack(side="left")
        self.delay_spin = ttk.Spinbox(self.other_frame, from_=3, to=120, width=5,
                                      textvariable=self.vars["delay_seconds"])
        self.delay_spin.pack(side="left", padx=(4, 0))
        ttk.Label(self.other_frame, text="秒").pack(side="left", padx=(3, 0))
        ttk.Button(self.other_frame, text="页面设置…", style="Ghost.TButton",
                   command=self.open_page_setup_dialog).pack(side="left", padx=(16, 0))
        ttk.Label(self.other_frame,
                  text="页面设置对话框与上方“纸张/边距”同源，用于快速填入常用值。",
                  foreground=self.colors["muted"]).pack(side="left", padx=(10, 0))

    def _toggle_advanced_other(self) -> None:
        if self.other_frame.winfo_manager():
            self.other_frame.grid_remove()
            self.other_toggle_btn.configure(text="其它（完成判断与页面设置） ▾")
        else:
            self.other_frame.grid()
            self.other_toggle_btn.configure(text="其它（完成判断与页面设置） ▴")

    def _draw_advanced_preview(self) -> None:
        """高级设置内嵌小预览：与主预览共用 plan + preview_metrics 数据源。

        只画几何（纸张/可打印区/槽位/缩放），不重复渲染位图，因此实时且轻量。
        """
        canvas = getattr(self, "adv_preview_canvas", None)
        if canvas is None:
            return
        try:
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            if width < 50:
                width, height = 420, int(canvas.cget("height")) or 118
        except tk.TclError:
            return
        c = self.colors
        canvas.delete("all")
        record = self._selected_record()
        if record is None:
            canvas.create_text(width / 2, height / 2, text="添加并选择文件后显示版式预览",
                               fill=c["muted"], font=_theme.FONT_SMALL)
            self.adv_preview_caption_var.set(
                "添加文件后，这里显示与主预览相同的纸张、槽位与缩放结果。")
            return
        try:
            abs_pages, _dropped, copies = self._resolve_record(record)
        except PdfError as exc:
            canvas.create_text(width / 2, height / 2,
                               text="该文件范围有误，修正后即可预览",
                               fill=c["danger"], font=_theme.FONT_SMALL)
            self.adv_preview_caption_var.set(f"范围有误：{str(exc).splitlines()[0]}")
            return
        if not abs_pages:
            canvas.create_text(width / 2, height / 2, text="当前范围没有可打印页",
                               fill=c["muted"], font=_theme.FONT_SMALL)
            self.adv_preview_caption_var.set("当前范围没有可打印页，打印时该文件会被跳过。")
            return
        layout = self._layout_options()
        plan = build_layout_plan(len(abs_pages), layout)
        metrics = self.preview_metrics_data or {}
        expected_nup = 2 if layout.booklet else max(1, plan.nup)
        paper = tuple(metrics.get("paper") or plan.paper_pt)
        if (not metrics or abs(paper[0] - plan.paper_pt[0]) > 0.5
                or int(metrics.get("nup") or 0) != expected_nup):
            metrics = preview_metrics(plan, layout, plan.paper_pt[0], plan.paper_pt[1])
            paper = tuple(metrics["paper"])
        paper_w, paper_h = paper
        margin = self._px(12)
        scale = min((width - margin * 2) / max(1.0, paper_w),
                    (height - margin * 2) / max(1.0, paper_h))
        draw_w, draw_h = paper_w * scale, paper_h * scale
        x0 = (width - draw_w) / 2
        y0 = (height - draw_h) / 2
        canvas.create_rectangle(x0, y0, x0 + draw_w, y0 + draw_h,
                                fill=c.get("surface", "#ffffff"), outline=c["border"],
                                width=1)
        printable = metrics.get("printable")
        if printable:
            px0, py0, px1, py1 = printable
            canvas.create_rectangle(x0 + px0 * scale, y0 + py0 * scale,
                                    x0 + px1 * scale, y0 + py1 * scale,
                                    outline=c["border"], dash=(2, 3))
        for index, (sx0, sy0, sx1, sy1) in enumerate(metrics.get("slots") or []):
            left = x0 + sx0 * scale
            top = y0 + sy0 * scale
            right = x0 + sx1 * scale
            bottom = y0 + sy1 * scale
            canvas.create_rectangle(left, top, right, bottom, outline=c["accent"],
                                    dash=(3, 3))
            canvas.create_text((left + right) / 2, (top + bottom) / 2,
                               text=str(index + 1), fill=c["accent"],
                               font=_theme.FONT_SMALL)
        percent = float(metrics.get("scale_percent") or 0.0)
        paper_text = f"{paper_w * 25.4 / 72:.0f}×{paper_h * 25.4 / 72:.0f} 毫米"
        self.adv_preview_caption_var.set(
            f"纸张 {paper_text} · 每张 {expected_nup} 页 · 实际缩放 {percent:.1f}% · "
            f"本次 {len(abs_pages)} 页 / 输出 {plan.output_count} 面 / {plan.output_count // 2} 张×{copies} 份")



    # -------------------------------------------------------------- 预览区
    def _build_preview_panel(self, frame: ttk.Frame) -> None:
        """预览区（方案 1/2 共用）：标题行 + 缩放行 + 翻页行 + Canvas + 缩略图条。

        缩放控件从“高级设置”移到预览工具条，改动后即时重绘（防抖 100ms）。
        不再单独放“文件汇总”卡片——它与文件列表内容重复，只会挤压画布宽度。
        """
        self.tab_preview = frame
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)        # Canvas 舞台

        # ---- 标题行：标题 + 实际缩放徽标 + 视图切换（与刷新用分隔线隔开） ----
        head = ttk.Frame(frame)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(2, weight=1)
        ttk.Label(head, text="打印预览", style="Subtitle.TLabel").grid(row=0, column=0, sticky="w")
        # 实际缩放百分比做成轻量徽标（浅底小胶囊），不再是一行纯文字
        self.actual_scale_var = tk.StringVar(value="实际 —")
        badge = tk.Frame(head, background=self.colors["surface_alt"],
                         highlightthickness=1, highlightbackground=self.colors["border"])
        badge.grid(row=0, column=1, sticky="w", padx=(self._px(GAP), self._px(GAP)))
        tk.Label(badge, textvariable=self.actual_scale_var, background=self.colors["surface_alt"],
                 foreground=self.colors["accent"], font=_theme.FONT_SMALL,
                 padx=self._px(6), pady=self._px(1)).pack()
        ttk.Label(head, text="视图", style="Muted.TLabel").grid(
            row=0, column=3, sticky="e", padx=(0, self._px(GAP_S)))
        ttk.Radiobutton(head, text="单页", value="single", variable=self.preview_view_var,
                        command=self._on_preview_view_change).grid(
            row=0, column=4, padx=(0, self._px(GAP_S)))
        ttk.Radiobutton(head, text="双页", value="spread", variable=self.preview_view_var,
                        command=self._on_preview_view_change).grid(row=0, column=5)
        ttk.Separator(head, orient="vertical", style="VSep.TSeparator").grid(
            row=0, column=6, sticky="ns", padx=self._px(GAP), pady=self._px(2))
        self.preview_refresh_btn = ttk.Button(head, text="刷新", width=4,
                                              command=self._refresh_layout_preview)
        self.preview_refresh_btn.grid(row=0, column=7)

        # ---- 工具条：翻页组 | 缩放组（组间 1px 细分隔线；窄窗口自动换行）------
        tool = ttk.Frame(frame)
        tool.grid(row=1, column=0, sticky="ew", pady=(self._px(GAP), 0))
        self.preview_tool = tool
        self.preview_frame = frame
        self.page_group = ttk.Frame(tool)
        self.page_group.grid(row=0, column=0, sticky="w")
        self.zoom_group = ttk.Frame(tool)
        self.zoom_group.grid(row=0, column=2, sticky="w")
        self.tool_sep = ttk.Separator(tool, orient="vertical", style="VSep.TSeparator")
        self.tool_sep.grid(row=0, column=1, sticky="ns", padx=self._px(GAP), pady=self._px(2))

        ttk.Label(self.page_group, text="页", width=2, anchor="e").grid(
            row=0, column=0, padx=(0, self._px(GAP_S)))
        self.preview_page_entry_var = tk.StringVar(value="1")
        self.preview_page_spin = ttk.Spinbox(
            self.page_group, from_=1, to=1, width=4,
            textvariable=self.preview_page_entry_var,
            command=self._on_preview_page_jump)
        self.preview_page_spin.grid(row=0, column=1)
        self.preview_page_spin.bind("<Return>", lambda _e: self._on_preview_page_jump())
        ttk.Label(self.page_group, textvariable=self.preview_page_var, style="Muted.TLabel").grid(
            row=0, column=2, sticky="w", padx=(self._px(GAP_S), self._px(GAP)))
        self.preview_prev_btn = ttk.Button(self.page_group, text="‹", width=2,
                                           command=lambda: self._set_preview_page(-1))
        self.preview_prev_btn.grid(row=0, column=3, padx=(0, self._px(GAP_S)))
        self.preview_next_btn = ttk.Button(self.page_group, text="›", width=2,
                                           command=lambda: self._set_preview_page(1))
        self.preview_next_btn.grid(row=0, column=4)

        # 缩放下拉显示中文，内部保存英文枚举值；改动立即重绘（100ms 防抖）
        self.scale_label_var = tk.StringVar(
            value=SCALE_LABELS.get(str(self.vars["scale_mode"].get()), SCALE_LABELS["shrink"]))
        self.scale_mode_combo = ttk.Combobox(
            self.zoom_group, textvariable=self.scale_label_var, state="readonly", width=12,
            values=tuple(SCALE_LABELS.values()))
        self.scale_mode_combo.grid(row=0, column=0, sticky="w", padx=(0, self._px(GAP)))
        self.scale_mode_combo.bind("<<ComboboxSelected>>", self._on_scale_mode_selected)
        self.scale_percent_spin = ttk.Spinbox(
            self.zoom_group, from_=10, to=500, width=5, textvariable=self.vars["scale_percent"],
            command=self._on_setting_change)
        self.scale_percent_spin.grid(row=0, column=1)
        self.scale_percent_spin.bind("<KeyRelease>", lambda _e: self._on_setting_change())
        ttk.Label(self.zoom_group, text="%").grid(row=0, column=2, padx=(self._px(2), 0))

        # ---- 舞台：Canvas 画纸张/可打印区/内容 -----------------------------
        # 舞台最小高度 180px：窗口缩到最小时也不会被压成一条窄缝；
        # 正常尺寸下由 row weight 拉满剩余高度。宽度不强制撑满预览卡片：
        # 按 4:3 上限收缩并水平居中（sticky 不含 ew），避免 1024 宽时画布
        # 变成“扁横幅”，纸张图始终按可用空间等比居中。
        self.preview_canvas = tk.Canvas(
            frame, background=self.colors["surface_alt"], highlightthickness=1,
            highlightbackground=self.colors["border"], borderwidth=0,
            width=self._px(420), height=self._px(240))
        self.preview_canvas.grid(row=2, column=0, sticky="ns", pady=(self._px(GAP), 0))
        self.preview_canvas.bind("<Configure>", self._on_preview_canvas_configure)
        frame.rowconfigure(2, weight=1, minsize=self._px(180))

        # ---- 底部细状态条：页码 / 缩放 / 纸张（位置固定，长度变化不推挤）----
        strip = tk.Frame(frame, background=self.colors["surface_alt"],
                         height=self._px(24), highlightthickness=1,
                         highlightbackground=self.colors["border"])
        strip.grid(row=3, column=0, sticky="ew", pady=(self._px(GAP_S), 0))
        strip.grid_propagate(False)
        self.preview_status_strip = strip
        self.preview_status_page_var = tk.StringVar(value="第 — 面")
        self.preview_status_scale_var = tk.StringVar(value="缩放 —")
        self.preview_status_paper_var = tk.StringVar(value="纸张 —")
        self.preview_status_labels = []
        for index, var in enumerate((self.preview_status_page_var,
                                     self.preview_status_scale_var,
                                     self.preview_status_paper_var)):
            if index:
                sep = tk.Frame(strip, background=self.colors["border"], width=1,
                               height=self._px(12))
                sep.pack(side="left", padx=self._px(GAP), pady=self._px(6))
            label = tk.Label(strip, textvariable=var, background=self.colors["surface_alt"],
                             foreground=self.colors["muted"], font=_theme.FONT_SMALL,
                             anchor="w")
            label.pack(side="left", padx=(self._px(8) if index == 0 else 0, 0))
            self.preview_status_labels.append(label)
        self.shortcut_hint_label = tk.Label(
            strip, text="Ctrl+O 添加 · Del 移除 · Ctrl+Enter 打印",
            background=self.colors["surface_alt"], foreground=self.colors["muted"],
            font=_theme.FONT_SMALL)
        self.shortcut_hint_label.pack(side="right", padx=self._px(8))
        # 预览区宽度变化时切换工具条分组布局（窄窗口自动换成两行）
        frame.bind("<Configure>", self._on_preview_frame_configure, add="+")

        # ---- 预览说明（限宽换行，长文本不会撑宽预览列）--------------------
        # 矮窗口（如 1024×720）自动收起这一行，把高度让给画布；高窗口才显示。
        self.preview_info_var = tk.StringVar(value="")
        self.preview_info_label = ttk.Label(
            frame, textvariable=self.preview_info_var, style="Muted.TLabel",
            wraplength=420, justify="left", anchor="w")
        self.preview_info_label.grid(row=4, column=0, sticky="w", pady=(self._px(GAP_S), 0))
        self.preview_info_var.trace_add(
            "write", lambda *_args: self._sync_preview_info_visibility())

    def _sync_preview_info_visibility(self) -> None:
        """矮窗口收起状态条与说明行（信息已由标题徽标/工具条覆盖），把高度让给画布。

        1024×720 下预览画布因此多出约 30px，避免被压成扁条。
        """
        if not hasattr(self, "preview_info_label"):
            return
        try:
            tall = self.root.winfo_height() >= self._px(880)
        except tk.TclError:
            return
        if hasattr(self, "preview_status_strip"):
            if tall:
                self.preview_status_strip.grid()
            else:
                self.preview_status_strip.grid_remove()
        show_info = tall and bool(self.preview_info_var.get().strip())
        if show_info:
            self.preview_info_label.grid()
        else:
            self.preview_info_label.grid_remove()

    # ------------------------------------------------------------------ 底部常驻日志
    def _build_log_drawer(self, parent: tk.Frame) -> None:
        """底部常驻日志区（默认展开、始终可见）。

        结构：可拖动分隔条 → 标题条（日志 + 导出/清空/打开目录 + 页序详情 +
        最新一条摘要）→ 日志正文。高度按像素固定（默认 ≈6 行），拖动分隔条
        在最小值与窗口 1/3 之间调整；窗口变高时自动按比例增高，但不挤压
        工作区（工作区所在行仍是唯一 weight=1 的伸缩行）。
        """
        c = self.colors
        self._log_min_height = self._px(134)
        self._log_height = self._log_min_height
        self._log_user_sized = False
        self._log_drag = None
        self.log_drawer_open = True

        sash = tk.Frame(parent, background=c["bg"], height=self._px(6),
                        cursor="sb_v_double_arrow", highlightthickness=0)
        sash.grid(row=4, column=0, sticky="ew", pady=(self._px(GAP_S), 0))
        sash.bind("<Button-1>", self._on_log_sash_press)
        sash.bind("<B1-Motion>", self._on_log_sash_drag)
        sash.bind("<ButtonRelease-1>", lambda _e: setattr(self, "_log_user_sized", True))
        self.log_sash = sash

        drawer = tk.Frame(parent, background=c["surface"], highlightthickness=1,
                          highlightbackground=c["border"], height=self._log_height)
        drawer.grid(row=5, column=0, sticky="ew", pady=(self._px(GAP_S), 0))
        drawer.grid_propagate(False)
        drawer.columnconfigure(0, weight=1)
        drawer.rowconfigure(1, weight=1)
        self.log_drawer = drawer

        head = tk.Frame(drawer, background=c["surface"], padx=self._px(10), pady=self._px(2))
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(5, weight=1)
        tk.Label(head, text="日志", background=c["surface"], foreground=c["text"],
                 font=_theme.FONT_SUBTITLE).grid(row=0, column=0, sticky="w")
        self.export_log_btn = ttk.Button(head, text="导出", style="Ghost.TButton",
                                         command=self.export_log)
        self.export_log_btn.grid(row=0, column=1, padx=(self._px(GAP), 0))
        ttk.Button(head, text="清空", style="Ghost.TButton",
                   command=self.clear_log).grid(row=0, column=2, padx=(self._px(GAP_S), 0))
        ttk.Button(head, text="打开目录", style="Ghost.TButton",
                   command=self.open_log_dir).grid(row=0, column=3, padx=(self._px(GAP_S), 0))
        self.log_detail_var = tk.StringVar(value="页序/缩略图 ▸")
        ttk.Button(head, textvariable=self.log_detail_var, style="Ghost.TButton",
                   command=self._toggle_log_detail).grid(row=0, column=4, padx=(self._px(GAP_S), 0))
        # 最新一条始终可见（按实际可用像素省略，不裁切半个字）
        self.log_summary_var = tk.StringVar(value="暂无日志")
        self._log_summary_full = "暂无日志"
        summary = tk.Label(head, textvariable=self.log_summary_var, background=c["surface"],
                           foreground=c["muted"], font=_theme.FONT_SMALL, anchor="e")
        summary.grid(row=0, column=5, sticky="e", padx=(self._px(GAP), 0))
        self.log_summary_label = summary
        summary.bind("<Configure>", lambda _e: self._render_log_summary())

        body = tk.Frame(drawer, background=c["surface"], padx=self._px(10), pady=self._px(2))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self.log_body = body

        # 页序详情（整批页序摘要 + 双面示意图 + 缩略图）：默认收起，按需展开
        detail = tk.Frame(body, background=c["surface"])
        detail.grid(row=0, column=0, sticky="ns", padx=(0, self._px(GAP)))
        detail.grid_remove()
        self.log_detail_frame = detail
        self.batch_plan_var = tk.StringVar(value="")
        tk.Label(detail, textvariable=self.batch_plan_var, justify="left", anchor="w",
                 background=c["surface_alt"], foreground=c["text"],
                 font=_theme.FONT_SMALL, padx=8, pady=6, wraplength=420,
                 highlightthickness=1, highlightbackground=c["border"]).grid(
            row=0, column=0, sticky="ew", pady=(0, self._px(GAP)))
        duplex_box = ttk.Frame(detail)
        duplex_box.grid(row=1, column=0, sticky="w", pady=(0, self._px(GAP)))
        self.preview_title_var = tk.StringVar(
            value="添加文件后，这里会以示意图显示每张纸的正反面页码。")
        self.preview_title_label = ttk.Label(duplex_box, textvariable=self.preview_title_var,
                                             foreground=self.colors["accent"])
        self.preview_title_label.grid(row=0, column=0, sticky="w")
        self.preview_box_frame = tk.Frame(duplex_box, background=self.colors["surface"])
        self.preview_box_frame.grid(row=1, column=0, sticky="w", pady=(2, 0))
        thumbs_row = ttk.Frame(detail)
        thumbs_row.grid(row=2, column=0, sticky="w")
        ttk.Label(thumbs_row, text="缩略图", style="Muted.TLabel").pack(side="left", padx=(0, self._px(GAP)))
        self.thumb_frame = ttk.Frame(thumbs_row)
        self.thumb_frame.pack(side="left")

        self.log_text = tk.Text(body, height=5, wrap="word", state="disabled",
                                font=_theme.FONT_BODY,
                                background=self.colors["input_bg"],
                                foreground=self.colors["text"],
                                insertbackground=self.colors["text"],
                                relief="flat", borderwidth=1,
                                padx=self._px(10), pady=self._px(4),
                                highlightthickness=1,
                                highlightbackground=self.colors["border"])
        self.log_text.grid(row=0, column=1, sticky="nsew")
        # 错误/失败日志用危险色高亮（只加 tag，不改写文本，搜索/断言不受影响）
        self.log_text.tag_configure("log_error",
                                    foreground=self.colors["danger_text"])
        scroll = ttk.Scrollbar(body, orient="vertical", style="Thin.Vertical.TScrollbar",
                               command=self.log_text.yview)
        scroll.grid(row=0, column=2, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

    # ---------------------------------------------------------- 日志高度与详情
    def _on_log_sash_press(self, event) -> None:
        self._log_drag = (event.y_root, self._log_height)

    def _on_log_sash_drag(self, event) -> None:
        if not self._log_drag:
            return
        start_y, start_height = self._log_drag
        self._set_log_height(start_height - (event.y_root - start_y))

    def _set_log_height(self, value: int) -> None:
        """把日志高度夹在“最小值 … 窗口 1/3”之间。"""
        try:
            window_h = max(1, self.root.winfo_height())
        except tk.TclError:
            window_h = self._log_min_height * 3
        limit = max(self._log_min_height, window_h // 3)
        height = int(max(self._log_min_height, min(int(value), limit)))
        if height == self._log_height:
            return
        self._log_height = height
        if hasattr(self, "log_drawer"):
            self.log_drawer.configure(height=height)

    def _resize_log_for_window(self) -> None:
        """窗口变高时让日志轻微增高（≈16%），用户手动拖过之后不再自动改。"""
        if getattr(self, "_log_user_sized", False):
            return
        try:
            window_h = self.root.winfo_height()
        except tk.TclError:
            return
        self._set_log_height(max(self._log_min_height, int(window_h * 0.16)))

    def _toggle_log_detail(self) -> None:
        if self.log_detail_frame.winfo_ismapped():
            self.log_detail_frame.grid_remove()
            self.log_detail_var.set("页序/缩略图 ▸")
            return
        self.log_detail_frame.grid()
        self.log_detail_var.set("页序/缩略图 ▾")
        self._set_log_height(max(self._log_height, self._px(360)))

    def _toggle_log_drawer(self) -> None:
        """兼容旧入口：日志常驻，不再收起，仅把焦点放到日志正文。"""
        self.log_text.see("end")

    def open_log_dir(self) -> None:
        """打开配置文件所在目录（日志导出默认位置也在这里）。"""
        folder = os.path.dirname(settings_path())
        try:
            open_directory(folder)
            self.status_var.set(f"已打开目录：{folder}")
        except PrintError as exc:
            messagebox.showerror("打开目录失败", str(exc), parent=self.root)

    # ------------------------------------------------------------------ 主操作行
    def _build_action_row(self, parent: ttk.Frame, vertical: bool = False) -> None:
        """开始打印/暂停/取消 + 状态与进度（两种布局共用同一套控件）。"""
        c = self.colors
        row = ttk.Frame(parent)
        row.pack(fill="x")

        if not vertical:
            self.start_btn = ttk.Button(row, text="▶ 开始打印", width=12, style="Accent.TButton",
                                        command=self.start_print)
            self.start_btn.grid(row=0, column=0, sticky="w")
            self.pause_btn = ttk.Button(row, text="暂停", width=6, style="TButton", state="disabled",
                                        command=self.toggle_pause)
            self.pause_btn.grid(row=0, column=1, sticky="w", padx=(self._px(GAP), 0))
            self.stop_btn = ttk.Button(row, text="✕ 取消", width=6, style="Danger.TButton",
                                       state="disabled", command=self.stop_print)
            self.stop_btn.grid(row=0, column=2, sticky="w", padx=(self._px(GAP), 0))
            info = ttk.Frame(row)
            info.grid(row=0, column=3, sticky="ew", padx=(self._px(16), 0))
            # 状态文字占弹性列；阶段徽标/进度条/页码列固定，长度变化不会推挤它们
            info.columnconfigure(0, weight=1)
            info.columnconfigure(3, minsize=self._px(84))
            self.status_var = tk.StringVar(value="就绪")
            self._status_full = "就绪"
            self.status_label = tk.Label(info, text="", background=c["surface"],
                                         foreground=c["accent"],
                                         font=_theme.FONT_SUBTITLE, anchor="w")
            self.status_label.grid(row=0, column=0, sticky="w")
            self.status_label.bind("<Configure>", lambda _e: self._render_status_text())
            self.status_var.trace_add("write", self._on_status_change)
            self._render_status_text()
            phase_chip = tk.Frame(info, background=c["surface_alt"], width=self._px(112),
                                  height=self._px(24), highlightthickness=1,
                                  highlightbackground=c["border"])
            phase_chip.grid(row=0, column=1, sticky="e", padx=(self._px(GAP), self._px(GAP)))
            phase_chip.grid_propagate(False)
            self.phase_var = tk.StringVar(value="待开始")
            self.phase_label = tk.Label(phase_chip, textvariable=self.phase_var,
                                        background=c["surface_alt"],
                                        foreground=c["muted"], font=_theme.FONT_SMALL)
            self.phase_label.place(relx=0.5, rely=0.5, anchor="center")
            self.progress = ttk.Progressbar(info, mode="determinate", length=self._px(200),
                                            maximum=100)
            self.progress.grid(row=0, column=2, sticky="ew", padx=(0, self._px(12)))
            self.current_page_var = tk.StringVar(value="")
            tk.Label(info, textvariable=self.current_page_var, background=c["surface"],
                     foreground=c["muted"], font=_theme.FONT_SMALL, anchor="e").grid(
                row=0, column=3, sticky="e")
        else:
            info = ttk.Frame(row)
            info.pack(fill="x")
            info.columnconfigure(0, weight=1)
            self.status_var = tk.StringVar(value="就绪")
            self._status_full = "就绪"
            self.status_label = tk.Label(info, text="", background=c["surface"],
                                         foreground=c["accent"],
                                         font=_theme.FONT_SUBTITLE, anchor="w")
            self.status_label.grid(row=0, column=0, sticky="w")
            self.status_label.bind("<Configure>", lambda _e: self._render_status_text())
            self.status_var.trace_add("write", self._on_status_change)
            self._render_status_text()
            self.current_page_var = tk.StringVar(value="")
            tk.Label(info, textvariable=self.current_page_var, background=c["surface"],
                     foreground=c["muted"], font=_theme.FONT_SMALL, anchor="e").grid(
                row=0, column=1, sticky="e")
            self.progress = ttk.Progressbar(info, mode="determinate", maximum=100)
            self.progress.grid(row=1, column=0, columnspan=2, sticky="ew",
                               pady=(self._px(GAP_S), self._px(GAP)))
            btn_row = ttk.Frame(row)
            btn_row.pack(fill="x")
            self.start_btn = ttk.Button(btn_row, text="▶ 开始打印", style="Accent.TButton",
                                        command=self.start_print)
            self.start_btn.pack(side="left", fill="x", expand=True)
            self.pause_btn = ttk.Button(btn_row, text="暂停", width=6, style="TButton",
                                        state="disabled", command=self.toggle_pause)
            self.pause_btn.pack(side="left", padx=(self._px(GAP), 0))
            self.stop_btn = ttk.Button(btn_row, text="✕ 取消", width=6, style="Danger.TButton",
                                       state="disabled", command=self.stop_print)
            self.stop_btn.pack(side="left", padx=(self._px(GAP), 0))

    _STATUS_ERROR_WORDS = ("失败", "错误", "不可用", "未选择", "请先", "无法", "有误",
                           "未检测到", "不支持", "未就绪")
    _STATUS_WARN_WORDS = ("暂停", "取消")
    _STATUS_SUCCESS_WORDS = ("完成", "已清空", "已移除", "已继续", "准备就绪", "已就绪",
                             "已恢复", "已打开")

    def _on_status_change(self, *_args) -> None:
        """状态文字变化：按语义着色，并按可用宽度省略，避免长文案挤压进度条。"""
        self._status_full = str(self.status_var.get())
        self._render_status_text()
        self._update_flow_steps()

    def _status_tone(self, text: str) -> str:
        text = str(text or "")
        if any(word in text for word in self._STATUS_ERROR_WORDS):
            return "danger"
        if any(word in text for word in self._STATUS_WARN_WORDS):
            return "warning"
        if any(word in text for word in self._STATUS_SUCCESS_WORDS):
            return "success"
        return "accent"

    def _render_status_text(self) -> None:
        label = getattr(self, "status_label", None)
        if label is None:
            return
        full = str(getattr(self, "_status_full", ""))
        try:
            width = label.winfo_width()
            if width <= 1:
                width = self._px(320)
            font = tkfont.Font(root=self.root, font=label.cget("font"))
            text = self._elide_text(full, font, max(self._px(24), width - self._px(2)))
            key = {"danger": "danger_text", "warning": "warning",
                   "success": "success"}.get(self._status_tone(full), "accent")
            label.configure(text=text, foreground=self.colors[key])
        except tk.TclError:
            return

    def _set_phase_text(self, text: str, tone: str = "muted") -> None:
        if not hasattr(self, "phase_var"):
            return
        self.phase_var.set(text)
        label = getattr(self, "phase_label", None)
        if label is None:
            return
        key = {"accent": "accent", "success": "success", "warning": "warning",
               "danger": "danger_text"}.get(tone, "muted")
        try:
            label.configure(foreground=self.colors[key])
        except tk.TclError:
            pass

    def _update_phase_from_status(self, status: str) -> None:
        """把逐文件状态映射为固定的阶段徽标（转换/阶段一正序/阶段二逆序）。"""
        text = str(status or "")
        if "转换中" in text:
            self._set_phase_text("② 转换中", "warning")
        elif "待转换" in text:
            self._set_phase_text("② 待转换", "warning")
        elif "转换失败" in text or "无法读取" in text:
            self._set_phase_text("② 转换失败", "danger")
        elif any(word in text for word in ("整批·正面", "第1面", "·正面", "正序")):
            self._set_phase_text("③ 阶段一 · 正序", "accent")
        elif any(word in text for word in ("整批·反面", "第2面", "·反面", "逆序")):
            self._set_phase_text("③ 阶段二 · 逆序", "accent")
        elif "单面打印" in text:
            self._set_phase_text("③ 单面打印", "accent")
        elif "排版中" in text:
            self._set_phase_text("③ 正在排版", "accent")
        elif "等待翻面" in text:
            self._set_phase_text("③ 等待翻面", "warning")
        elif "已完成" in text:
            self._set_phase_text("③ 已完成", "success")
        elif "已取消" in text:
            self._set_phase_text("③ 已取消", "warning")
        elif "失败" in text:
            self._set_phase_text("③ 打印失败", "danger")

    def _setup_tab_order(self) -> None:
        """显式焦点链：Tab 顺序与视觉顺序一致（文件区 → 预览 → 设置 → 操作）。"""
        order = [
            getattr(self, "add_file_btn", None), getattr(self, "add_dir_btn", None),
            getattr(self, "remove_btn", None), getattr(self, "up_btn", None),
            getattr(self, "down_btn", None), getattr(self, "dedupe_btn", None),
            getattr(self, "explore_btn", None), getattr(self, "clear_btn", None),
            getattr(self, "recent_combo", None),
            getattr(self, "preview_page_spin", None), getattr(self, "preview_prev_btn", None),
            getattr(self, "preview_next_btn", None), getattr(self, "scale_mode_combo", None),
            getattr(self, "scale_percent_spin", None), getattr(self, "preview_refresh_btn", None),
            getattr(self, "printer_combo", None), getattr(self, "refresh_btn", None),
            getattr(self, "copies_spin", None), getattr(self, "range_mode_combo", None),
            getattr(self, "range_entry", None),
            getattr(self, "single_radio", None), getattr(self, "manual_duplex_check", None),
            getattr(self, "flow_batch_all_radio", None),
            getattr(self, "flow_batch_take_radio", None),
            getattr(self, "flow_per_file_radio", None),
            getattr(self, "copy_exec_check", None),
            getattr(self, "advanced_btn", None), getattr(self, "restore_btn", None),
            getattr(self, "restore_acrobat_btn", None),
            getattr(self, "start_btn", None), getattr(self, "pause_btn", None),
            getattr(self, "stop_btn", None),
        ]
        self._focus_order = [w for w in order if w is not None and w.winfo_exists()]
        # 只在本窗口内生效：高级设置/进度提示等 Toplevel 不受影响
        self.root.bind("<Tab>", lambda _e: self._cycle_focus(1))
        self.root.bind("<Shift-Tab>", lambda _e: self._cycle_focus(-1))

    def _cycle_focus(self, step: int):
        order = [w for w in getattr(self, "_focus_order", [])
                 if w.winfo_exists() and w.winfo_ismapped()
                 and str(w.cget("state")) != "disabled"]
        if not order:
            return None
        try:
            current = self.root.focus_get()
        except tk.TclError:
            current = None
        start = order.index(current) + step if current in order else (0 if step > 0 else -1)
        target = order[start % len(order)]
        target.focus_set()
        return "break"

    # 兼容旧命名（外部脚本可能调用）
    def _build_bottom_bar(self, parent: tk.Frame) -> None:
        self._build_action_row(parent, vertical=False)

    # -------------------------------------------------------------- 设置与联动
    def _on_setting_change(self) -> None:
        """任意设置变化：先立即同步控件可用状态，再防抖刷新预览。

        防抖（LIVE_REFRESH_DEBOUNCE_MS=100ms）只作用于“重算预览/页序”这类
        开销较大的工作，避免在范围输入框里每敲一个字符就重开 PDF；
        控件启用/禁用状态仍然即时生效。
        """
        self._sync_widget_states()
        self._update_scope_hint()
        self._schedule_live_refresh()

    def _schedule_live_refresh(self) -> None:
        if self.shutting_down:
            return
        if self._live_refresh_id is not None:
            try:
                self.root.after_cancel(self._live_refresh_id)
            except tk.TclError:
                pass
        self._live_refresh_id = self.root.after(
            LIVE_REFRESH_DEBOUNCE_MS, self._apply_live_refresh)

    def _apply_live_refresh(self) -> None:
        """防抖结束后的统一刷新入口：保证最后一次输入一定被执行。"""
        self._live_refresh_id = None
        if self.shutting_down:
            return
        try:
            self._sync_widget_states()
            self._update_advanced_summary()
            self._update_preview()
            # 窗口拉伸/设置变化后立即重绘成纸预览（用缓存位图，不重开 PDF）
            self._draw_paper_preview()
            self._draw_advanced_preview()
        except Exception:
            self._append_log("刷新预览失败：\n" + traceback.format_exc())

    def _sync_widgets(self) -> None:
        """立即同步控件状态并刷新预览（供测试与外部调用，不走防抖）。"""
        self._refresh_choice_combos()
        self._sync_widget_states()
        self._update_scope_hint()
        self._update_selection_bar()
        self._update_advanced_summary()
        self._update_preview()
        # 摘要文案变化会改变右侧区块的请求宽度：重新判断“同行/拆行”，
        # 避免控件被挤压后超出父容器被裁剪。
        self._reflow_settings_strip()

    def _mode_hint_text(self) -> str:
        """三段打印流程的停顿点，放在文件列表提示行常驻显示。"""
        if not bool(self.vars["manual_duplex"].get()):
            return "单面打印无翻面与取纸停顿"
        flow = str(self.vars["print_flow"].get() or "batch_all")
        if flow == "per_file":
            return "每个文件：正面 → 提醒翻面 → 反面 → 提醒取纸 → 下一个文件"
        if flow == "batch_take":
            return "正面全部一次打完 → 翻面一次 → 反面逐文件打印，每个文件后暂停取纸"
        return "正面全部一次打完 → 翻面一次 → 反面全部打完 → 完成后统一取纸"

    def _update_scope_hint(self) -> None:
        """逐文件范围的说明：独立一行、口语化，不使用旧的覆盖术语。

        未单独设置的文件直接显示它实际生效的范围文本（浅灰整行）；
        单独设置过的行用深色 + ＊ 前缀区分。
        未开启逐文件设置时，这一行用来常驻展示两种翻面模式的停顿区别。
        """
        per_file = bool(self.vars["per_file"].get())
        if per_file:
            self._drop_hint_text = "可多选后批量设置，也可双击单行修改"
        else:
            # 旧版固定写“可拖入…”，与当前进度无关；改为按阶段提示下一步动作。
            pending = [f for f in self.files
                       if f.get("kind") in ("word", "wps", "image")
                       and not f.get("pdf_path") and f.get("status") != "转换失败"]
            if not self.files:
                self._drop_hint_text = "第 1 步 · 拖入文件或点“＋ 添加文件”"
            elif pending:
                self._drop_hint_text = f"第 2 步 · 还有 {len(pending)} 个文件待转换"
            else:
                self._drop_hint_text = "第 3 步 · 可以开始打印"
        self._restore_drop_hint()
        if not hasattr(self, "scope_hint_label"):
            return
        if per_file:
            # 单行显示：1024 宽下列表卡约 423px，这句约 300px，不换行也不截断
            self.scope_hint_var.set(
                "未单独设置的文件按上方范围打印，带 ＊ 的是单独设置过的。")
        else:
            self.scope_hint_var.set(self._mode_hint_text())
        self.scope_hint_label.grid()

    def _set_widget_state(self, widget, state: str) -> None:
        """安全设置控件状态：高级设置未打开时控件可能尚未创建。"""
        if widget is None:
            return
        try:
            if widget.winfo_exists():
                widget.configure(state=state)
        except tk.TclError:
            pass

    def _update_start_state(self) -> None:
        """开始打印按钮的禁用态：无文件或仍有文件未转换时置灰。

        旧版无论列表是否为空，“开始打印”都保持可点，用户点下才弹提示；
        现在空列表/转换中直接禁用，主操作的可用性一目了然。
        """
        button = getattr(self, "start_btn", None)
        if button is None:
            return
        if getattr(self, "print_active", False):
            self._set_widget_state(button, "disabled")
            return
        pending = [f for f in self.files
                   if f.get("kind") in ("word", "wps", "image")
                   and not f.get("pdf_path") and f.get("status") != "转换失败"]
        self._set_widget_state(button, "normal" if self.files and not pending
                               else "disabled")

    def _sync_widget_states(self) -> None:
        mode = str(self.vars["range_mode"].get())
        custom = mode in ("custom", "recent")
        # 范围输入框与错误行始终占位，只切换可用状态：避免输入/清空时整条设置栏重排
        self._set_widget_state(self.range_entry, "normal" if custom else "disabled")
        if custom:
            self._update_range_error()
        else:
            self.range_error_var.set("")

        # “本次范围”列只在启用“逐文件单设范围/份数”时显示，平时把宽度让给
        # “输出与补页”和“状态”，避免窄窗口下列被截断
        if hasattr(self, "tree"):
            per_file = bool(self.vars["per_file"].get())
            columns = (("name", "pages", "scope", "plan", "status") if per_file
                       else ("name", "pages", "plan", "status"))
            if getattr(self, "_display_columns", None) != columns:
                self._display_columns = columns
                self.tree["displaycolumns"] = columns
                if per_file:
                    self.tree.column("scope", width=self._px(132), minwidth=self._px(110))

        queue_mode = str(self.vars["wait_mode"].get()) == "queue"
        self._set_widget_state(getattr(self, "delay_spin", None),
                               "disabled" if queue_mode else "normal")

        custom_scale = str(self.vars["scale_mode"].get()) == "custom"
        self._set_widget_state(self.scale_percent_spin, "normal" if custom_scale else "disabled")

        custom_paper = str(self.vars["paper_name"].get()) == CUSTOM_PAPER_NAME
        for entry in (getattr(self, "custom_paper_w_entry", None),
                      getattr(self, "custom_paper_h_entry", None)):
            self._set_widget_state(entry, "normal" if custom_paper else "disabled")

        booklet = bool(self.vars["booklet"].get())
        self._set_widget_state(getattr(self, "nup_combo", None),
                               "disabled" if booklet else "readonly")
        self._set_widget_state(getattr(self, "order_combo", None),
                               "disabled" if booklet else "readonly")

        # 打印流程三选一：单面时三项都没有意义，整体置灰。
        # 份数>1 且未按份执行时，反面会跨文件混在一起，无法按文件抽纸：
        # 流程 2 置灰，若当前已选中则自动切回流程 1 并在日志说明。
        duplex = bool(self.vars["manual_duplex"].get())
        multi_copies = self._int_setting("copies", 1) > 1
        flow = str(self.vars["print_flow"].get() or "batch_all")
        if flow not in ("batch_all", "batch_take", "per_file"):
            flow = "batch_all"
            self.vars["print_flow"].set(flow)
        flow2_usable = duplex and not (multi_copies and not bool(
            self.vars["copy_exec_on"].get()))
        for radio in (getattr(self, "flow_batch_all_radio", None),
                      getattr(self, "flow_batch_take_radio", None),
                      getattr(self, "flow_per_file_radio", None)):
            self._set_widget_state(radio, "normal" if duplex else "disabled")
        self._set_widget_state(getattr(self, "flow_batch_take_radio", None),
                               "normal" if flow2_usable else "disabled")
        if duplex and multi_copies and not bool(self.vars["copy_exec_on"].get()) \
                and flow == "batch_take":
            self.vars["print_flow"].set("batch_all")
            flow = "batch_all"
            if not getattr(self, "_flow2_warned", False):
                self._append_log("份数>1 且未按份执行：流程 2 无法按文件抽纸，已切换为流程 1。")
                self._flow2_warned = True
        elif flow2_usable:
            self._flow2_warned = False

        batch_flow = flow in ("batch_all", "batch_take")
        self._set_widget_state(
            self.copy_exec_check,
            "normal" if duplex and batch_flow and multi_copies else "disabled")
        if not (duplex and batch_flow and multi_copies) and bool(
                self.vars["copy_exec_on"].get()):
            # 仅在确实需要复位时写变量，避免 set 触发 trace 造成递归
            self.vars["copy_exec_on"].set(False)
        if hasattr(self, "flow_note_var"):
            self.flow_note_var.set(
                "" if (flow2_usable or not duplex)
                else "份数>1 需按份执行才可按文件取纸")
        self._update_scope_hint()

        # 中文下拉与内部枚举值同步（恢复默认 / 读配置后也保持一致）
        scale_value = str(self.vars["scale_mode"].get() or "shrink")
        if hasattr(self, "scale_label_var"):
            label = SCALE_LABELS.get(scale_value, SCALE_LABELS["shrink"])
            if str(self.scale_label_var.get()) != label:
                self.scale_label_var.set(label)
        if hasattr(self, "range_label_var"):
            range_value = str(self.vars["range_mode"].get() or "all")
            label = RANGE_LABELS.get(range_value, RANGE_LABELS["all"])
            if str(self.range_label_var.get()) != label:
                self.range_label_var.set(label)

        # 翻页按钮可用状态与页码显示随预览页数即时更新
        if hasattr(self, "preview_prev_btn"):
            has_pages = self.preview_pages > 0
            self._set_widget_state(self.preview_prev_btn,
                                   "normal" if has_pages and self.preview_page_index > 0
                                   else "disabled")
            self._set_widget_state(self.preview_next_btn,
                                   "normal" if has_pages and
                                   self.preview_page_index < self.preview_pages - 1
                                   else "disabled")
            self._set_widget_state(getattr(self, "preview_page_spin", None),
                                   "normal" if has_pages else "disabled")
            try:
                self.preview_page_spin.configure(
                    from_=1, to=max(1, self.preview_pages))
            except tk.TclError:
                pass

    def _effective_range_text(self, mode: Optional[str] = None) -> str:
        """当前生效的范围文本；最近使用范围取上次成功打印的范围。"""
        mode = mode or str(self.vars["range_mode"].get())
        if mode == "recent":
            last = str(self.settings.get("last_range") or "").strip()
            return last or str(self.vars["range_text"].get())
        return str(self.vars["range_text"].get())

    def _update_range_error(self) -> None:
        mode = str(self.vars["range_mode"].get())
        text = self._effective_range_text(mode)
        try:
            parse_page_range(text)
        except PdfError as exc:
            self.range_error_var.set(f"范围有误：{str(exc).splitlines()[0]}")
            self.range_error_label.configure(foreground=self.colors["danger"])
            # 错误态只换输入框描边（1px→1px），错误行始终占位，不改变布局高度
            self._set_entry_error(True)
            return
        self.range_error_var.set("")
        self._set_entry_error(False)

    def _set_entry_error(self, has_error: bool) -> None:
        entry = getattr(self, "range_entry", None)
        if entry is None:
            return
        try:
            entry.configure(style="Error.TEntry" if has_error else "TEntry")
        except tk.TclError:
            pass

    def _float_setting(self, key: str, default: float) -> float:
        try:
            return float(str(self.vars[key].get()).strip())
        except (TypeError, ValueError):
            return float(default)

    def _int_setting(self, key: str, default: int) -> int:
        try:
            return int(float(str(self.vars[key].get()).strip()))
        except (TypeError, ValueError):
            return int(default)

    def _get_printable_insets(self):
        """取当前打印机的可打印区域内缩（点，左/上/右/下），按打印机缓存。

        驱动没返回时返回 None：排版只用用户页边距，预览虚线贴着纸张边缘并提示。
        """
        printer = self.current_printer()
        if not printer:
            return None
        if self.printable_insets_printer == printer:
            return self.printable_insets
        insets = get_printable_insets_pt(printer)
        self.printable_insets_printer = printer
        self.printable_insets = insets
        return insets

    def _layout_options(self) -> LayoutOptions:
        """把界面变量转换为 LayoutOptions（非法数值回退默认值）。

        可打印区域内缩来自驱动（GetDeviceCaps），与用户页边距取并集后再排版，
        使“收缩到可打印区”的留白比例与 Acrobat 一致。
        """
        nup = self._int_setting("nup", 1)
        if nup not in NUP_CHOICES:
            nup = 1
        booklet = bool(self.vars["booklet"].get())
        source_code = 7
        source_name = str(self.vars["source_name"].get())
        for name, code in self.paper_sources:
            if name == source_name:
                source_code = code
                break
        return LayoutOptions(
            paper_name=str(self.vars["paper_name"].get() or "A4"),
            paper_width_mm=self._float_setting("paper_width_mm", 210.0),
            paper_height_mm=self._float_setting("paper_height_mm", 297.0),
            orientation=str(self.vars["orientation"].get() or "auto"),
            source_name=source_name,
            source_code=source_code,
            auto_rotate=bool(self.vars["auto_rotate"].get()),
            scale_mode=str(self.vars["scale_mode"].get() or "shrink"),
            scale_percent=self._int_setting("scale_percent", 100),
            nup=2 if booklet else nup,
            nup_border=bool(self.vars["nup_border"].get()),
            page_order=str(self.vars["page_order"].get() or "row"),
            booklet=booklet,
            margin_left_mm=self._float_setting("margin_left_mm", 10.0),
            margin_top_mm=self._float_setting("margin_top_mm", 10.0),
            margin_right_mm=self._float_setting("margin_right_mm", 10.0),
            margin_bottom_mm=self._float_setting("margin_bottom_mm", 8.0),
            crop_left_mm=self._float_setting("crop_left_mm", 0.0),
            crop_top_mm=self._float_setting("crop_top_mm", 0.0),
            crop_right_mm=self._float_setting("crop_right_mm", 0.0),
            crop_bottom_mm=self._float_setting("crop_bottom_mm", 0.0),
            manual_duplex=bool(self.vars["manual_duplex"].get()),
            annots=bool(self.vars["annots"].get()),
            skip_blank=bool(self.vars["skip_blank"].get()),
            printable_insets_pt=self._get_printable_insets(),
        )

    def _printer_options(self, landscape: bool) -> PrinterOptions:
        layout = self._layout_options()
        return PrinterOptions(
            landscape=landscape,
            paper_name=layout.paper_name,
            paper_width_mm=layout.paper_width_mm,
            paper_height_mm=layout.paper_height_mm,
            source_code=layout.source_code,
            color_mode=str(self.vars["color_mode"].get() or "color"),
            print_quality=str(self.vars["print_quality"].get() or "normal"),
            dpi=self._int_setting("dpi", 300),
            ink_saver=bool(self.vars["ink_saver"].get()),
            annots=bool(self.vars["annots"].get()),
        )

    def _collect_settings(self) -> dict:
        data = default_settings()
        for key, var in self.vars.items():
            if isinstance(var, tk.BooleanVar):
                data[key] = bool(var.get())
                continue
            text = str(var.get()).strip()
            if key == "printer":
                data[key] = self.current_printer()
            elif key in ("copies", "dpi", "scale_percent", "nup"):
                try:
                    data[key] = int(float(text))
                except ValueError:
                    pass
            elif key in ("paper_width_mm", "paper_height_mm", "margin_left_mm",
                         "margin_top_mm", "margin_right_mm", "margin_bottom_mm",
                         "crop_left_mm", "crop_top_mm", "crop_right_mm",
                         "crop_bottom_mm", "delay_seconds"):
                try:
                    data[key] = float(text)
                except ValueError:
                    pass
            else:
                data[key] = text
        data["recent_files"] = list(self.recent_files)[:20]
        data["recent_ranges"] = list(self.recent_ranges)[:20]
        data["ui_style"] = "A"
        data["preview_view"] = str(self.preview_view_var.get() or "single")
        data["window_geometry"] = str(self.settings.get("window_geometry") or "")
        return data

    def _save_settings_now(self) -> None:
        try:
            save_settings(self._collect_settings())
        except SettingsError as exc:
            self._append_log(str(exc))

    def restore_defaults(self) -> None:
        """恢复默认打印设置（保留文件列表与最近记录）。"""
        if not messagebox.askokcancel("恢复默认", "确定恢复所有打印设置为默认值吗？",
                                       parent=self.root):
            return
        defaults = default_settings()
        for key, var in self.vars.items():
            if key == "printer":
                continue
            if key in defaults:
                var.set(defaults[key])
        self.duplex_mode_var.set("single" if not defaults.get("manual_duplex") else "duplex")
        self.preview_view_var.set(str(defaults.get("preview_view") or "single"))
        self._sync_widgets()
        self._update_advanced_summary()
        self._save_settings_now()
        self._append_log("已恢复默认打印设置。")

    def restore_acrobat_defaults(self) -> None:
        """一键回到“Acrobat 直接打印”的基线（不弹窗，立即生效）。"""
        data = acrobat_defaults()
        for key, var in self.vars.items():
            if key == "printer":
                continue
            if key in data:
                var.set(data[key])
        self.duplex_mode_var.set("single")
        self.preview_view_var.set("single")
        self._sync_widget_states()
        self._update_advanced_summary()
        self._update_preview()
        self._save_settings_now()
        self.status_var.set("已恢复 Acrobat 等效默认：全部页面 · 1 份 · 收缩到可打印区 · 单面")
        self._append_log("已恢复 Acrobat 等效默认参数（单面输出，缩放=收缩到可打印区）。")

    # ------------------------------------------------------- 高级设置对话框
    def open_advanced_settings(self, visible: bool = True) -> None:
        """打开高级设置（非模态，改动实时回传主界面与预览）。

        visible=False 供离屏截图/自动化使用：窗口构建完成后仍保持隐藏，
        由调用方自行设置透明度后再 deiconify，避免上屏。
        """
        if self.advanced_window is not None and self.advanced_window.winfo_exists():
            self.advanced_window.deiconify()
            self.advanced_window.lift()
            return
        win = tk.Toplevel(self.root)
        # 首帧不闪黑：Toplevel 创建即映射，此时内容还没构建；先 withdraw，
        # 设成与主题一致的底色，构建完内容与滚动区域后再 deiconify。
        win.withdraw()
        win.title("高级设置 · PDF 双面打印")
        win.transient(self.root)
        win.configure(background=self.colors["bg"])
        win.minsize(1040, 720)
        win.protocol("WM_DELETE_WINDOW", self._close_advanced_settings)

        footer = ttk.Frame(win, padding=(14, 10))
        footer.pack(side="bottom", fill="x")
        ttk.Label(footer, text="改动立即生效，无需“应用”。", style="Muted.TLabel").pack(side="left")
        self.restore_acrobat_btn = ttk.Button(
            footer, text="恢复 Acrobat 等效默认（单面）", style="Ghost.TButton",
            command=self.restore_acrobat_defaults)
        self.restore_acrobat_btn.pack(side="right")
        ttk.Button(footer, text="关闭", command=self._close_advanced_settings).pack(
            side="right", padx=(0, 8))

        # 内容区可滚动：小屏/缩放比例较大时也不会裁掉底部选项
        canvas = tk.Canvas(win, background=self.colors["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        body = ttk.Frame(canvas, padding=14, style="Page.TFrame")
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")

        # 内容高度变化时刷新 scrollregion（不能只在初始化时设一次）；
        # Canvas 尺寸变化时同时刷新宽度与 scrollregion，保证可滚动区域跟随窗口。
        def _refresh_adv_scrollregion(_event=None) -> None:
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except tk.TclError:
                pass

        body.bind("<Configure>", _refresh_adv_scrollregion)
        canvas.bind("<Configure>",
                    lambda e: (canvas.itemconfigure(body_id, width=max(600, e.width)),
                               _refresh_adv_scrollregion()))
        self._build_advanced_content(body)
        self._sync_widget_states()
        self._update_advanced_summary()
        # 滚轮：对话框构建完成即绑定（bind_all + 归属校验），关闭时解绑
        self._bind_advanced_wheel(win, canvas)
        try:
            win.bind("<Destroy>",
                     lambda e: self._unbind_advanced_wheel()
                     if (getattr(e, "widget", None) is win
                         and getattr(self, "_adv_window", None) is win) else None)
        except tk.TclError:
            pass

        # 内容全部就绪后再决定位置并显示（只定位一次，避免可见后重排）
        win.update_idletasks()
        self._place_toplevel_over_root(win, 1200, 1000)
        self.advanced_window = win
        if visible:
            win.deiconify()
            win.lift()

    def _place_toplevel_over_root(self, win, width: int, height: int) -> None:
        """在显示前一次性算好 Toplevel 的尺寸与位置（相对主窗居中且不越界）。

        必须在 deiconify 之前调用：可见后再改 geometry 会让用户看到一次重排。
        """
        try:
            self.root.update_idletasks()
            root_w = max(1, self.root.winfo_width())
            root_h = max(1, self.root.winfo_height())
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
        except tk.TclError:
            return
        try:
            areas = list_work_areas()
        except Exception:
            areas = []
        area = areas[0] if areas else (0, 0, self.root.winfo_screenwidth(),
                                       self.root.winfo_screenheight())
        area_w = max(400, area[2] - area[0])
        area_h = max(300, area[3] - area[1])
        width = max(600, min(int(width), area_w))
        height = max(400, min(int(height), area_h))
        x = root_x + max(0, (root_w - width) // 2)
        y = root_y + max(0, (root_h - height) // 2)
        x = max(area[0], min(x, area[2] - width))
        y = max(area[1], min(y, area[3] - height))
        win.geometry(f"{width}x{height}+{x}+{y}")

    def _close_advanced_settings(self) -> None:
        # 先解绑滚轮再销毁：否则 bind_all 会残留在主窗口上，导致主界面滚轮被吞
        self._unbind_advanced_wheel()
        if self.advanced_window is not None:
            try:
                self.advanced_window.destroy()
            except tk.TclError:
                pass
            self.advanced_window = None

    # --------------------------------------------------- 高级设置滚轮滚动
    def _bind_advanced_wheel(self, win, canvas) -> None:
        """把滚轮统一绑到 Canvas 及其所有子控件上。

        Tk 的 bind_all 会在指针所在控件上触发，因此无需逐个子控件 bind；
        对话框关闭时必须 unbind_all，否则主窗口的滚轮也会被这个处理器吃掉。
        """
        self._adv_canvas = canvas
        self._adv_window = win
        try:
            self.root.bind_all("<MouseWheel>", self._on_advanced_wheel, add="+")
            self.root.bind_all("<Button-4>", self._on_advanced_wheel, add="+")
            self.root.bind_all("<Button-5>", self._on_advanced_wheel, add="+")
        except tk.TclError:
            self._adv_canvas = None

    def _unbind_advanced_wheel(self) -> None:
        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            try:
                self.root.unbind_all(sequence)
            except tk.TclError:
                pass
        self._adv_canvas = None
        self._adv_window = None

    def _adv_widget_inside_dialog(self, widget) -> bool:
        """事件发生在高级设置对话框内部（Canvas 或它的任意子控件）。"""
        win = getattr(self, "_adv_window", None)
        if win is None:
            return False
        while widget is not None:
            if widget is win:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _adv_content_overflows(self) -> bool:
        """内容是否真的超出视口：未超出时滚轮不滚动。"""
        canvas = getattr(self, "_adv_canvas", None)
        if canvas is None:
            return False
        try:
            box = canvas.bbox("all")
            if not box:
                return False
            content_h = int(box[3]) - int(box[1])
            view_h = max(1, canvas.winfo_height())
        except tk.TclError:
            return False
        return content_h > view_h

    def _on_advanced_wheel(self, event) -> Optional[str]:
        """滚轮处理器：按 delta 滚动 Canvas；到顶/到底不越界。

        - Spinbox 上不滚动页面（保留 Spinbox 自身的滚轮改值行为）；
        - Button-4/5 兼容 Linux/X11；Windows 用 delta（每格 120）。
        """
        canvas = getattr(self, "_adv_canvas", None)
        widget = getattr(event, "widget", None)
        if canvas is None or not self._adv_widget_inside_dialog(widget):
            return None
        # 指针在 Spinbox 上时让 Spinbox 自己处理（不滚页面）
        try:
            if isinstance(widget, (tk.Spinbox, ttk.Spinbox)):
                return None
        except AttributeError:
            pass
        if not self._adv_content_overflows():
            return "break"
        delta = int(getattr(event, "delta", 0) or 0)
        if getattr(event, "num", 0) == 4:
            delta = 120
        elif getattr(event, "num", 0) == 5:
            delta = -120
        if delta == 0:
            return "break"
        # 每格滚 3 行：120 → 3 行，240 → 6 行，高精度滚轮至少 1 行
        lines = max(1, int(abs(delta) * 3 / 120))
        steps = -lines if delta > 0 else lines
        try:
            first, last = canvas.yview()
        except tk.TclError:
            return "break"
        if steps < 0 and first <= 0.0:      # 已在顶部，继续向上不越界
            return "break"
        if steps > 0 and last >= 1.0:       # 已在底部，继续向下不越界
            return "break"
        try:
            canvas.yview_scroll(steps, "units")
        except tk.TclError:
            pass
        return "break"

    def _update_advanced_summary(self) -> None:
        """主界面显示一行“当前高级设置摘要”，只列非默认项。"""
        if not hasattr(self, "advanced_summary_var"):
            return
        parts: List[str] = []
        defaults = acrobat_defaults()
        nup = self._int_setting("nup", 1)
        if bool(self.vars["booklet"].get()):
            parts.append("小册子")
        elif nup > 1:
            parts.append(f"{nup} 合 1")
        color = str(self.vars["color_mode"].get())
        if color == "gray":
            parts.append("灰度")
        elif color == "mono":
            parts.append("黑白")
        if bool(self.vars["ink_saver"].get()):
            parts.append("省墨")
        if str(self.vars["print_quality"].get()) != "normal":
            parts.append({"draft": "草稿质量", "high": "高质量"}.get(
                str(self.vars["print_quality"].get()), str(self.vars["print_quality"].get())))
        paper = str(self.vars["paper_name"].get())
        if paper and paper != defaults.get("paper_name"):
            parts.append(paper)
        if str(self.vars["orientation"].get()) != "auto":
            parts.append("横向" if str(self.vars["orientation"].get()) == "landscape" else "纵向")
        if bool(self.vars["annots"].get()):
            parts.append("打印批注")
        if bool(self.vars["skip_blank"].get()):
            parts.append("忽略空白页")
        if bool(self.vars["reverse"].get()):
            parts.append("逆序")
        if str(self.vars["collate"].get()) == "uncollated":
            parts.append("逐页不分页")
        if bool(self.vars["per_file"].get()):
            parts.append("逐文件范围")
        sides = ("left", "top", "right", "bottom")
        margins = tuple(self._float_setting(f"margin_{side}_mm", 0.0) for side in sides)
        expected_margins = tuple(float(defaults.get(f"margin_{side}_mm", 0.0)) for side in sides)
        if margins != expected_margins:
            parts.append("自定义边距")
        crops = (self._float_setting("crop_left_mm", 0.0), self._float_setting("crop_top_mm", 0.0),
                 self._float_setting("crop_right_mm", 0.0), self._float_setting("crop_bottom_mm", 0.0))
        if any(value > 0 for value in crops):
            parts.append("裁切")
        self.advanced_summary_var.set(
            "高级设置：" + (" · ".join(parts) if parts else "全部为默认值"))

    def _apply_window_geometry(self) -> None:
        """启动时恢复窗口几何；越界/屏幕外/尺寸过小则回退为主屏居中。"""
        saved = str(self.settings.get("window_geometry") or "")
        try:
            areas = list_work_areas()
        except Exception:
            areas = []
        width, height, x, y = choose_geometry(
            saved, areas, WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE)
        self.root.geometry(format_geometry(width, height, x, y))

    def _remember_window_geometry(self) -> None:
        """关闭时记忆窗口位置与大小（仅 normal 状态，最大化不写入）。"""
        try:
            if self.root.state() != "normal":
                return
            geometry = self.root.geometry()
            width, height, x, y = choose_geometry(
                geometry, list_work_areas(), WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE)
            self.settings["window_geometry"] = format_geometry(width, height, x, y)
        except Exception:
            pass

    def current_printer(self) -> str:
        """当前选择的打印机（去掉“（默认）”标记）。"""
        text = str(self.vars["printer"].get())
        return text.replace("（默认）", "").strip()

    # ------------------------------------------------------------------ 打印机
    def _refresh_printers(self) -> None:
        try:
            printers = list_printers()
        except Exception as exc:
            printers = []
            self._append_log(f"枚举打印机失败：{exc}")
        default = get_default_printer()
        values = [p + ("（默认）" if p == default else "") for p in printers]
        self.printer_combo["values"] = values
        current = self.current_printer()
        if not printers:
            self.status_var.set("未检测到打印机，请先在 Windows 中安装打印机。")
            self._set_printer_status("未检测到打印机", self.colors["danger"])
            self.paper_sources = [("自动选择", 7)]
            if hasattr(self, "source_combo"):
                self.source_combo["values"] = ["自动选择"]
            self.vars["source_name"].set("自动选择")
            return
        if current not in printers:
            chosen = default if default in printers else printers[0]
            self.vars["printer"].set(chosen + ("（默认）" if chosen == default else ""))
        # 纸张来源随打印机更新
        try:
            self.paper_sources = get_printer_sources(self.current_printer())
        except Exception:
            self.paper_sources = [("自动选择", 7)]
        names = [name for name, _code in self.paper_sources]
        if hasattr(self, "source_combo"):
            self.source_combo["values"] = names
        if str(self.vars["source_name"].get()) not in names:
            self.vars["source_name"].set(names[0] if names else "自动选择")
        self._poll_printer_status(reschedule=False)

    def _poll_printer_status(self, reschedule: bool = True) -> None:
        # 退出中：直接返回，不查询驱动、不再重排 after
        if self.shutting_down:
            return
        printer = self.current_printer()
        if not printer:
            self._set_printer_status("未选择打印机", self.colors["danger"])
            return
        try:
            status = get_printer_status(printer)
        except Exception as exc:
            self._set_printer_status(f"状态未知：{exc}", self.colors["danger"])
            return
        text = str(status["text"])
        if status["offline"] or status["jam"] or status["paper_out"]:
            color = self.colors["danger"]
        elif status["printing"] or int(status["jobs"] or 0) > 0:
            color = self.colors["warning"]
        else:
            color = self.colors["success"]
        tag = "（默认打印机）" if status.get("default") else ""
        self._set_printer_status(f"状态：{text}{tag}", color)
        if reschedule and not self.shutting_down and self.root.winfo_exists() and not (
                self.worker and self.worker.is_alive()):
            # 先取消旧 id，避免手动刷新与轮询重排产生未跟踪的重复 after
            if self.printer_poll_id is not None:
                try:
                    self.root.after_cancel(self.printer_poll_id)
                except tk.TclError:
                    pass
            self.printer_poll_id = self.root.after(5000, self._poll_printer_status)

    def _set_printer_status(self, text: str, color: str = "") -> None:
        self.printer_status_var.set(text)
        try:
            self.printer_status_label.configure(
                foreground=color or self.colors["accent"])
        except tk.TclError:
            pass

    def _browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="选择排版 PDF 输出目录", parent=self.root)
        if folder:
            self.vars["output_dir"].set(folder)

    # ------------------------------------------------------------------ 文件列表
    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 PDF、Word、图片或文本文件",
            filetypes=[("支持的文件", "*.pdf *.doc *.docx *.rtf *.txt *.odt *.wps "
                                  "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"),
                       ("PDF 文件", "*.pdf"),
                       ("Word / WPS / 文本", "*.doc *.docx *.rtf *.txt *.odt *.wps"),
                       ("图片", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"),
                       ("所有文件", "*.*")],
            parent=self.root,
        )
        if paths:
            self._add_paths(paths)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含 PDF/Word 的文件夹", parent=self.root)
        if not folder:
            return
        try:
            names = sorted(os.listdir(folder))
        except OSError as exc:
            messagebox.showerror("无法读取文件夹", str(exc), parent=self.root)
            return
        paths = [os.path.join(folder, name) for name in names
                 if (name.lower().endswith(".pdf") or is_word_path(name)
                     or is_wps_path(name) or is_image_path(name))
                 and os.path.isfile(os.path.join(folder, name))]
        if not paths:
            messagebox.showinfo("提示", "该文件夹内没有 PDF、Word、WPS 或图片文件。",
                                parent=self.root)
            return
        self._add_paths(paths)

    def _add_paths(self, paths) -> None:
        existing = {os.path.normcase(os.path.abspath(f["path"])) for f in self.files}
        failures = []
        ignored = []
        added = 0
        for path in paths:
            text = str(path).strip()
            if not text:
                continue
            word = is_word_path(text)
            wps = is_wps_path(text)
            image = is_image_path(text)
            if not (is_pdf_path(text) or word or wps or image):
                ignored.append(os.path.basename(text))
                continue
            abspath = os.path.abspath(text)
            key = os.path.normcase(abspath)
            if key in existing:
                continue
            if word or wps or image:
                # Word/WPS/图片：页数要等转换后才知道，先入列表并标记“待转换”
                record = {"path": abspath, "pdf_path": "", "pages": None,
                          "status": "待转换",
                          "kind": "image" if image else ("wps" if wps else "word"),
                          "range_text": "", "copies": None}
                self.recent_files = add_recent(self.recent_files, abspath)
            else:
                try:
                    pages = self._pdf_pages_cached(abspath)
                    status = "待打印"
                except PdfError as exc:
                    pages = None
                    status = "无法读取"
                    failures.append((os.path.basename(abspath), str(exc)))
                record = {"path": abspath, "pdf_path": abspath, "pages": pages,
                          "status": status, "kind": "pdf",
                          "range_text": "", "copies": None}
                self.recent_files = add_recent(self.recent_files, abspath)
            iid = self.tree.insert(
                "", "end",
                values=(os.path.basename(abspath), record["pages"] or "?", "", "",
                        self._status_display(record["status"])),
            )
            record["iid"] = iid
            self.files.append(record)
            existing.add(key)
            added += 1

        if added:
            self.status_var.set(f"已添加 {added} 个文件，当前共 {len(self.files)} 个待打印。")
            self._append_log(f"已添加 {added} 个文件，当前共 {len(self.files)} 个待打印。")
        elif not failures and not ignored:
            self.status_var.set("没有新增文件（所选文件已在列表中或不是受支持的文件）。")
        if ignored:
            names = "、".join(ignored[:5]) + (" …" if len(ignored) > 5 else "")
            self._append_log(
                f"已忽略 {len(ignored)} 个不支持的文件（支持 PDF、Word/WPS/文本、图片）：{names}")
        self._populate_recent_menu()
        self._start_word_conversions()
        self._update_preview()
        self._save_settings_now()

        if failures:
            detail = "\n".join(f"• {name}：{msg.splitlines()[0]}" for name, msg in failures[:5])
            if len(failures) > 5:
                detail += f"\n…另有 {len(failures) - 5} 个文件无法读取。"
            messagebox.showwarning(
                "部分文件无法读取",
                "以下文件已加入列表，但打印时会被跳过：\n\n" + detail,
                parent=self.root,
            )

    # ------------------------------------------------------ Word → 临时 PDF
    def _record_pdf_path(self, record: dict) -> str:
        """返回该记录可用于打印/预览的 PDF 路径（Word 转换完成后才有）。"""
        return str(record.get("pdf_path") or "")

    def _start_word_conversions(self) -> None:
        """有未转换的 Word/WPS/图片文件时，启动串行后台转换线程。"""
        pending = [f for f in self.files
                   if f.get("kind") in ("word", "wps", "image")
                   and not f.get("pdf_path")
                   and f.get("status") != "转换失败"]
        if not pending:
            return
        if self.word_thread is not None and self.word_thread.is_alive():
            return
        self.word_thread = threading.Thread(target=self._word_worker, daemon=True,
                                            name="file-convert")
        self.word_thread.start()

    def _word_worker(self) -> None:
        """串行把 Word/WPS/图片转成临时 PDF；COM 在子进程里执行，界面不冻结。"""
        jobs = [f for f in list(self.files)
                if f.get("kind") in ("word", "wps", "image") and not f.get("pdf_path")
                and f.get("status") != "转换失败"]
        if not jobs:
            return
        self.ui_queue.put(
            ("status", f"正在转换 {len(jobs)} 个文件，转换完成前无法打印…"))
        converters = available_converters()
        has_office_job = any(f.get("kind") in ("word", "wps") for f in jobs)
        if not converters and has_office_job:
            # 图片不需要 Office，仍继续处理图片；Office 文档统一给中文提示
            for record in jobs:
                if record.get("kind") == "image":
                    continue
                message = ("未检测到 WPS，无法转换 .wps 文件。请安装 WPS，"
                           "或用 WPS 打开后另存为 PDF/Word 再导入。"
                           if record.get("kind") == "wps" else
                           "未检测到 Word 或 WPS，无法把文档转成 PDF。"
                           "请先安装 Microsoft Word / WPS，或手动另存为 PDF 后重新导入。")
                self.ui_queue.put(("word_failed", record.get("iid"), message))
        if converters:
            self.ui_queue.put(("log",
                               f"检测到转换组件：{'、'.join(converters)}，开始转换文档…"))
        image_jobs = [f for f in jobs if f.get("kind") == "image"]
        office_jobs = [f for f in jobs if f.get("kind") in ("word", "wps")]
        # 多份 Office 文档走“一个子进程 + 常驻 Word 实例”批量转换：
        # 同批 4 个文件实测总耗时由约 29s 降到约 11s（首个文件仍含冷启动）。
        batch_office = bool(converters) and len(office_jobs) > 1
        if image_jobs:
            self.ui_queue.put(("log", f"开始转换 {len(image_jobs)} 个图片文件…"))
        for record in jobs:
            if self.shutting_down or self.stop_event.is_set():
                return
            kind = record.get("kind")
            if kind in ("word", "wps") and not converters:
                continue
            if kind in ("word", "wps") and batch_office:
                continue                      # 交由批量转换统一处理
            source = record["path"]
            if not os.path.isfile(source):
                self.ui_queue.put(("word_failed", record.get("iid"), "文件不存在，无法转换。"))
                continue
            temp_dir = create_temp_dir()
            dst = os.path.join(temp_dir,
                               os.path.splitext(os.path.basename(source))[0] + ".pdf")
            self.ui_queue.put(("rows", [(record.get("iid"), "转换中…")]))
            try:
                if kind == "image":
                    info = image_to_pdf(source, dst)
                else:
                    info = convert_file(source, dst, timeout=180)
            except (WordConvertError, ImageConvertError) as exc:
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.ui_queue.put(("word_failed", record.get("iid"), str(exc)))
                continue
            except Exception as exc:  # 转换异常不能拖垮程序
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.ui_queue.put(("word_failed", record.get("iid"),
                                   f"文件转换失败：{exc}"))
                continue
            if self.shutting_down:
                shutil.rmtree(temp_dir, ignore_errors=True)
                return
            self.word_temp_dirs[source] = temp_dir
            self.ui_queue.put(("word_done", record.get("iid"), dst,
                               info.get("pages") or 0,
                               info.get("converter") or ("图片" if kind == "image" else "Word"),
                               {"missing_fonts": info.get("missing_fonts"),
                                "unembedded_fonts": info.get("unembedded_fonts")}))
        if batch_office:
            self._convert_office_batch(office_jobs)
        if not (self.shutting_down or self.stop_event.is_set()):
            self.ui_queue.put(
                ("status", "文件转换结束，请确认列表状态后再开始打印。"))

    def _convert_office_batch(self, records: List[dict]) -> None:
        """一个常驻 Word 实例批量转换多个文档；逐文件回调更新界面。

        任一文件失败只影响该行（其余文件继续）；超时/中止由 convert_batch
        统一结束子进程并清理临时目录，不会残留 WINWORD 进程。
        """
        items: List[dict] = []
        staged: dict = {}
        for record in records:
            if self.shutting_down or self.stop_event.is_set():
                self._cleanup_unsaved_batch(staged)
                return
            source = record["path"]
            if not os.path.isfile(source):
                self.ui_queue.put(("word_failed", record.get("iid"), "文件不存在，无法转换。"))
                continue
            temp_dir = create_temp_dir()
            dst = os.path.join(temp_dir,
                               os.path.splitext(os.path.basename(source))[0] + ".pdf")
            staged[record.get("iid")] = (record, source, temp_dir, dst)
            self.ui_queue.put(("rows", [(record.get("iid"), "转换中…")]))
            items.append({"iid": record.get("iid"), "src": source, "dst": dst,
                          "kind": record.get("kind")})

        def handle(payload: dict) -> None:
            entry = staged.get(payload.get("iid"))
            if not entry:
                return
            record, source, temp_dir, dst = entry
            if payload.get("ok"):
                self.word_temp_dirs[source] = temp_dir
                self.ui_queue.put(("word_done", record.get("iid"), dst,
                                   payload.get("pages") or 0,
                                   payload.get("converter") or "Word",
                                   {"missing_fonts": payload.get("missing_fonts"),
                                    "unembedded_fonts": payload.get("unembedded_fonts")}))
            else:
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.ui_queue.put(("word_failed", record.get("iid"),
                                   payload.get("error") or "转换失败。"))

        try:
            convert_batch(items, timeout_per_file=180, on_result=handle,
                          stop_event=self.stop_event)
        except BatchCancelled:
            self._cleanup_unsaved_batch(staged)
        except Exception as exc:  # 批量链路异常：仅把未成功的文件标记失败
            for _record, source, temp_dir, _dst in staged.values():
                if source not in self.word_temp_dirs:
                    shutil.rmtree(temp_dir, ignore_errors=True)
            for record in records:
                entry = staged.get(record.get("iid"))
                if entry and entry[1] not in self.word_temp_dirs:
                    self.ui_queue.put(("word_failed", record.get("iid"),
                                       f"批量转换失败：{exc}"))

    def _cleanup_unsaved_batch(self, staged: dict) -> None:
        """中止批量转换时删除尚未登记成功的临时目录（已成功的不动）。"""
        for _record, source, temp_dir, _dst in staged.values():
            if source not in self.word_temp_dirs:
                shutil.rmtree(temp_dir, ignore_errors=True)

    def _release_word_record(self, record: dict) -> None:
        """删除某个转换记录的临时 PDF 目录，并把状态复位为“待转换”。"""
        source = record.get("path", "")
        temp_dir = self.word_temp_dirs.pop(source, "")
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        if record.get("kind") in ("word", "wps", "image"):
            record["pdf_path"] = ""
            record["pages"] = None
            record["status"] = "待转换"

    def _cleanup_word_temps(self, records=None) -> None:
        """清理转换临时 PDF（打印完成后与退出时调用，异常路径也覆盖）。"""
        targets = records if records is not None else list(self.files)
        cleaned = 0
        for record in list(targets):
            if record.get("kind") in ("word", "wps", "image") and record.get("pdf_path"):
                self._release_word_record(record)
                self._update_row(record)
                cleaned += 1
        if cleaned:
            self._append_log(
                f"已清理 {cleaned} 个转换临时 PDF；下次打印这些文件时会重新转换。")

    def _pdf_pages_cached(self, path: str, force: bool = False) -> int:
        """读取 PDF 页数，按 (mtime_ns, size) 缓存。

        * 同一路径在文件未变化时不会重复打开 PDF（预览刷新/改设置都走缓存）；
        * force=True 或 mtime/大小变化时才重新解析；
        * 解析失败抛 PdfError，由调用方转成界面状态。
        """
        stat = os.stat(path)
        key = (stat.st_mtime_ns, stat.st_size)
        cached = self.pdf_info_cache.get(path)
        if not force and cached and (cached[0], cached[1]) == key:
            return cached[2]
        pages = get_pdf_page_count(path)
        self.pdf_info_cache[path] = (key[0], key[1], pages)
        return pages

    def _refresh_file_metadata(self, notify: bool = True) -> None:
        """检查列表中的 PDF 是否被外部修改；变化时重新读取页数并更新状态列。

        F5 或下一次交互（刷新预览）都会调用，满足“外部变更自动同步”的要求。
        """
        if not self.files:
            return
        changed: List[str] = []
        for record in self.files:
            # Word/WPS/图片：只监视源文件，转换后的临时 PDF 由程序自己管理
            if record.get("kind") in ("word", "wps", "image"):
                continue
            path = record["path"]
            if not os.path.isfile(path):
                if record.get("status") != "文件不存在":
                    record["status"] = "文件不存在"
                    self._update_row(record)
                continue
            stat = os.stat(path)
            cached = self.pdf_info_cache.get(path)
            if cached and (cached[0], cached[1]) == (stat.st_mtime_ns, stat.st_size):
                continue
            try:
                pages = self._pdf_pages_cached(path, force=True)
                record["pages"] = pages
                record["status"] = "待打印" if record.get("status") in ("无法读取", "文件不存在") \
                    else record.get("status", "待打印")
                changed.append(os.path.basename(path))
            except PdfError as exc:
                record["pages"] = None
                record["status"] = "无法读取"
                self._append_log(f"重新读取失败：{os.path.basename(path)}：{str(exc).splitlines()[0]}")
            self._update_row(record)
        if changed:
            message = f"文件已变更，已重新读取：{'、'.join(changed[:5])}"
            self._append_log(message)
            if notify:
                self.status_var.set(message)

    def _update_row(self, record: dict) -> None:
        """把 record 的页数/状态写回文件列表行。"""
        iid = record.get("iid")
        if not iid or not self.tree.exists(iid):
            return
        try:
            _valid, _dropped, _copies = self._resolve_record(record)
            plan_text = self._plan_text(record)
        except PdfError:
            plan_text = "范围有误"
        values = list(self.tree.item(iid, "values"))
        while len(values) < 5:
            values.append("")
        values[1] = record.get("pages") if record.get("pages") else "?"
        values[2] = self._scope_text(record)
        values[3] = plan_text
        values[4] = self._status_display(record.get("status") or "")
        self.tree.item(iid, values=values)
        self._apply_row_tags()

    def _populate_recent_menu(self) -> None:
        self.recent_combo["values"] = list(self.recent_files)
        if self.recent_files and not self.recent_var.get():
            self.recent_var.set("")

    def _add_recent_file(self) -> None:
        path = str(self.recent_var.get()).strip()
        if path and os.path.isfile(path):
            self._add_paths([path])
        else:
            self.status_var.set("请先在最近使用列表中选择一个仍存在的文件。")

    def _clear_recent(self) -> None:
        self.recent_files.clear()
        self._populate_recent_menu()
        self._save_settings_now()

    def remove_selected(self) -> None:
        selected = list(self.tree.selection())
        if not selected:
            self.status_var.set("请先在列表中选中要移除的文件。")
            return
        # 批量移除属于破坏性操作：2 个及以上先确认，单个移除保持一步完成。
        if len(selected) > 1 and not messagebox.askokcancel(
                "移除文件",
                f"确定从列表移除选中的 {len(selected)} 个文件吗？\n"
                "原文件不会被删除，只是不再参与本次打印。",
                parent=self.root):
            return
        for iid in selected:
            if self.tree.exists(iid):
                self.tree.delete(iid)
            removed = [f for f in self.files if f["iid"] == iid]
            for record in removed:
                self._release_word_record(record)
            self.files = [f for f in self.files if f["iid"] != iid]
        self.status_var.set(f"已移除 {len(selected)} 个文件，剩余 {len(self.files)} 个。")
        self._append_log(f"已移除 {len(selected)} 个文件，剩余 {len(self.files)} 个。")
        self._update_preview()

    def clear_files(self) -> None:
        if not self.files:
            return
        if not messagebox.askokcancel("清空列表", "确定要清空待打印文件列表吗？",
                                       parent=self.root):
            return
        for record in self.files:
            self._release_word_record(record)
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.status_var.set("待打印列表已清空。")
        self._append_log("待打印列表已清空。")
        self._update_preview()

    def _sync_files_order(self) -> None:
        by_iid = {f["iid"]: f for f in self.files}
        self.files = [by_iid[iid] for iid in self.tree.get_children() if iid in by_iid]

    def move_selected(self, direction: int) -> None:
        selected = list(self.tree.selection())
        if not selected:
            return
        selected_set = set(selected)
        children = list(self.tree.get_children())
        ordered = sorted(selected, key=lambda iid: children.index(iid))
        if direction < 0:
            for iid in ordered:
                index = self.tree.index(iid)
                if index > 0 and children[index - 1] not in selected_set:
                    self.tree.move(iid, "", index - 1)
                    children = list(self.tree.get_children())
        else:
            for iid in reversed(ordered):
                index = self.tree.index(iid)
                if index < len(children) - 1 and children[index + 1] not in selected_set:
                    self.tree.move(iid, "", index + 1)
                    children = list(self.tree.get_children())
        self._sync_files_order()
        self._update_preview()

    def dedupe_files(self) -> None:
        seen = set()
        removed = 0
        for record in list(self.files):
            key = os.path.normcase(os.path.abspath(record["path"]))
            if key in seen:
                if self.tree.exists(record["iid"]):
                    self.tree.delete(record["iid"])
                self._release_word_record(record)
                self.files.remove(record)
                removed += 1
            else:
                seen.add(key)
        self._append_log(f"去重完成：移除 {removed} 个重复文件。")
        self._update_preview()

    def show_in_explorer(self) -> None:
        record = self._selected_record() or (self.files[0] if self.files else None)
        if not record:
            return
        try:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(record["path"])])
        except Exception as exc:
            self._append_log(f"无法打开资源管理器：{exc}")

    def open_selected_preview(self) -> None:
        record = self._selected_record()
        if record is None:
            return
        self._update_preview()

    def _find_record(self, iid) -> Optional[dict]:
        for record in self.files:
            if record.get("iid") == iid:
                return record
        return None

    def _selected_records(self) -> List[dict]:
        """当前选中的文件记录（保持列表顺序）。"""
        selected = set(self.tree.selection())
        return [record for record in self.files if record["iid"] in selected]

    def _selected_record(self) -> Optional[dict]:
        selected = set(self.tree.selection())
        for record in self.files:
            if record["iid"] in selected:
                return record
        return None

    def edit_selected_scope(self, focus: str = "both") -> None:
        """批量/单行设置范围与份数（紧凑对话框，带常用预设）。

        focus: range=定位到范围输入；copies=定位到份数；both=默认。
        """
        selected = self._selected_records()
        if not selected:
            self.status_var.set("请先选中要设置的文件。")
            return
        dialog = tk.Toplevel(self.root)
        # 同高级设置：先隐藏构建，避免可见后补内容闪一下系统底色
        dialog.withdraw()
        dialog.title("设置范围与份数")
        dialog.transient(self.root)
        dialog.configure(background=self.colors["bg"])
        dialog.columnconfigure(0, weight=1)

        max_pages = max([int(record.get("pages") or 0) for record in selected] + [1])
        odd_text = ",".join(str(page) for page in range(1, max_pages + 1, 2))
        even_text = ",".join(str(page) for page in range(2, max_pages + 1, 2))
        last_text = str(self.recent_ranges[0]) if self.recent_ranges else ""
        preset_row = ttk.Frame(dialog)
        preset_row.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))
        ttk.Label(preset_row, text="常用：", style="Muted.TLabel").pack(side="left", padx=(0, 4))

        range_var = tk.StringVar(value=str(selected[0].get("range_text") or ""))
        copies_var = tk.StringVar(
            value="" if selected[0].get("copies") is None else str(selected[0]["copies"]))

        def pick(value: str) -> None:
            range_var.set(value)
            preview_text()

        for label, value in (("全部", f"1-{max_pages}"), ("仅奇数页", odd_text),
                             ("仅偶数页", even_text), ("上次使用范围", last_text)):
            ttk.Button(preset_row, text=label, style="Ghost.TButton",
                       command=lambda v=value: pick(v)).pack(side="left", padx=(0, 4))

        ttk.Label(dialog, text="范围（绝对页码，留空表示按上方范围打印）：").grid(
            row=1, column=0, sticky="w", padx=10, pady=(6, 4))
        range_entry = ttk.Entry(dialog, textvariable=range_var, width=32)
        range_entry.grid(row=2, column=0, sticky="ew", padx=10)
        ttk.Label(dialog, text="份数（留空表示使用上方份数）：").grid(
            row=3, column=0, sticky="w", padx=10, pady=(8, 4))
        copies_spin = ttk.Spinbox(dialog, from_=1, to=99, textvariable=copies_var, width=8)
        copies_spin.grid(row=4, column=0, sticky="w", padx=10)
        scope_error_var = tk.StringVar()
        ttk.Label(dialog, textvariable=scope_error_var,
                  foreground=self.colors["danger"]).grid(
            row=5, column=0, sticky="w", padx=10, pady=(4, 0))
        summary_var = tk.StringVar()
        ttk.Label(dialog, textvariable=summary_var, foreground=self.colors["muted"],
                  justify="left", wraplength=420).grid(
            row=6, column=0, sticky="w", padx=10, pady=(6, 0))
        ttk.Label(dialog, text="将应用到当前选中的 %d 个文件；双击单行可快速修改同一设置。"
                  % len(selected), foreground=self.colors["muted"], justify="left",
                  wraplength=420).grid(row=7, column=0, sticky="w", padx=10, pady=(4, 0))

        def preview_text() -> None:
            """实时显示本次选取页数/补页/输出面，非法范围只在这里标红。"""
            text = range_var.get().strip()
            copies = self._int_setting("copies", 1)
            if copies_var.get().strip():
                try:
                    copies = max(1, min(99, int(copies_var.get())))
                except ValueError:
                    copies = 1
            error_text = ""
            if text:
                try:
                    pages = parse_page_range(text)
                except PdfError as exc:
                    error_text = f"范围有误：{str(exc).splitlines()[0]}"
                    pages = []
            else:
                try:
                    pages, _dropped, copies = self._resolve_record(selected[0])
                except PdfError:
                    pages, copies = [], 1
            if error_text:
                scope_error_var.set(error_text)
                range_entry.configure(style="Error.TEntry")
                summary_var.set("修正范围后才能应用；其它文件不受影响。")
                return
            scope_error_var.set("")
            range_entry.configure(style="TEntry")
            if not pages:
                summary_var.set("该范围在所选文件中没有有效页（打印时会被跳过）。")
                return
            plan = build_layout_plan(len(pages), self._layout_options())
            summary_var.set(
                f"本次选取 {len(pages)} 页；补 {plan.source_blank_added} 页后 "
                f"{plan.padded_count} 页；输出 {plan.output_count} 面、"
                f"{plan.output_count // 2} 张纸 × {copies} 份。")

        range_var.trace_add("write", lambda *_a: preview_text())
        copies_var.trace_add("write", lambda *_a: preview_text())

        def collect():
            """校验输入；合法返回 (范围文本, 份数)，非法返回 None（只在对话框标红）。"""
            text = range_var.get().strip()
            if text:
                try:
                    parse_page_range(text)
                except PdfError as exc:
                    scope_error_var.set(f"范围有误：{str(exc).splitlines()[0]}")
                    range_entry.configure(style="Error.TEntry")
                    return None
            copies = None
            if copies_var.get().strip():
                try:
                    copies = int(copies_var.get())
                    if not 1 <= copies <= 99:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("份数有误", "份数需为 1-99 的整数。", parent=dialog)
                    return None
            return text, copies

        def apply() -> None:
            values = collect()
            if values is None:
                return
            text, copies = values
            dialog.destroy()
            self._apply_scope_to_records(selected, text, copies)

        def apply_all() -> None:
            """把当前范围/份数一次性写到列表内所有文件。"""
            values = collect()
            if values is None:
                return
            text, copies = values
            dialog.destroy()
            self._apply_scope_to_records(self.files, text, copies)

        def restore_inherit() -> None:
            self.vars["per_file"].set(True)
            for record in selected:
                record["range_text"] = ""
                record["copies"] = None
            self._append_log(f"已恢复 {len(selected)} 个文件为按上方范围打印。")
            dialog.destroy()
            self._sync_widgets()
            self._update_preview()

        buttons = ttk.Frame(dialog)
        buttons.grid(row=8, column=0, sticky="ew", padx=10, pady=10)
        ttk.Button(buttons, text="确定", style="Accent.TButton", command=apply).pack(side="left")
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="恢复按上方范围", style="Ghost.TButton",
                   command=restore_inherit).pack(side="right")
        ttk.Button(buttons, text="应用到全部文件", style="Ghost.TButton",
                   command=apply_all).pack(side="right", padx=(0, 8))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        preview_text()
        dialog.update_idletasks()
        self._place_toplevel_over_root(dialog, dialog.winfo_reqwidth() + 24,
                                       dialog.winfo_reqheight() + 24)
        dialog.deiconify()
        dialog.lift()
        dialog.grab_set()
        target = {"range": range_entry, "copies": copies_spin}.get(focus, range_entry)
        target.focus_set()

    # ------------------------------------------------------------- 选中操作条
    def _on_tree_select(self, _event=None) -> None:
        self._update_selection_bar()
        self._update_preview()

    def _select_all_on_tree(self, _event=None) -> str:
        self.tree.selection_set([record["iid"] for record in self.files])
        self._update_selection_bar()
        return "break"

    def _clear_tree_selection(self, _event=None) -> str:
        self.tree.selection_remove(*self.tree.selection())
        self._update_selection_bar()
        return "break"

    def _update_selection_bar(self) -> None:
        """选中操作条：显示单选文件名或多选计数，无选中时按钮置灰。"""
        if not getattr(self, "sel_info_var", None):
            return
        records = self._selected_records()
        if not records:
            self.sel_info_var.set("未选择文件")
        elif len(records) == 1:
            name = os.path.basename(records[0]["path"])
            # 独占一行且给足宽度：超过 26 字才省略，省略号前保留完整短名
            self.sel_info_var.set(f"已选：{name if len(name) <= 26 else name[:25] + '…'}")
        else:
            self.sel_info_var.set(f"已选 {len(records)} 个文件")
        state = ["!disabled"] if records else ["disabled"]
        for button in getattr(self, "sel_buttons", ()):
            button.state(state)

    def _apply_scope_to_records(self, records: List[dict], text: str,
                                copies: Optional[int]) -> None:
        """一次性把范围/份数应用到指定记录，并写日志、刷新行与顶部统计。"""
        self.vars["per_file"].set(True)
        for record in records:
            record["range_text"] = text
            record["copies"] = copies
        self._append_log(
            f"已为 {len(records)} 个文件设置范围 {text or '全部'}、"
            f"份数 {copies if copies is not None else '使用上方份数'}。")
        self._sync_widgets()
        self._update_preview()

    def restore_global_for_selected(self) -> None:
        """把选中文件恢复为“按上方范围打印”（清空逐文件设置）。"""
        records = self._selected_records()
        if not records:
            return
        for record in records:
            record["range_text"] = ""
            record["copies"] = None
        self._append_log(f"已恢复 {len(records)} 个文件为按上方范围打印。")
        self._sync_widgets()
        self._update_preview()

    def apply_selected_scope_to_all(self) -> None:
        """把首个选中文件的设置复制到全部文件。"""
        records = self._selected_records()
        if not records:
            return
        source = records[0]
        text = str(source.get("range_text") or "")
        copies = source.get("copies")
        self.vars["per_file"].set(True)
        for record in self.files:
            record["range_text"] = text
            record["copies"] = copies
        self._append_log(
            f"已将《{os.path.basename(source['path'])}》的范围 {text or '全部'}、"
            f"份数 {copies if copies is not None else '使用上方份数'} "
            f"应用到全部 {len(self.files)} 个文件。")
        self._sync_widgets()
        self._update_preview()

    # ------------------------------------------------------------- 范围与计划计算
    def _record_has_custom_scope(self, record: dict) -> bool:
        """该文件是否单独设置过范围或份数（用于“＊”标记与浅灰整行）。"""
        if not bool(self.vars["per_file"].get()):
            return False
        return bool(str(record.get("range_text") or "").strip() or record.get("copies"))

    def _record_scope_invalid(self, record: dict) -> bool:
        """该文件是否设置了非法范围（只影响本行，不打断整批）。

        合法但超出页数的范围不算“非法”（会按规则裁剪并提示补页）。
        """
        if not (bool(self.vars["per_file"].get()) and record.get("range_text")):
            return False
        try:
            parse_page_range(str(record["range_text"]))
        except PdfError:
            return True
        return False

    def _resolve_record(self, record: dict, current_page: Optional[int] = None
                        ) -> Tuple[List[int], List[int], int]:
        """返回（有效绝对页码, 被忽略页码, 份数）。"""
        page_count = record.get("pages") or 0
        mode = str(self.vars["range_mode"].get() or "all")
        text = self._effective_range_text(mode)
        if bool(self.vars["per_file"].get()) and record.get("range_text"):
            mode = "custom"
            text = record["range_text"]
        copies = self._int_setting("copies", 1)
        if bool(self.vars["per_file"].get()) and record.get("copies"):
            copies = int(record["copies"])
        if not page_count:
            return [], [], max(1, copies)
        try:
            valid, dropped = resolve_range_pages(
                mode, text, page_count,
                current_page if current_page is not None else self.preview_page_index or 1,
            )
        except PdfError:
            raise
        return valid, dropped, max(1, copies)

    def _scope_text(self, record: dict) -> str:
        """“本次范围”列：未单设的文件直接显示实际生效的范围文本。

        逐文件开关关闭时该列不显示；开启后：
          * 未单独设置 → 实际生效范围（如“全部”），整行浅灰（按上方范围）；
          * 已单独设置 → “＊3-7”/“＊3-7×2”（＊ 前缀 + 深色，表示单独设置过）；
          * 非法 → “＊abc 有误”；越界 → “＊9-10 越界”。
        文案都控制在短句内，1024 宽下不出现半截文字。
        """
        mode = str(self.vars["range_mode"].get() or "all")
        labels = {
            "all": "全部", "current": f"当前页 {self.preview_page_index or 1}",
            "custom": self._effective_range_text("custom") or "全部",
            "odd": "仅奇数页", "even": "仅偶数页",
            "recent": "最近 " + (self._effective_range_text("recent") or "—"),
        }
        if not bool(self.vars["per_file"].get()):
            return labels.get(mode, mode)
        custom_range = str(record.get("range_text") or "").strip()
        custom_copies = record.get("copies")
        if not custom_range and not custom_copies:
            return labels.get(mode, mode)
        if custom_range:
            try:
                parse_page_range(custom_range)
            except PdfError:
                text = f"＊{custom_range} 有误"
            else:
                valid, dropped = resolve_range_pages(
                    "custom", custom_range, int(record.get("pages") or 0),
                    self.preview_page_index or 1)
                if not valid:
                    text = f"＊{custom_range} 越界"
                else:
                    text = f"＊{custom_range}"
        else:
            text = "＊全部"
        if custom_copies:
            text += f"×{int(custom_copies)}"
        return text

    def _plan_text(self, record: dict) -> str:
        """文件列表“输出与补页”列：短文案，列宽 1024 下也放得下。"""
        if not record.get("pages"):
            return "-"
        try:
            valid, dropped, copies = self._resolve_record(record)
        except PdfError:
            return "范围有误"
        if not valid:
            return "范围外"
        plan = build_layout_plan(len(valid), self._layout_options())
        # “总页”列已显示页数，这里只写补页与输出面；例：→6面 / 补1→6面 / 2合1→6面
        text = ""
        if plan.source_blank_added:
            text += f"补{plan.source_blank_added}"
        if plan.booklet:
            text += "册"
        elif plan.nup > 1:
            text += f"{plan.nup}合1"
        text += f"→{plan.output_count}面"
        if copies > 1:
            text += f"×{copies}"
        if dropped:
            text += f" 略{len(dropped)}"
        return text.strip()

    def _update_empty_state(self) -> None:
        """空列表引导与表格的显隐切换（同一单元格，互斥显示，不产生重叠）。"""
        frame = getattr(self, "empty_state", None)
        tree = getattr(self, "tree", None)
        if frame is None or tree is None:
            return
        scroll = getattr(self, "tree_scroll", None)
        if self.files:
            frame.grid_remove()
            tree.grid()
            if scroll is not None:
                scroll.grid()
        else:
            tree.grid_remove()
            frame.grid()
            if scroll is not None:
                scroll.grid_remove()

    def _refresh_file_rows(self) -> None:
        if hasattr(self, "file_count_var"):
            self.file_count_var.set(f"{len(self.files)} 个文件")
        self._update_empty_state()
        for record in self.files:
            if not self.tree.exists(record["iid"]):
                continue
            values = list(self.tree.item(record["iid"], "values"))
            while len(values) < 5:
                values.append("")
            values[1] = record["pages"] if record.get("pages") else "?"
            values[2] = self._scope_text(record)
            values[3] = self._plan_text(record)
            self.tree.item(record["iid"], values=values)
        self._apply_row_tags()
        # 行内容变化后强制重算（宽度不变也要按新的最长文案调整）
        self._last_tree_width = -1
        self._fit_tree_columns()
        self._update_scope_hint()
        self._update_flow_steps()
        self._update_start_state()

    # --------------------------------------------------------------- 预览
    def _current_preview_record(self) -> Optional[dict]:
        usable = [f for f in self.files if f.get("pages")]
        if not usable:
            return None
        selected = set(self.tree.selection())
        for record in usable:
            if record["iid"] in selected:
                return record
        return usable[0]

    def _preview_error(self, message: str) -> None:
        self.preview_title_label.configure(foreground=self.colors["danger"])
        self.preview_title_var.set(message)

    def _update_preview(self) -> None:
        """刷新文件列、双面页序示意与每个文件汇总，并触发版式预览。"""
        try:
            # 每次交互刷新前先做一次轻量的 mtime/大小校验（命中缓存则零开销），
            # 外部修改过的 PDF 会自动重新读取页数并更新状态列。
            self._refresh_file_metadata(notify=False)
            self._refresh_file_rows()
            self._update_metrics()
            self._update_batch_plan()
            for widget in self.preview_box_frame.winfo_children():
                widget.destroy()

            record = self._current_preview_record()
            if record is None and self.files:
                # 全是未转换的 Word 时也要给出“待转换”提示，而不是“添加文件”
                record = self._selected_record() or self.files[0]
            if record is None:
                self.preview_title_label.configure(foreground=self.colors["accent"])
                self.preview_title_var.set("添加文件后，这里会以示意图显示每张纸的正反面页码。")
                self._schedule_layout_preview()
                return

            name = os.path.basename(record["path"])
            if record.get("kind") in ("word", "wps", "image") and not record.get("pdf_path"):
                self.preview_title_label.configure(foreground=self.colors["muted"])
                if record.get("status") == "转换失败":
                    self.preview_title_var.set(
                        f"《{name}》转换失败：请在日志中查看原因，或手动另存为 PDF 后重新导入。")
                else:
                    kind_text = {"image": "图片", "wps": "WPS 文档"}.get(
                        record.get("kind"), "Word 文档")
                    self.preview_title_var.set(f"《{name}》是{kind_text}，正在转换为 PDF…")
                self._clear_preview_images()
                self.preview_info_var.set("待转换")
                return
            try:
                abs_pages, dropped, copies = self._resolve_record(record)
            except PdfError as exc:
                self._preview_error(f"页码范围有误：{str(exc).splitlines()[0]}")
                self._clear_preview_images()
                return
            if not abs_pages:
                self._preview_error(f"范围全部超出《{name}》的页数（共 {record['pages']} 页）。")
                self._clear_preview_images()
                return

            layout = self._layout_options()
            plan = build_layout_plan(len(abs_pages), layout)
            title = f"以《{name}》为例：选取 {len(abs_pages)} 页"
            if plan.source_blank_added:
                title += f"，补 {plan.source_blank_added} 张空白页 → {plan.padded_count} 页"
            if plan.booklet:
                title += f"；小册子排版 → {plan.output_count} 个输出面"
            elif plan.nup > 1:
                title += f"；{plan.nup} 合 1 → {plan.output_count} 个输出面"
            else:
                title += f"；共 {plan.output_count} 个输出面"
            if plan.output_blank_added:
                title += "（另补 1 个空白输出面用于双面配对）"
            if copies > 1:
                title += f"；以下页序共重复 {copies} 份"
            if dropped:
                title += f"；{len(dropped)} 个超出页数已忽略"
            if len([f for f in self.files if f.get("pages")]) > 1:
                title += "（其他文件按各自页数同理计算）"
            self.preview_title_label.configure(foreground=self.colors["accent"])
            self.preview_title_var.set(title)

            self._build_duplex_boxes(plan, abs_pages, layout)
            # 版式未变（例如只拉伸窗口/切换单双页）时跳过重开 PDF，仅由 Canvas 重绘
            self._schedule_layout_preview(self._preview_signature_for(record, abs_pages, layout))
        except Exception as exc:  # 预览不允许把主界面拖垮
            self._append_log("刷新预览时出错：\n" + traceback.format_exc())
            self._preview_error(f"预览出错：{exc}")

    # 兼容旧命名
    def _refresh_plan_column(self) -> None:
        self._refresh_file_rows()

    def _update_metrics(self) -> None:
        """刷新顶部指标徽标：文件数 / 选取页 / 预计张数 / 翻面次数。

        指标与“整批一次翻面”规则一致：
          * 每个文件各自补空白页保证输出面为偶数；
          * 整批模式翻面 1 次（按份逐份执行时为份数）；
          * 逐文件模式翻面次数 = 有可打印内容的文件数；
          * 单面打印翻面 0 次。
        """
        usable = []
        for record in self.files:
            if not record.get("pages"):
                continue
            try:
                valid, _dropped, copies = self._resolve_record(record)
            except PdfError:
                continue
            if not valid:
                continue
            plan = build_layout_plan(len(valid), self._layout_options())
            usable.append(max(1, int(copies)) * plan.output_count)

        files_count = len(usable)
        pages_count = 0
        for record in self.files:
            if not record.get("pages"):
                continue
            try:
                valid, _dropped, _copies = self._resolve_record(record)
            except PdfError:
                continue
            pages_count += len(valid)

        flow = str(self.vars["print_flow"].get() or "batch_all")
        if not bool(self.vars["manual_duplex"].get()):
            sheets = sum(usable)
            flips = 0
        elif flow in ("batch_all", "batch_take"):
            sheets = sum(usable) // 2
            copies = self._int_setting("copies", 1)
            flips = copies if (copies > 1 and bool(self.vars["copy_exec_on"].get())) else 1
            if files_count == 0:
                flips = 0
        else:
            sheets = sum(usable) // 2
            flips = files_count

        self.metric_vars["files"].set(str(files_count))
        self.metric_vars["pages"].set(str(pages_count))
        self.metric_vars["sheets"].set(str(sheets))
        self.metric_vars["flips"].set(str(flips))

    def _update_batch_plan(self) -> None:
        """预览页的整批视图：

        * 阶段一：合并任务全局奇数位置页（按文件分组展示）；
        * 阶段二：翻面后对全局页序整体倒序得到的偶数位置页；
        * 同时显示“预计翻面 N 次、共 M 张纸”。
        """
        counts: List[int] = []
        names: List[str] = []
        single_sided = not bool(self.vars["manual_duplex"].get())
        for record in self.files:
            if not record.get("pages"):
                continue
            try:
                valid, _dropped, copies = self._resolve_record(record)
            except PdfError:
                continue
            if not valid:
                continue
            plan = build_layout_plan(len(valid), self._layout_options())
            counts.append(plan.output_count)
            names.append(os.path.basename(record["path"]))

        if not counts:
            self.batch_plan_var.set("添加文件后，这里显示整批两阶段页序与翻面次数。")
            return
        if single_sided or sum(counts) <= 0:
            self.batch_plan_var.set(
                "单面打印：不进行翻面。\n"
                f"合并后共 {sum(counts)} 个输出面；勾选“手动双面”后显示整批两阶段页序。")
            return
        try:
            flow = str(self.vars["print_flow"].get() or "batch_all")
            summary = plan_batch_summary(
                counts, self._int_setting("copies", 1),
                exec_mode="by_file" if flow == "per_file" else "batch",
                copy_exec="per_copy" if bool(self.vars["copy_exec_on"].get()) else "all",
                reverse=bool(self.vars["reverse"].get()),
            )
        except PdfError:
            self.batch_plan_var.set("当前设置无法计算整批页序，请检查范围与份数。")
            return

        ranges = merged_side_ranges(counts)
        total = sum(counts)
        lines = [
            f"{summary['note']}｜预计翻面 {summary['flip_count']} 次、共 "
            f"{summary['sheets']} 张纸｜共 {total} 个输出面",
            f"阶段一（正面，全局正序）：{format_page_list(summary['fronts'], limit=14)}",
            f"阶段二（翻面后全局逆序）：{format_page_list(summary['backs'], limit=14)}",
        ]
        # 按文件分组：便于核对文件边界（每个文件都会从奇数全局页开始）。
        # 文件多于 1 个时合并成一行展示，避免把预览大图挤出可视区域。
        if len(counts) > 1:
            groups = []
            for (start, end), name in zip(ranges, names):
                local_fronts = [p for p in summary["fronts"] if start <= p <= end]
                local_backs = [p for p in summary["backs"] if start <= p <= end]
                groups.append(
                    f"《{name}》{start}-{end}：正 "
                    f"{format_page_list(local_fronts, limit=6)}｜反 "
                    f"{format_page_list(local_backs, limit=6)}")
            lines.append("文件分组：" + "　".join(groups))
        lines.append("每个文件已各自补页，文件边界不会跨到下一个文件的纸背面。")
        self.batch_plan_var.set("\n".join(lines))

    def _build_duplex_boxes(self, plan, abs_pages: List[int], layout: LayoutOptions) -> None:
        total_sides = plan.output_count
        phases = ((1, "第一面 · 正序"), (2, "第二面 · 翻面后逆序"))
        for phase, label in phases:
            row = tk.Frame(self.preview_box_frame, background=self.colors["surface"])
            row.pack(anchor="w", pady=1)
            tk.Label(row, text=label, font=("Microsoft YaHei UI", 9), width=18,
                     anchor="e", background=self.colors["surface"],
                     foreground=self.colors["muted"]).pack(side="left", padx=(0, 8))
            numbers = [n for n in range(1, total_sides + 1)
                       if (n % 2 == 1 if phase == 1 else n % 2 == 0)]
            if phase == 2:
                numbers.reverse()
            cap = 16
            for side_no in numbers[:cap]:
                self._side_box(row, plan, side_no, abs_pages, back=(phase == 2))
            if len(numbers) > cap:
                tk.Label(row, text=f"… 共 {len(numbers)} 面", font=("Microsoft YaHei UI", 8),
                         fg=self.colors["muted"],
                         background=self.colors["surface"]).pack(side="left", padx=4)
            if phase == 1 and plan.manual_duplex:
                tk.Label(self.preview_box_frame,
                         text="↓ 第一面打印完成后，将整叠纸【整体翻面】放回进纸盘 ↓",
                         font=("Microsoft YaHei UI", 8), fg=self.colors["warning"],
                         background=self.colors["surface"]).pack(anchor="w", pady=(2, 2))

    def _side_box(self, parent: tk.Frame, plan, side_no: int,
                  abs_pages: List[int], back: bool) -> None:
        entries = plan.sides[side_no - 1]
        parts = []
        for rel in entries:
            if rel is None:
                continue
            if rel > len(abs_pages):
                parts.append("空白(补)")
            elif plan.nup == 1 and not plan.booklet:
                parts.append(f"原{abs_pages[rel - 1]}")
            else:
                parts.append(f"原{abs_pages[rel - 1]}")
        if not parts:
            text = f"{side_no} · 空白"
            bg, fg = self.colors["surface_alt"], self.colors["muted"]
        else:
            # 单行紧凑排版：面号 + 该面上的原文档页码，降低双面示意区高度
            text = f"{side_no} · " + "、".join(parts)
            bg = self.colors["accent_soft"] if back else self.colors["surface"]
            fg = self.colors["text"]
        tk.Label(parent, text=text, font=("Microsoft YaHei UI", 8), width=10,
                 relief="flat", borderwidth=1, bg=bg, fg=fg,
                 highlightthickness=1,
                 highlightbackground=self.colors["border"]).pack(side="left", padx=2, pady=1)

    # ------------------------------------------------------------ 成纸预览
    def _clear_preview_images(self) -> None:
        """清空成纸预览：Canvas、缩略图、位图缓存与页码显示。"""
        self.preview_photo = None
        self.preview_pngs.clear()
        self.preview_pages = 0
        self.preview_page_index = 0
        self.preview_metrics_data = {}
        self._paper_photo_cache.clear()
        for widget in self.thumb_frame.winfo_children():
            widget.destroy()
        self._thumb_photos = []
        try:
            self.preview_canvas.delete("all")
        except tk.TclError:
            pass
        self.preview_page_var.set("第 0 / 0 面")
        self.actual_scale_var.set("实际 —")
        if hasattr(self, "preview_page_entry_var"):
            self.preview_page_entry_var.set("1")
        self._draw_paper_preview()

    def _preview_signature_for(self, record: dict, abs_pages: List[int],
                               layout: LayoutOptions):
        """预览签名：文件、页面与版式参数的组合，用于避免无意义的重新渲染。"""
        return (record.get("path", ""), tuple(abs_pages), repr(layout))

    def _schedule_layout_preview(self, signature=None) -> None:
        if not self.preview_enabled:
            return
        # 版式/页面未变且已有位图：不重开 PDF，交由 _draw_paper_preview 重绘
        if signature is not None and signature == self._preview_signature and self.preview_pngs:
            return
        self._preview_signature = signature
        if self.preview_after_id is not None:
            try:
                self.root.after_cancel(self.preview_after_id)
            except tk.TclError:
                pass
        self.preview_after_id = self.root.after(500, self._refresh_layout_preview)

    def _refresh_layout_preview(self) -> None:
        """立即（后台线程）重新生成成纸预览。"""
        self.preview_after_id = None
        if not self.preview_enabled:
            return
        record = self._current_preview_record()
        if record is None:
            self._clear_preview_images()
            return
        try:
            abs_pages, _dropped, _copies = self._resolve_record(record)
        except PdfError:
            self._clear_preview_images()
            return
        if not abs_pages:
            self._clear_preview_images()
            return
        self.preview_generation += 1
        generation = self.preview_generation
        layout = self._layout_options()
        pdf_path = self._record_pdf_path(record)
        if not pdf_path or not os.path.isfile(pdf_path):
            self._clear_preview_images()
            return
        self.preview_info_var.set("正在生成成纸预览…")
        self.preview_thread = threading.Thread(
            target=self._preview_worker,
            args=(pdf_path, abs_pages, layout, generation),
            daemon=True,
        )
        self.preview_thread.start()

    def _preview_worker(self, src_path: str, abs_pages: List[int],
                        layout: LayoutOptions, generation: int) -> None:
        """后台生成“真实成纸”位图与几何数据，避免界面冻结。

        * 用 PyMuPDF 把排版后的 PDF 逐面渲染为 PNG（预览即是最终打印效果，
          包含边距留白、缩放、N-up 分格、溢出裁切）；
        * 同时用 preview_metrics 计算纸张、可打印区、槽位与实际缩放百分比，
          交给 Canvas 绘制纸张轮廓与虚线；
        * 只渲染前 PREVIEW_MAX_THUMBS 面，超出部分仅提示，避免超大文档卡顿。
        """
        temp_dir = ""
        try:
            import fitz
            temp_dir = tempfile.mkdtemp(prefix="pdf_duplex_preview_")
            out_path = os.path.join(temp_dir, "preview.pdf")
            result = compose_layout_pdf(src_path, abs_pages, out_path, layout)

            # 取第一张源页的显示尺寸（已考虑 /Rotate），用于计算缩放百分比
            src_w, src_h = 595.0, 842.0
            try:
                src_doc = fitz.open(src_path)
                src_page = src_doc[abs_pages[0] - 1]
                src_w, src_h = src_page.rect.width, src_page.rect.height
                src_doc.close()
            except Exception:
                pass
            metrics = preview_metrics(result.plan, layout, src_w, src_h)

            doc = fitz.open(out_path)
            page_count = min(doc.page_count, PREVIEW_MAX_THUMBS)
            pngs: List[bytes] = []
            thumbs: List[bytes] = []
            paper_pts: List[Tuple[float, float]] = []
            for index in range(page_count):
                page = doc[index]
                paper_pts.append((page.rect.width, page.rect.height))
                # 按最长边渲染：纸张轮廓清晰，Canvas 再按控件大小缩放（可缓存）
                zoom = PREVIEW_LONG_EDGE / max(1.0, page.rect.width, page.rect.height)
                pngs.append(page.get_pixmap(
                    matrix=fitz.Matrix(zoom, zoom), alpha=False).tobytes("png"))
                tzoom = 72.0 / max(1.0, page.rect.width)
                thumbs.append(page.get_pixmap(
                    matrix=fitz.Matrix(tzoom, tzoom), alpha=False).tobytes("png"))
            total = doc.page_count
            doc.close()
            self.ui_queue.put(("preview", generation, {
                "temp_dir": temp_dir, "pages": total, "pngs": pngs, "thumbs": thumbs,
                "info": describe_plan(result.plan), "metrics": metrics,
                "paper_pts": paper_pts,
            }))
        except ToolError as exc:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.ui_queue.put(("preview_error", generation, str(exc)))
        except Exception as exc:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.ui_queue.put(("preview_error", generation, f"生成预览失败：{exc}"))

    def _on_preview_data(self, payload: dict) -> None:
        old_dir = self.preview_temp_dir
        self.preview_temp_dir = payload["temp_dir"]
        self.preview_pages = int(payload["pages"])
        self.preview_pngs = {i: data for i, data in enumerate(payload["pngs"])}
        self.preview_info_var.set(str(payload.get("info", "")))
        self.preview_metrics_data = dict(payload.get("metrics") or {})
        self._draw_advanced_preview()
        self.preview_paper_pts = list(payload.get("paper_pts") or [])
        self._paper_photo_cache.clear()
        if old_dir and old_dir != self.preview_temp_dir:
            shutil.rmtree(old_dir, ignore_errors=True)
        for widget in self.thumb_frame.winfo_children():
            widget.destroy()
        self._thumb_photos = []
        # 缩略图每行 3 个、最多展示 1 行：超出部分通过翻页/页码跳转查看，
        # 避免几十页的缩略图把预览舞台高度挤成 0，也避免挤压文件汇总表
        for index, data in enumerate(payload["thumbs"][:3]):
            try:
                photo = tk.PhotoImage(data=data)
                if photo.height() > 64:
                    photo = photo.subsample(2)
            except tk.TclError:
                continue
            self._thumb_photos.append(photo)
            label = tk.Label(self.thumb_frame, image=photo, relief="solid", borderwidth=1,
                             cursor="hand2")
            label.grid(row=index // 3, column=index % 3, padx=2, pady=2)
            label.bind("<Button-1>", lambda _e, i=index: self._show_preview_page(i))
        self.preview_page_index = min(self.preview_page_index, max(0, self.preview_pages - 1))
        self._update_actual_scale_text()
        self._show_preview_page(self.preview_page_index)

    def _update_actual_scale_text(self) -> None:
        """显示当前缩放模式下的实际百分比（例如“实际 87%”）。"""
        metrics = self.preview_metrics_data or {}
        percent = float(metrics.get("scale_percent") or 0.0)
        if percent <= 0:
            self.actual_scale_var.set("实际 —")
            return
        text = f"实际 {percent:.0f}%"
        if metrics.get("overflow"):
            text += "（超出，已裁切）"
        self.actual_scale_var.set(text)
        if hasattr(self, "preview_status_scale_var"):
            self.preview_status_scale_var.set(f"缩放 {percent:.0f}%")

    def _set_preview_page(self, delta: int) -> None:
        if self.preview_pages <= 0:
            return
        step = 2 if str(self.preview_view_var.get()) == "spread" else 1
        self._show_preview_page(
            max(0, min(self.preview_page_index + delta * step, self.preview_pages - 1)))

    def _show_preview_page(self, index: int) -> None:
        """切换到指定输出面并立即重绘成纸预览（复用已缓存的位图）。"""
        if self.preview_pages <= 0:
            self.preview_page_index = 0
        else:
            self.preview_page_index = max(0, min(int(index), self.preview_pages - 1))
        self.preview_page_var.set(f"{self.preview_page_index + 1} / {self.preview_pages}")
        if hasattr(self, "preview_page_entry_var"):
            self.preview_page_entry_var.set(str(self.preview_page_index + 1))
        self._sync_widget_states()
        self._draw_paper_preview()

    def _get_paper_photo(self, index: int, width: int, height: int):
        """把渲染好的 PNG 按 Canvas 上的纸张像素尺寸缩放（带缓存）。"""
        key = (int(index), int(width), int(height))
        cached = self._paper_photo_cache.get(key)
        if cached is not None:
            return cached
        data = self.preview_pngs.get(index)
        if not data:
            return None
        photo = None
        try:
            import io
            from PIL import Image, ImageTk
            image = Image.open(io.BytesIO(data)).convert("RGB")
            image = image.resize((max(1, int(width)), max(1, int(height)),), Image.LANCZOS)
            photo = ImageTk.PhotoImage(image)
        except Exception:
            try:
                photo = tk.PhotoImage(data=data)
            except tk.TclError:
                return None
        # 缓存上限，避免窗口反复拉伸导致内存膨胀
        if len(self._paper_photo_cache) > 80:
            self._paper_photo_cache.clear()
        self._paper_photo_cache[key] = photo
        return photo

    def _draw_paper_preview(self) -> None:
        """在 Canvas 上绘制“纸张轮廓 + 可打印区虚线 + 页面实际落位”。

        绘制顺序（从下到上）：
          1. 纸张阴影与白色纸面（纸张边界为实线）；
          2. 排版后的页面位图（已包含缩放、边距留白、N-up 分格与裁切结果）；
          3. 可打印区域虚线（驱动返回的物理不可打印边距与用户边距的并集）；
          4. N-up/小册子的槽位分格虚线与槽位页序编号。
        单页/双页并排两种视图共用同一套缓存位图，切换时不重新打开 PDF。
        """
        canvas = getattr(self, "preview_canvas", None)
        if canvas is None:
            return
        try:
            canvas.delete("all")
            width = canvas.winfo_width()
            height = canvas.winfo_height()
        except tk.TclError:
            return
        if width < 80 or height < 80:
            return
        if self.preview_pages <= 0 or not self.preview_pngs:
            canvas.configure(background=self.colors["surface_alt"])
            self._draw_preview_empty(canvas, width, height)
            self._update_preview_status_strip([])
            return

        spread = str(self.preview_view_var.get()) == "spread"
        indexes = [self.preview_page_index]
        if spread and self.preview_page_index + 1 < self.preview_pages:
            indexes.append(self.preview_page_index + 1)

        margin = self._px(12)
        gap = self._px(16)
        # 单元格占满可用高度：纸张在单元格内等比缩放后水平/垂直居中，
        # 四周留白均匀，不会贴边、也不会偏上（面号标签落在上方留白里）。
        cell_w = (width - margin * 2 - gap * (len(indexes) - 1)) / len(indexes)
        cell_h = height - margin * 2
        for position, index in enumerate(indexes):
            self._draw_one_paper(canvas, index, margin + position * (cell_w + gap),
                                 margin, cell_w, cell_h)
        self._update_preview_status_strip(indexes)

    def _draw_preview_empty(self, canvas: tk.Canvas, width: int, height: int) -> None:
        """空状态：虚线拖拽框 + 纸张图形 + 引导文案，而不是一块空灰。"""
        c = self.colors
        has_files = bool(self.files)
        box_w = min(width - self._px(40), self._px(360))
        box_h = min(height - self._px(48), self._px(210))
        x0, y0 = (width - box_w) / 2, (height - box_h) / 2
        x1, y1 = x0 + box_w, y0 + box_h
        canvas.create_rectangle(x0, y0, x1, y1, outline=c["border"], dash=(6, 4), width=1)

        # 纸张图形：白纸 + 折角 + 三条正文占位线（纯 Canvas 绘制，无位图依赖）
        sheet_w = self._px(54)
        sheet_h = self._px(70)
        sx, sy = width / 2 - sheet_w / 2, y0 + self._px(22)
        canvas.create_rectangle(sx, sy, sx + sheet_w, sy + sheet_h,
                                fill=c["surface"], outline=c["border"], width=1)
        fold = self._px(14)
        canvas.create_polygon(sx + sheet_w - fold, sy, sx + sheet_w, sy + fold,
                              sx + sheet_w - fold, sy + fold,
                              fill=c["surface_alt"], outline=c["border"])
        for line in range(3):
            ly = sy + self._px(30) + line * self._px(10)
            canvas.create_line(sx + self._px(10), ly, sx + sheet_w - self._px(10), ly,
                               fill=c["border"])
        canvas.create_text(width / 2, sy + sheet_h + self._px(18),
                           text="将 PDF / Word / 图片拖到这里" if not has_files else "正在生成预览…",
                           fill=c["text"], font=(_theme.FONT_FAMILY, 10, "bold"))
        canvas.create_text(
            width / 2, sy + sheet_h + self._px(40),
            text="或点击左侧「＋ 添加文件」\nWord / 图片会自动转成 PDF 再打印",
            fill=c["muted"], font=_theme.FONT_SMALL, justify="center", anchor="n")

    def _update_preview_status_strip(self, indexes: List[int]) -> None:
        """预览底部细状态条：页码 / 缩放 / 纸张（长度变化不改变布局）。"""
        if not hasattr(self, "preview_status_page_var"):
            return
        metrics = self.preview_metrics_data or {}
        if not indexes:
            self.preview_status_page_var.set("等待预览")
            self.preview_status_scale_var.set("缩放 —")
            self.preview_status_paper_var.set("纸张 —")
            if not self.files:
                self.preview_info_var.set("")
            return
        page_text = f"第 {indexes[0] + 1} 面"
        if len(indexes) > 1:
            page_text += f"–{indexes[-1] + 1} 面"
        page_text += f" / 共 {self.preview_pages} 面"
        percent = float(metrics.get("scale_percent") or 0.0)
        scale_text = f"缩放 {percent:.0f}%" if percent > 0 else "缩放 —"
        paper_w, paper_h = self._paper_size_for(indexes[0])
        paper_text = f"纸张 {paper_w / 72 * 25.4:.0f}×{paper_h / 72 * 25.4:.0f} mm"
        nup = int(metrics.get("nup") or 1)
        if nup > 1:
            paper_text += f" · {nup} 合 1"
        self.preview_status_page_var.set(page_text)
        self.preview_status_scale_var.set(scale_text)
        self.preview_status_paper_var.set(paper_text)

    def _draw_one_paper(self, canvas: tk.Canvas, index: int,
                        cell_x: float, cell_y: float, cell_w: float, cell_h: float) -> None:
        """绘制一张纸：纸面、页面位图、可打印区虚线、N-up 槽位。"""
        metrics = self.preview_metrics_data or {}
        paper_w_pt, paper_h_pt = self._paper_size_for(index)
        aspect = paper_w_pt / max(1.0, paper_h_pt)
        # 等比缩放：四周留 16px，另在上方多留 18px 放“第 N 面”标签；
        # 纸张整体在单元格内水平/垂直居中（不贴边、不偏上）。
        pad = self._px(16)
        label_band = self._px(18)
        usable_w = max(40.0, cell_w - pad * 2)
        usable_h = max(40.0, cell_h - pad * 2 - label_band)
        paper_h = usable_h
        paper_w = paper_h * aspect
        if paper_w > usable_w:
            paper_w = usable_w
            paper_h = paper_w / max(0.01, aspect)
        x0 = cell_x + (cell_w - paper_w) / 2
        y0 = cell_y + (cell_h - paper_h) / 2
        x1, y1 = x0 + paper_w, y0 + paper_h

        # 轻投影（2px、比边框更浅）+ 纸面（1px 浅灰描边，抢眼程度低于文字）
        canvas.create_rectangle(x0 + self._px(2), y0 + self._px(2),
                                x1 + self._px(2), y1 + self._px(2),
                                fill=self.colors["shadow_light"], outline="",
                                tags=("paper_shadow",))
        canvas.create_rectangle(x0, y0, x1, y1, fill=self.colors["surface"],
                                outline=self.colors["border"], width=1, tags=("paper",))

        photo = self._get_paper_photo(index, int(paper_w), int(paper_h))
        if photo is not None:
            canvas.create_image(x0, y0, anchor="nw", image=photo)
        else:
            canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2,
                               text="该面尚未渲染\n（仅预览前 %d 面）" % PREVIEW_MAX_THUMBS,
                               fill=self.colors["muted"], justify="center",
                               font=(_theme.FONT_FAMILY, 9))

        def _map(rect_pt):
            """纸张坐标（点）→ Canvas 坐标。"""
            px = x0 + float(rect_pt[0]) / max(1.0, paper_w_pt) * paper_w
            py = y0 + float(rect_pt[1]) / max(1.0, paper_h_pt) * paper_h
            return px, py

        # 可打印区域虚线（物理不可打印边距 + 用户边距）
        printable = metrics.get("printable")
        if printable:
            rx0, ry0 = _map((printable[0], printable[1]))
            rx1, ry1 = _map((printable[2], printable[3]))
            canvas.create_rectangle(rx0, ry0, rx1, ry1, outline=self.colors["accent"],
                                    dash=(5, 3), width=1)

        # N-up / 小册子：槽位分格虚线 + 槽位页序编号
        slots = metrics.get("slots") or []
        if int(metrics.get("nup") or 1) > 1 and slots:
            for slot_no, slot in enumerate(slots, start=1):
                sx0, sy0 = _map((slot[0], slot[1]))
                sx1, sy1 = _map((slot[2], slot[3]))
                canvas.create_rectangle(sx0, sy0, sx1, sy1, outline=self.colors["muted"],
                                        dash=(3, 3), width=1)
                canvas.create_text(sx0 + 6, sy0 + 6, anchor="nw", text=str(slot_no),
                                   fill=self.colors["accent"],
                                   font=(_theme.FONT_FAMILY, 8, "bold"))

        # 面号标签（纸张上方留白内；随纸张居中，不贴画布顶边）
        canvas.create_text(x0, y0 - 3, anchor="sw", text=f"第 {index + 1} 面",
                           fill=self.colors["text"], font=(_theme.FONT_FAMILY, 9, "bold"),
                           tags=("paper_label",))

    def _paper_size_for(self, index: int) -> Tuple[float, float]:
        """取该输出面的纸张尺寸（点）；缺失时退回当前版式计算的纸张。"""
        scalable = getattr(self, "preview_paper_pts", None) or []
        if 0 <= index < len(scalable):
            return scalable[index]
        metrics = self.preview_metrics_data or {}
        paper = metrics.get("paper")
        if paper:
            return float(paper[0]), float(paper[1])
        return 595.0, 842.0


    # -------------------------------------------------------------- 快捷键
    def _bind_shortcuts(self) -> None:
        """Ctrl+O 打开文件、Delete 移除、F5 刷新打印机与文件变更、Ctrl+Enter 开始打印。"""
        self.root.bind_all("<Control-o>", lambda _e: self.add_files())
        self.root.bind_all("<Control-O>", lambda _e: self.add_files())
        self.root.bind_all("<Delete>", self._on_delete_key)
        self.root.bind_all("<F5>", lambda _e: self._refresh_all())
        self.root.bind_all("<Control-Return>", lambda _e: self.start_print())

    def _refresh_all(self) -> None:
        """F5：刷新打印机列表 + 重新校验文件页数（外部修改后自动同步）。"""
        self._refresh_printers()
        self._refresh_file_metadata(notify=True)
        self._update_preview()

    def _on_delete_key(self, _event=None) -> str:
        # 只有焦点在文件列表时才执行删除，避免在输入框里按 Delete 误删文件
        try:
            focus = self.root.focus_get()
        except KeyError:
            focus = None
        if focus is self.tree:
            self.remove_selected()
            return "break"
        return ""

    # ------------------------------------------------------------------ 拖拽
    def _setup_drag_drop(self) -> None:
        """注册拖拽：优先 tkinterdnd2，缺失时回退 Win32 WM_DROPFILES。"""
        self.drop_supported = False
        self._drop_hint_text = "可将 PDF 文件直接拖入本窗口（支持多文件）"
        self._drop_hint_color = self.colors["muted"]
        if _tkdnd is not None and hasattr(self.root, "drop_target_register"):
            try:
                for widget in (self.root, self.tree):
                    widget.drop_target_register(_tkdnd.DND_FILES)
                    widget.dnd_bind("<<Drop>>", self._on_drop)
                    widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                    widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)
                self.drop_supported = True
                return
            except tk.TclError as exc:
                self._append_log(f"tkinterdnd2 拖拽注册失败，改用系统拖拽：{exc}")
        self._setup_dropfiles_win32()

    def _on_drop(self, event):
        paths = split_drop_data(getattr(event, "data", "") or "")
        self._restore_drop_hint()
        if paths:
            self._add_paths(paths)
        return getattr(event, "action", None)

    def _on_drag_enter(self, _event=None):
        try:
            self.drop_hint_label.configure(foreground=self.colors["accent"])
            self.drop_hint_var.set("松开鼠标即可添加这些 PDF 文件")
        except tk.TclError:
            pass

    def _on_drag_leave(self, _event=None):
        self._restore_drop_hint()

    def _restore_drop_hint(self) -> None:
        try:
            self.drop_hint_label.configure(foreground=self.colors["muted"])
            self.drop_hint_var.set(getattr(self, "_drop_hint_text",
                                           "可将 PDF 文件直接拖入本窗口（支持多文件）"))
        except tk.TclError:
            pass

    def _setup_dropfiles_win32(self) -> None:
        """Win32 回退：DragAcceptFiles + 子类化窗口过程处理 WM_DROPFILES。

        只在 tkinterdnd2 不可用时启用；路径通过 DragQueryFileW 读取，
        支持含空格/中文/多个文件的真实 HDROP 数据。
        """
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            shell32 = ctypes.WinDLL("shell32", use_last_error=True)
            hwnd = int(self.root.winfo_id())
            shell32.DragAcceptFiles(hwnd, True)
            user32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT,
                                              wintypes.LPWSTR, wintypes.UINT]
            user32.DragQueryFileW.restype = wintypes.UINT
            shell32.DragFinish.argtypes = [ctypes.c_void_p]
            user32.CallWindowProcW.restype = ctypes.c_long
            user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
            user32.GetWindowLongPtrW.restype = ctypes.c_void_p
            user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int,
                                                 ctypes.c_void_p]
            user32.SetWindowLongPtrW.restype = ctypes.c_void_p

            WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT,
                                         wintypes.WPARAM, wintypes.LPARAM)
            GWL_WNDPROC = -4
            WM_DROPFILES = 0x0233

            self._old_wndproc = user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)

            def handler(hwnd_value, msg, wparam, lparam):
                if msg == WM_DROPFILES:
                    hdrop = ctypes.c_void_p(wparam)
                    count = user32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
                    paths = []
                    for index in range(count):
                        length = user32.DragQueryFileW(hdrop, index, None, 0)
                        buffer = ctypes.create_unicode_buffer(length + 1)
                        user32.DragQueryFileW(hdrop, index, buffer, length + 1)
                        paths.append(buffer.value)
                    shell32.DragFinish(hdrop)
                    if paths:
                        self.root.after(0, lambda p=list(paths): self._add_paths(p))
                    return 0
                return user32.CallWindowProcW(self._old_wndproc, hwnd_value, msg, wparam, lparam)

            self._wndproc = WNDPROC(handler)
            user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC,
                                     ctypes.cast(self._wndproc, ctypes.c_void_p))
            self.drop_supported = True
        except Exception as exc:
            self._append_log(f"系统拖拽初始化失败（可改用“添加 PDF 文件”按钮）：{exc}")

    # ------------------------------------------------------- 页面设置对话框
    def open_page_setup_dialog(self) -> None:
        """独立“页面设置”对话框：纸张大小/来源/方向/页边距（对应参考截图）。"""
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()  # 先构建后显示，避免首帧闪系统底色
        dialog.title("页面设置")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(background=self.colors["bg"])

        body = ttk.Frame(dialog, padding=16, style="Card.TFrame")
        body.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(body, text="纸张", style="Subtitle.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(body, text="纸张大小：").grid(row=1, column=0, sticky="w", pady=(8, 4))
        paper_var = tk.StringVar(value=str(self.vars["paper_name"].get()))
        paper_combo = ttk.Combobox(body, textvariable=paper_var, state="readonly", width=16,
                                   values=tuple(PAPER_SIZES_MM.keys()) + (CUSTOM_PAPER_NAME,))
        paper_combo.grid(row=1, column=1, sticky="w", pady=(8, 4))

        width_var = tk.StringVar(value=str(self.vars["paper_width_mm"].get()))
        height_var = tk.StringVar(value=str(self.vars["paper_height_mm"].get()))
        ttk.Label(body, text="自定义宽×高(mm)：").grid(row=1, column=2, sticky="e", pady=(8, 4))
        size_entry = ttk.Frame(body)
        size_entry.grid(row=1, column=3, sticky="w", pady=(8, 4))
        ttk.Entry(size_entry, textvariable=width_var, width=7).pack(side="left")
        ttk.Label(size_entry, text=" × ").pack(side="left")
        ttk.Entry(size_entry, textvariable=height_var, width=7).pack(side="left")

        ttk.Label(body, text="纸张来源：").grid(row=2, column=0, sticky="w", pady=4)
        source_var = tk.StringVar(value=str(self.vars["source_name"].get()))
        source_combo = ttk.Combobox(body, textvariable=source_var, state="readonly", width=26,
                                    values=[name for name, _code in self.paper_sources])
        source_combo.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(body, text="方向：").grid(row=3, column=0, sticky="w", pady=4)
        orient_var = tk.StringVar(value=str(self.vars["orientation"].get()))
        orient_row = ttk.Frame(body)
        orient_row.grid(row=3, column=1, columnspan=3, sticky="w", pady=4)
        for value, label in (("auto", "自动"), ("portrait", "纵向"), ("landscape", "横向")):
            ttk.Radiobutton(orient_row, text=label, value=value,
                            variable=orient_var).pack(side="left", padx=(0, 12))

        ttk.Label(body, text="页边距（毫米）", style="Subtitle.TLabel").grid(
            row=4, column=0, columnspan=4, sticky="w", pady=(12, 4))
        margin_vars = {}
        margin_labels = (("margin_left_mm", "上"), ("margin_top_mm", "左"),
                         ("margin_right_mm", "右"), ("margin_bottom_mm", "下"))
        for index, (key, label) in enumerate(margin_labels):
            margin_vars[key] = tk.StringVar(value=str(self.vars[key].get()))
            ttk.Label(body, text=f"{label}：").grid(row=5 + index // 2, column=(index % 2) * 2,
                                                   sticky="w", pady=3)
            ttk.Entry(body, textvariable=margin_vars[key], width=8).grid(
                row=5 + index // 2, column=(index % 2) * 2 + 1, sticky="w", padx=(0, 16), pady=3)

        hint = ttk.Label(body, text="边距由本工具排版时留白；纸张/来源/方向通过 DEVMODE 交给驱动。",
                         style="Muted.TLabel")
        hint.grid(row=7, column=0, columnspan=4, sticky="w", pady=(10, 0))

        def apply() -> None:
            paper = paper_var.get()
            values = (width_var.get() + " " + height_var.get()).split()
            try:
                width = float(values[0])
                height = float(values[1]) if len(values) > 1 else 297.0
                if not (10 <= width <= 2000 and 10 <= height <= 2000):
                    raise ValueError
            except (ValueError, IndexError):
                messagebox.showerror("尺寸有误", "自定义纸张宽/高需为 10-2000 毫米的数字。",
                                     parent=dialog)
                return
            for key, var in margin_vars.items():
                try:
                    number = float(var.get())
                    if not 0 <= number <= 100:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("边距有误", "页边距需为 0-100 毫米的数字。", parent=dialog)
                    return
            self.vars["paper_name"].set(paper)
            self.vars["paper_width_mm"].set(_fmt_number(width))
            self.vars["paper_height_mm"].set(_fmt_number(height))
            self.vars["source_name"].set(source_var.get())
            self.vars["orientation"].set(orient_var.get())
            for key, var in margin_vars.items():
                self.vars[key].set(var.get().strip())
            dialog.destroy()
            self._sync_widgets()
            self._save_settings_now()

        buttons = ttk.Frame(body)
        buttons.grid(row=8, column=0, columnspan=4, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="确定", style="Accent.TButton", command=apply).pack(side="left")
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=(8, 0))

        dialog.update_idletasks()
        self._place_toplevel_over_root(dialog, dialog.winfo_reqwidth(),
                                       dialog.winfo_reqheight())
        dialog.deiconify()
        dialog.lift()
        dialog.grab_set()

    # ------------------------------------------------------------------ 日志
    def _append_log(self, text: str) -> None:
        if not text:
            return
        # 同步落盘（handler 未安装时为标准库空操作），便于打印失败后排查
        _logger.info(text.rstrip())
        stamp = time.strftime("%H:%M:%S")
        try:
            self.log_text.configure(state="normal")
            # 不在末尾留空行：否则“最新一条”会被空白行顶出可视区
            prefix = "" if self.log_text.index("end-1c") == "1.0" else "\n"
            line_tags = ("log_error",) if any(
                word in text for word in ("失败", "错误", "无法", "中断", "异常",
                                          "未检测到", "不可用")) else ()
            self.log_text.insert("end", f"{prefix}[{stamp}] {text.rstrip()}", line_tags)
            # 限制日志长度，防止长时间运行内存膨胀
            if int(self.log_text.index("end-1c").split(".")[0]) > 5000:
                self.log_text.delete("1.0", "2000.0")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
            # 收起状态下只显示最新一条摘要（固定一行，不占纵向空间）
            summary = text.strip().splitlines()[0] if text.strip() else ""
            if len(summary) > 68:
                summary = summary[:68] + "…"
            self._log_summary_full = f"[{stamp}] {summary}"
            self._render_log_summary()
        except tk.TclError:
            pass

    def _render_log_summary(self) -> None:
        """按摘要标签当前像素宽度做省略，保证任何缩放档位都不裁切文字。"""
        label = getattr(self, "log_summary_label", None)
        full = getattr(self, "_log_summary_full", "")
        if label is None or not full:
            return
        try:
            width = label.winfo_width()
            if width <= 1:
                width = self._px(420)
            font = tkfont.Font(root=self.root, font=label.cget("font"))
            # 标签自身的水平留白/边框不参与文字绘制，必须从可用宽度里扣除，
            # 否则省略后仍可能比控件宽几个像素（被判为裁切）。
            try:
                overhead = 2 * (int(label.cget("padx")) + int(label.cget("borderwidth"))
                                + int(label.cget("highlightthickness")))
            except (tk.TclError, ValueError):
                overhead = 4
            budget_width = max(1, width - overhead)
            text = full
            if font.measure(text) > budget_width:
                ellipsis = "…"
                budget = budget_width - font.measure(ellipsis)
                while text and font.measure(text) > budget:
                    text = text[:-1]
                text = text.rstrip() + ellipsis
            self.log_summary_var.set(text)
        except tk.TclError:
            pass

    def _log_worker(self, text: str) -> None:
        """后台线程写日志：通过队列回到主线程。"""
        self._post(("log", text))

    def export_log(self) -> None:
        text = self.log_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo("导出日志", "当前没有日志内容。", parent=self.root)
            return
        default_name = time.strftime("PDF打印日志_%Y%m%d_%H%M%S.txt")
        path = filedialog.asksaveasfilename(
            title="导出日志", defaultextension=".txt", initialfile=default_name,
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")], parent=self.root)
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(text)
            self.status_var.set(f"日志已导出：{path}")
        except OSError as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------- UI 队列
    def _post(self, item: tuple) -> None:
        self.ui_queue.put(item)

    def _ui_call(self, fn):
        """在后台线程中同步调用主线程函数（用于翻面提示、确认框）。"""
        box = {}
        done = threading.Event()

        def runner():
            try:
                box["result"] = fn()
            except Exception as exc:  # 主线程异常传回工作线程
                box["error"] = exc
            finally:
                done.set()

        self.ui_queue.put(("call", runner))
        while not done.wait(0.2):
            if self.stop_event.is_set():
                return None
        if "error" in box:
            raise box["error"]
        return box.get("result")

    def _poll_ui_queue(self) -> None:
        # 退出中：不再处理队列，也不再重新挂 after（避免窗口销毁后 Tcl 报残留回调）
        if self.shutting_down:
            return
        try:
            while True:
                item = self.ui_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._append_log(item[1])
                elif kind == "status":
                    self.status_var.set(item[1])
                elif kind == "current":
                    self.current_page_var.set(item[1])
                elif kind == "progress":
                    _index, total = item[1], item[2]
                    self.progress.configure(maximum=max(1, total), value=_index)
                elif kind == "rows":
                    statuses = item[1]
                    for iid, status in statuses:
                        if self.tree.exists(iid):
                            values = list(self.tree.item(iid, "values"))
                            while len(values) < 5:
                                values.append("")
                            values[4] = self._status_display(status)
                            self.tree.item(iid, values=values)
                        self._update_phase_from_status(status)
                    self._apply_row_tags()
                    self._update_flow_steps()
                elif kind == "busy":
                    self._set_busy(bool(item[1]))
                elif kind == "call":
                    try:
                        item[1]()
                    except Exception:
                        self._append_log("界面回调出错：\n" + traceback.format_exc())
                elif kind == "preview":
                    if item[1] == self.preview_generation:
                        self._on_preview_data(item[2])
                elif kind == "preview_error":
                    if item[1] == self.preview_generation:
                        self._clear_preview_images()
                        self.preview_info_var.set(item[2])
                elif kind == "word_done":
                    iid, pdf_path, pages, converter = item[1], item[2], item[3], item[4]
                    extra = item[5] if len(item) > 5 and isinstance(item[5], dict) else {}
                    record = self._find_record(iid)
                    if record is not None:
                        record["pdf_path"] = pdf_path
                        record["pages"] = pages
                        kind = record.get("kind")
                        record["status"] = {
                            "image": "图片（已转换为 PDF）",
                            "wps": "WPS（已转换为 PDF）",
                        }.get(kind, "Word（已转换为 PDF）")
                        self._update_row(record)
                        unit = "图片" if kind == "image" else "文档"
                        name = os.path.basename(record["path"])
                        self._append_log(
                            f"已用 {converter} 把{unit}转换为 PDF："
                            f"{name}（{pages} 页）")
                        # 字体风险只在确实存在缺失/未嵌入时提示，避免日志噪音。
                        warning = font_warning_text(
                            name, extra.get("missing_fonts") or [])
                        if warning:
                            self._append_log(warning)
                        pdf_warning = pdf_font_warning_text(
                            name, extra.get("unembedded_fonts") or [])
                        if pdf_warning:
                            self._append_log(pdf_warning)
                    self._update_preview()
                    pending = [f for f in self.files
                               if f.get("kind") in ("word", "wps", "image")
                               and not f.get("pdf_path")
                               and f.get("status") != "转换失败"]
                    if not pending and not self.print_active:
                        self.status_var.set("文件已准备就绪，可以开始打印。")
                elif kind == "word_failed":
                    record = self._find_record(item[1])
                    if record is not None:
                        record["status"] = "转换失败"
                        self._update_row(record)
                        self._append_log(f"{os.path.basename(record['path'])}：{item[2]}")
                        if not self.print_active:
                            self.status_var.set(
                                f"转换失败：{os.path.basename(record['path'])}，"
                                "请查看列表中的红色行。")
                    self._update_preview()
                elif kind == "pump_word":
                    self._start_word_conversions()
                elif kind == "open_dir":
                    try:
                        open_directory(item[1])
                    except PrintError as exc:
                        self._append_log(str(exc))
        except queue.Empty:
            pass
        if not self.shutting_down and self.root.winfo_exists():
            self.ui_poll_id = self.root.after(100, self._poll_ui_queue)

    def _set_busy(self, busy: bool) -> None:
        self.print_active = busy
        state = "disabled" if busy else "normal"
        for button in (self.add_file_btn, self.add_dir_btn, self.remove_btn, self.clear_btn,
                       self.up_btn, self.down_btn, self.dedupe_btn, self.explore_btn,
                       self.refresh_btn, self.restore_btn if hasattr(self, "restore_btn") else None):
            if button is not None:
                button.configure(state=state)
        self.start_btn.configure(state="disabled" if busy else "normal")
        self.pause_btn.configure(state="normal" if busy else "disabled")
        self.stop_btn.configure(state="normal" if busy else "disabled")
        if not busy:
            self.pause_btn.configure(text="暂停")
            self.paused = False
            self.pause_event.clear()
            self.progress.configure(value=0)
        self._update_flow_steps()
        self._update_start_state()

    def _on_close(self) -> None:
        self._hide_tree_tooltip()
        if self.worker and self.worker.is_alive():
            if not messagebox.askokcancel(
                    "正在打印", "打印任务仍在进行，确定要取消并退出吗？", parent=self.root):
                return
            self.stop_event.set()
            self.pause_event.clear()
            if self.current_job_id:
                cancel_job(self.current_printer(), self.current_job_id)
            self.worker.join(timeout=3.0)
            if self.worker.is_alive():
                # daemon 线程会在进程退出时被强制结束，但留下记录便于排查卡死原因
                _logger.warning("退出时打印线程未能在 3 秒内结束，将由 daemon 兜底终止")
        self.shutting_down = True
        if self.preview_after_id is not None:
            try:
                self.root.after_cancel(self.preview_after_id)
            except tk.TclError:
                pass
        # 取消队列与打印机状态轮询，保证销毁窗口后不再有残留 after 回调
        for attr in ("ui_poll_id", "printer_poll_id"):
            poll_id = getattr(self, attr, None)
            if poll_id is not None:
                try:
                    self.root.after_cancel(poll_id)
                except tk.TclError:
                    pass
                setattr(self, attr, None)
        self.preview_enabled = False
        if self.preview_temp_dir:
            shutil.rmtree(self.preview_temp_dir, ignore_errors=True)
        # Word 临时 PDF 目录：退出时无条件清理（含转换进行中的半成品目录）
        for temp_dir in list(self.word_temp_dirs.values()):
            shutil.rmtree(temp_dir, ignore_errors=True)
        self.word_temp_dirs.clear()
        self._remember_window_geometry()
        self._save_settings_now()
        self.root.destroy()

    # ---------------------------------------------------------- 打印执行
    def start_print(self) -> None:
        """校验所有输入并启动后台打印线程（界面不冻结）。"""
        if self.print_active:
            return
        printer = self.current_printer()
        if not printer:
            messagebox.showerror("未选择打印机",
                                 "请先在“打印设置”中选择目标打印机（可点刷新）。",
                                 parent=self.root)
            return

        # Word/WPS/图片必须先转成 PDF 才能确定页数与整批翻面顺序
        pending_word = [f for f in self.files
                        if f.get("kind") in ("word", "wps", "image")
                        and not f.get("pdf_path")]
        if pending_word:
            failed = [f for f in pending_word if f.get("status") == "转换失败"]
            self._start_word_conversions()
            if failed:
                names = "、".join(os.path.basename(f["path"]) for f in failed[:5])
                messagebox.showerror(
                    "文件转换失败",
                    f"以下文件未能转换为 PDF：{names}\n\n"
                    "Word/WPS 请安装/修复 Microsoft Word 或 WPS，或手动另存为 PDF "
                    "后重新导入；图片请确认文件未损坏。",
                    parent=self.root)
            else:
                messagebox.showinfo(
                    "文件正在转换",
                    f"还有 {len(pending_word)} 个文件正在转换为 PDF，"
                    "请等状态变为“已转换”后再开始打印。",
                    parent=self.root)
            return

        try:
            copies = int(float(str(self.vars["copies"].get()).strip()))
            if not 1 <= copies <= 99:
                raise ValueError
        except ValueError:
            messagebox.showerror("份数有误", "打印份数需为 1-99 的整数。", parent=self.root)
            return

        delay = self._float_setting("delay_seconds", 15.0)
        if not 0 <= delay <= 600:
            messagebox.showerror("延时设置有误", "固定延时需为 0-600 秒。", parent=self.root)
            return

        percent = self._int_setting("scale_percent", 100)
        if not 10 <= percent <= 500:
            messagebox.showerror("缩放比例有误", "自定义缩放比例需为 10%-500%。", parent=self.root)
            return
        for key, label in (("margin_left_mm", "左"), ("margin_top_mm", "上"),
                           ("margin_right_mm", "右"), ("margin_bottom_mm", "下"),
                           ("crop_left_mm", "裁切左"), ("crop_top_mm", "裁切上"),
                           ("crop_right_mm", "裁切右"), ("crop_bottom_mm", "裁切下")):
            value = self._float_setting(key, 0.0)
            if not 0 <= value <= 100:
                messagebox.showerror("边距/裁切有误",
                                     f"{label}边距/裁切需为 0-100 毫米。", parent=self.root)
                return
        if str(self.vars["paper_name"].get()) == CUSTOM_PAPER_NAME:
            width = self._float_setting("paper_width_mm", 0.0)
            height = self._float_setting("paper_height_mm", 0.0)
            if not (10 <= width <= 2000 and 10 <= height <= 2000):
                messagebox.showerror("自定义纸张有误",
                                     "自定义纸张宽/高需为 10-2000 毫米。", parent=self.root)
                return

        keep_output = bool(self.vars["keep_output"].get())
        output_dir = str(self.vars["output_dir"].get()).strip()
        if keep_output and not output_dir:
            messagebox.showerror("输出目录未设置",
                                 "已勾选“保留排版 PDF”，请先选择输出目录。", parent=self.root)
            return
        if keep_output:
            try:
                os.makedirs(output_dir, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("输出目录不可用", str(exc), parent=self.root)
                return

        # 在主线程完成所有 Tk 变量读取与范围解析，工作线程只处理纯数据
        jobs = []
        range_errors = []
        skipped = []
        for record in self.files:
            if not record.get("pages"):
                skipped.append(os.path.basename(record.get("path") or "?"))
                continue
            try:
                valid, dropped, file_copies = self._resolve_record(record)
            except PdfError as exc:
                range_errors.append(f"《{os.path.basename(record['path'])}》："
                                    f"{str(exc).splitlines()[0]}")
                continue
            if not valid:
                self._append_log(
                    f"《{os.path.basename(record['path'])}》的打印范围全部超出页数，已跳过。")
                continue
            jobs.append({
                "path": self._record_pdf_path(record),
                "name": os.path.basename(record["path"]),
                "pages": record["pages"],
                "valid": valid,
                "dropped": len(dropped),
                "copies": file_copies,
                "iid": record["iid"],
            })
        if range_errors:
            messagebox.showerror("页码范围有误", "\n".join(range_errors[:6]), parent=self.root)
            return
        if skipped:
            names = "、".join(skipped[:5]) + (" …" if len(skipped) > 5 else "")
            self._append_log(
                f"已跳过 {len(skipped)} 个没有可用页数的文件（未转换或无法读取）：{names}")
        if not jobs:
            messagebox.showerror("没有可打印内容",
                                 "列表中没有可用于打印的页面，请检查文件与打印范围。",
                                 parent=self.root)
            return

        mode = str(self.vars["range_mode"].get())
        if mode in ("custom", "recent"):
            self.settings["last_range"] = self._effective_range_text(mode)
            self.recent_ranges = add_recent(self.recent_ranges,
                                            self.settings["last_range"])
        self._save_settings_now()

        snapshot = {
            "printer": printer,
            "jobs": jobs,
            "copies": copies,
            "collate": str(self.vars["collate"].get()) == "collate",
            "manual_duplex": bool(self.vars["manual_duplex"].get()),
            "print_flow": str(self.vars["print_flow"].get() or "batch_all"),
            "copy_exec_on": bool(self.vars["copy_exec_on"].get()),
            "reverse": bool(self.vars["reverse"].get()),
            "wait_mode": str(self.vars["wait_mode"].get() or "queue"),
            "delay": delay,
            "keep_output": keep_output,
            "auto_open": bool(self.vars["auto_open"].get()),
            "output_dir": output_dir,
            "layout": self._layout_options(),
            "printer_options": self._printer_options(self._paper_is_landscape()),
        }

        self.stop_event.clear()
        self.pause_event.clear()
        self.paused = False
        self.current_job_id = 0
        self._set_busy(True)
        self.status_var.set("正在准备打印…")
        self.current_page_var.set("")
        self.progress.configure(value=0, maximum=100)
        flow_names = {
            "batch_all": "整批翻面·打完全部再取纸",
            "batch_take": "整批翻面·逐文件取纸",
            "per_file": "逐文件翻面·逐文件取纸",
        }
        self._append_log(
            f"开始打印：{len(jobs)} 个文件，份数 {copies}，"
            f"{flow_names.get(snapshot['print_flow'], '整批翻面·打完全部再取纸')}"
            f"{'，按份逐份执行' if copies > 1 and snapshot['copy_exec_on'] else ''}")
        self.worker = threading.Thread(target=self._worker_main, args=(snapshot,), daemon=True)
        self.worker.start()

    def _paper_is_landscape(self) -> bool:
        orientation = str(self.vars["orientation"].get() or "auto")
        if orientation == "landscape":
            return True
        if orientation == "portrait":
            return False
        name = str(self.vars["paper_name"].get() or "A4")
        if name in PAPER_SIZES_MM:
            width, height = PAPER_SIZES_MM[name]
        else:
            width = self._float_setting("paper_width_mm", 210.0)
            height = self._float_setting("paper_height_mm", 297.0)
        return width > height

    def toggle_pause(self) -> None:
        if not self.print_active:
            return
        self.paused = not self.paused
        if self.paused:
            self.pause_event.set()
            self.pause_btn.configure(text="继续")
            self.status_var.set("已暂停：当前页打印完后暂停；翻面提示期间不会自动继续")
            if self.current_job_id:
                set_job_paused(self.current_printer(), self.current_job_id, True)
        else:
            self.pause_event.clear()
            self.pause_btn.configure(text="暂停")
            self.status_var.set("已继续打印…")
            if self.current_job_id:
                set_job_paused(self.current_printer(), self.current_job_id, False)

    def stop_print(self) -> None:
        if not self.print_active:
            return
        if not messagebox.askokcancel("取消打印",
                                       "确定要取消本次整批打印吗？\n"
                                       "已打印的纸张无法收回，临时文件会被清理。",
                                       parent=self.root):
            return
        self.stop_event.set()
        self.pause_event.clear()
        if self.current_job_id:
            cancel_job(self.current_printer(), self.current_job_id)
        self.status_var.set("正在取消…（请等待当前页结束）")
        self._append_log("用户请求取消打印。")

    # ------------------------------------------------------------- 工作线程
    def _worker_main(self, snapshot: dict) -> None:
        temp_dir = ""
        prepared: List[dict] = []
        completed = False
        cancelled = False
        error_text = ""
        try:
            check_printer_ready(snapshot["printer"])
            temp_dir = create_temp_dir(PRINT_TEMP_PREFIX)
            target_dir = snapshot["output_dir"] if snapshot["keep_output"] else temp_dir
            if snapshot["keep_output"]:
                os.makedirs(target_dir, exist_ok=True)
            stamp = time.strftime("%Y%m%d_%H%M%S")

            # 1) 每个文件独立排版（各自补页，保证文件边界干净）
            for index, job in enumerate(snapshot["jobs"], start=1):
                self._raise_if_stopped()
                self._post(("status", f"正在排版 {index}/{len(snapshot['jobs'])}：{job['name']}"))
                self._post(("rows", [(job["iid"], "排版中")]))
                safe_name = self._safe_filename(os.path.splitext(job["name"])[0])
                out_path = os.path.join(target_dir, f"{index:02d}_{safe_name}_{stamp}.pdf")
                result = compose_layout_pdf(job["path"], job["valid"], out_path,
                                            snapshot["layout"])
                prepared.append({
                    "job": job,
                    "path": out_path,
                    "sides": result.plan.output_count,
                    "plan": result.plan,
                    "paper": result.plan.paper_pt,
                })
                self._post(("log", f"《{job['name']}》{describe_plan(result.plan)}"))
                self._post(("rows", [(job["iid"], "已排版")]))

            if not prepared:
                raise PdfError("没有可打印的页面。")

            # 2) 纸张尺寸/方向一致性检查（整批流程）
            batch_flow = snapshot["print_flow"] in ("batch_all", "batch_take")
            if batch_flow and len(prepared) > 1:
                papers = {(round(item["paper"][0], 1), round(item["paper"][1], 1))
                          for item in prepared}
                if len(papers) > 1:
                    detail = "、".join(
                        f"《{item['job']['name']}》{item['paper'][0] / 72 * 25.4:.0f}×"
                        f"{item['paper'][1] / 72 * 25.4:.0f}mm" for item in prepared[:4])
                    proceed = self._ui_call(lambda: messagebox.askokcancel(
                        "纸张尺寸/方向不一致",
                        "检测到各文件排版后的纸张尺寸或方向不一致：\n" + detail +
                        "\n\n整批流程会在一次任务中请求不同纸张，部分驱动会拒绝或触发换纸盒。"
                        "\n建议取消后改用「逐文件翻面 · 逐文件取纸」流程。\n\n仍要继续整批打印吗？",
                        parent=self.root))
                    if not proceed:
                        raise CancelledByUser("用户取消：纸张不一致，建议切换逐文件流程。")
                    self._post(("log", "用户选择继续整批打印（纸张尺寸不一致）。"))

            # 3) 执行打印流程
            if not snapshot["manual_duplex"]:
                self._run_single_sided(snapshot, prepared, target_dir, stamp)
            elif batch_flow:
                self._run_batch_duplex(snapshot, prepared, target_dir, stamp)
            else:
                self._run_by_file_duplex(snapshot, prepared)
            completed = True
        except CancelledByUser as exc:
            cancelled = True
            error_text = str(exc)
            self._post(("log", f"已取消：{error_text}"))
        except (ToolError, PrinterOfflineError) as exc:
            error_text = str(exc)
            self._post(("log", f"打印失败：{error_text}"))
        except Exception as exc:
            error_text = f"发生未预期错误：{exc}"
            self._post(("log", error_text + "\n" + traceback.format_exc()))
        finally:
            for item in prepared:
                if not snapshot["keep_output"]:
                    self._remove_temp_file(item["path"])
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
            statuses = [("已取消" if cancelled else ("已完成" if completed else "失败"), item["job"]["iid"])
                        for item in prepared]
            self._post(("rows", statuses))
            # 转换产物与打印临时文件一并在完成后删除（异常/取消路径同样执行）
            self._post(("call", self._cleanup_word_temps))
            self._post(("progress", 0, 1))
            self._post(("current", ""))
            if completed:
                self._post(("status", "整批打印完成"))
            elif cancelled:
                self._post(("status", "已取消"))
            else:
                self._post(("status", "打印失败"))
            self._post(("busy", False))
            if completed and snapshot["auto_open"] and snapshot["keep_output"]:
                self._post(("open_dir", snapshot["output_dir"]))
            self._show_summary(completed, cancelled, error_text, prepared)
            self.current_job_id = 0

    def _safe_filename(self, name: str) -> str:
        cleaned = "".join(ch for ch in name if ch not in '\\/:*?"<>|')
        return (cleaned or "document")[:60]

    def _show_summary(self, completed: bool, cancelled: bool, error_text: str,
                      prepared: List[dict]) -> None:
        sheets = sum(item["sides"] for item in prepared) // 2 if prepared else 0
        if completed:
            title, message = "打印完成", (
                f"共 {len(prepared)} 个文件、约 {sheets} 张纸的打印任务已提交完成。\n"
                "请按提示完成翻面与取纸，并核对页序。")
        elif cancelled:
            title, message = "已取消", "打印已取消，剩余的临时文件已清理。"
        else:
            title, message = "打印失败", error_text or "未知错误。"
        self._ui_call(lambda: messagebox.showinfo(title, message, parent=self.root)
                      if completed else
                      (messagebox.showwarning(title, message, parent=self.root)
                       if cancelled else
                       messagebox.showerror(title, message, parent=self.root)))

    def _raise_if_stopped(self) -> None:
        if self.stop_event.is_set():
            raise CancelledByUser("用户取消。")
        while self.pause_event.is_set():
            if self.stop_event.is_set():
                raise CancelledByUser("用户取消。")
            time.sleep(0.2)

    def _sleep_interruptible(self, seconds: float) -> None:
        """可被取消/暂停打断的等待（暂停不顺延结束时间）。"""
        deadline = time.time() + max(0.0, float(seconds))
        while time.time() < deadline:
            if self.stop_event.is_set():
                raise CancelledByUser("用户取消。")
            time.sleep(min(0.2, max(0.0, deadline - time.time())))

    def _remove_temp_file(self, path: str) -> None:
        """临时文件删除重试：打印机驱动可能仍在占用文件句柄。"""
        for _attempt in range(3):
            try:
                if os.path.exists(path):
                    os.remove(path)
                return
            except OSError:
                time.sleep(0.5)

    def _run_single_sided(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
        paths = [item["path"] for item in prepared]
        merged_path = os.path.join(target_dir, f"单面合并_{stamp}.pdf")
        total, counts = merge_layout_pdfs(paths, merged_path) if len(paths) > 1 else (
            prepared[0]["sides"], [prepared[0]["sides"]])
        order = compute_single_sided_orders(total, snapshot["copies"],
                                            snapshot["collate"], snapshot["reverse"])
        self._post(("rows", [(item["job"]["iid"], "单面打印") for item in prepared]))
        self._print_phase(merged_path if len(paths) > 1 else paths[0], order,
                          "PDF 单面打印", snapshot, "单面打印 · 正序")
        self._post(("log", f"单面打印完成：{total} 个输出面 × {snapshot['copies']} 份。"))

    def _run_batch_duplex(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
        """整批翻面：所有文件合并 → 全部正面 → 翻 1 次 → 反面。

        flow=batch_all：反面一次按全局逆序打完，最后统一提示取纸；
        flow=batch_take：反面按文件从后往前打，文件之间暂停取纸，最后统一提示。
        """
        flow = str(snapshot.get("print_flow") or "batch_all")
        if flow not in ("batch_all", "batch_take"):
            flow = "batch_all"
        pause_between_files = flow == "batch_take"
        counts = [item["sides"] for item in prepared]
        total = sum(counts)
        sheets = total // 2
        merged_path = os.path.join(target_dir, f"整批合并_{stamp}.pdf")
        merged_pages, _ = merge_layout_pdfs([item["path"] for item in prepared], merged_path)
        if merged_pages != total:
            raise PdfError("合并任务页数与预期不一致，请重试。")
        ranges = merged_side_ranges(counts)

        self._post(("log", (
            f"整批打印：{len(prepared)} 个文件合并为 {total} 个输出面（{sheets} 张纸），"
            f"预计翻面 {1 if not (snapshot['copies'] > 1 and snapshot['copy_exec_on']) else snapshot['copies']} 次。")))
        for (start, end), item in zip(ranges, prepared):
            self._post(("log", f"《{item['job']['name']}》占全局第 {start}-{end} 面。"))

        copies = snapshot["copies"]
        if copies > 1 and snapshot["copy_exec_on"]:
            items = compute_merged_per_copy_orders(counts, copies, snapshot["reverse"])
            for copy_no, fronts, backs in items:
                round_prefix = f"第 {copy_no}/{copies} 份 · "
                self._post(("rows", [(item["job"]["iid"], f"第{copy_no}份·正面") for item in prepared]))
                self._post(("log", f"第 {copy_no}/{copies} 份：整批阶段一（全部正面）。"))
                self._print_phase(merged_path, fronts, f"PDF 整批打印 第{copy_no}份 正面",
                                  snapshot, f"{round_prefix}整批阶段一（全部正面）")
                self._post(("rows", [(item["job"]["iid"], f"第{copy_no}份·反面") for item in prepared]))
                if not self._ask_flip_batch(counts, fronts, backs, ranges, prepared,
                                            round_no=f"第 {copy_no}/{copies} 份"):
                    raise CancelledByUser("用户在翻面提示中选择取消。")
                self._post(("log", f"第 {copy_no}/{copies} 份：整批阶段二（翻面后全局逆序反面）。"))
                self._print_batch_backs(snapshot, prepared, ranges, backs, merged_path,
                                        prefix=round_prefix,
                                        pause_between_files=pause_between_files)
                # 每份必须取出后才能换纸打印下一份，否则下一份正面会叠在上一份反面。
                action = self._ui_call(lambda c=copy_no: self._batch_done_modal(
                    c, copies, sheets, final=(c == copies)))
                if action == "cancel":
                    raise CancelledByUser("用户在完成提示中选择取消。")
            return

        fronts, backs = compute_merged_duplex_orders(
            counts, copies, snapshot["collate"], snapshot["reverse"])
        self._post(("rows", [(item["job"]["iid"], "整批·正面") for item in prepared]))
        self._post(("log", "整批阶段一：正序打印全部正面（全局奇数面）。"))
        self._print_phase(merged_path, fronts, "PDF 整批打印 阶段一（正面）",
                          snapshot,
                          f"整批阶段一 · 全部正面 · 本批 {sheets} 张 · 已完成 0 张")
        self._raise_if_stopped()
        self._post(("rows", [(item["job"]["iid"], "等待翻面") for item in prepared]))
        if not self._ask_flip_batch(counts, fronts, backs, ranges, prepared, round_no=""):
            raise CancelledByUser("用户在翻面提示中选择取消。")
        self._post(("rows", [(item["job"]["iid"], "整批·反面") for item in prepared]))
        self._post(("log", "整批阶段二：翻面后按全局逆序打印全部反面（全局偶数面）。"))
        self._print_batch_backs(snapshot, prepared, ranges, backs, merged_path,
                                pause_between_files=pause_between_files)
        action = self._ui_call(lambda: self._batch_done_modal(1, 1, sheets, final=True))
        if action == "cancel":
            raise CancelledByUser("用户在完成提示中选择取消。")
        self._post(("log", f"整批打印完成：{sheets} 张纸 × {copies} 份。"))

    def _print_batch_backs(self, snapshot: dict, prepared: List[dict],
                           ranges: List[Tuple[int, int]], backs: List[int],
                           merged_path: str, prefix: str = "",
                           *, pause_between_files: bool = False) -> None:
        """整批阶段二：打印全部反面。

        pause_between_files=True（流程 2）：按文件从后往前分组打印，文件之间
        弹“请取出这 N 张纸”，最后一组不弹（由完成提示统一收尾）。
        False（流程 1）：全局逆序一次打完，中途不停顿。
        两种方式的页序完全一致（见页序等价性单测）。
        """
        groups: List[Tuple[dict, List[int], int]] = []
        for file_no, ((start, end), item) in enumerate(zip(ranges, prepared), start=1):
            chunk = [page for page in backs if start <= page <= end]
            if chunk:
                groups.append((item, chunk, file_no))
        if not pause_between_files or len(groups) <= 1:
            self._print_phase(merged_path, backs, f"PDF {prefix}整批打印 阶段二（反面）",
                              snapshot, f"{prefix}整批阶段二 · 全部反面（全局逆序）")
            return

        total_groups = len(groups)
        ordered = list(reversed(groups))
        done_sheets = 0
        for position, (item, chunk, file_no) in enumerate(ordered, start=1):
            self._raise_if_stopped()
            job = item["job"]
            sheets = len(chunk)
            self._post(("rows", [(job["iid"], f"整批·反面 {position}/{total_groups}")]))
            self._print_phase(
                merged_path, chunk,
                f"PDF {prefix}整批打印 阶段二（反面）·{job['name']}", snapshot,
                f"{prefix}整批反面 · 文件 {file_no}/{len(prepared)} · 第二面（反面）· "
                f"《{job['name']}》 · 本文件 {sheets} 张 · 已完成 {done_sheets} 张")
            done_sheets += sheets
            if position < total_groups:
                next_name = ordered[position][0]["job"]["name"]
                action = self._ui_call(lambda: self._file_done_modal(
                    job, file_no, len(prepared), item["sides"], next_name,
                    mode="batch", group_done=position, group_total=total_groups))
                if action == "cancel":
                    raise CancelledByUser("用户在取纸提示中选择取消。")
        self._post(("log", f"{prefix}整批反面全部完成：共 {done_sheets} 张纸。"))

    def _run_by_file_duplex(self, snapshot: dict, prepared: List[dict]) -> None:
        """逐文件翻面：每个文件各自 正面 → 翻面确认 → 反面 → 取纸确认 → 下一个。

        流程 3 固定两处确认：每个文件正面打完弹“请翻面”，反面打完弹
        “请取出这 N 张纸”；最后一个文件的反面完成后弹
        “全部完成，请取出最后一叠纸”。任一确认点取消都会整批中止
        （由上层清理临时文件），未确认前不得进入下一阶段。
        """
        total_files = len(prepared)
        done_sheets = 0
        for index, item in enumerate(prepared, start=1):
            self._raise_if_stopped()
            job = item["job"]
            sheets = item["sides"] // 2
            fronts, backs = compute_duplex_orders(
                item["sides"], snapshot["copies"], snapshot["collate"],
                snapshot["reverse"])
            self._post(("rows", [(job["iid"], "第1面")]))
            self._post(("log", f"《{job['name']}》逐文件阶段一：正面。"))
            self._post(("status", f"逐文件 {index}/{total_files} · 第一面（正面）· "
                                  f"《{job['name']}》 · 本文件 {sheets} 张 · 已完成 {done_sheets} 张"))
            self._print_phase(item["path"], fronts, f"PDF {job['name']} 正面",
                              snapshot, f"逐文件 {index}/{total_files} · 第一面（正面）· "
                                        f"《{job['name']}》 · 本文件 {sheets} 张 · 已完成 {done_sheets} 张")
            self._raise_if_stopped()
            if backs:
                self._post(("rows", [(job["iid"], "等待翻面")]))
                if not self._ask_flip_file(job, item["sides"], backs, index, total_files):
                    raise CancelledByUser("用户在翻面提示中选择取消。")
            self._post(("rows", [(job["iid"], "第2面")]))
            self._post(("log", f"《{job['name']}》逐文件阶段二：反面（逆序）。"))
            self._post(("status", f"逐文件 {index}/{total_files} · 第二面（反面）· "
                                  f"《{job['name']}》 · 本文件 {sheets} 张 · 已完成 {done_sheets} 张"))
            self._print_phase(item["path"], backs, f"PDF {job['name']} 反面",
                              snapshot, f"逐文件 {index}/{total_files} · 第二面（反面）· "
                                        f"《{job['name']}》 · 本文件 {sheets} 张 · 已完成 {done_sheets} 张")
            self._post(("rows", [(job["iid"], "已完成")]))
            done_sheets += sheets
            remaining = total_files - index
            self._post(("status", f"逐文件进度：已完成 {index}/{total_files} 个文件、"
                                  f"{done_sheets} 张，剩余 {remaining} 个"))
            next_name = (prepared[index]["job"]["name"] if remaining else "")
            action = self._ui_call(lambda: self._file_done_modal(
                job, index, total_files, item["sides"], next_name,
                mode="by_file", final=(remaining == 0)))
            if action == "cancel":
                raise CancelledByUser("用户在取纸提示中选择取消。")

    def _file_done_modal(self, job: dict, done: int, total: int, sides: int,
                         next_name: str, *, mode: str = "by_file",
                         final: bool = False,
                         group_done: Optional[int] = None,
                         group_total: Optional[int] = None) -> str:
        """取纸确认框：返回 continue / cancel。

        - mode=by_file：逐文件反面打印完，取出后打印下一个；最后一个文件
          显示“全部完成，请取出最后一叠纸”。
        - mode=batch：整批反面按文件分组打印，文件之间取出后继续打印反面。
          group_done/group_total 表示“取纸分组”的进度（batch 的反面从最后
          一个文件开始，用原始文件序号会更清楚）。
        关闭窗口或按 Esc 一律按 cancel 处理（未确认不得继续下一个阶段）。
        """
        result = {"action": "cancel"}
        sheets = sides // 2
        if mode == "batch":
            title = f"该文件已打完，请取出这 {sheets} 张纸"
            next_line = f"下一个：{next_name} 的反面（同一叠纸继续打印）"
        elif final:
            title = "全部完成，请取出最后一叠纸"
            next_line = ""
        else:
            title = f"该文件已打完，请取出这 {sheets} 张纸"
            next_line = f"下一个：{next_name}"
        try:
            dialog = tk.Toplevel(self.root)
            dialog.withdraw()
            dialog.title(title)
            dialog.configure(background=self.colors["surface"])
            dialog.resizable(False, False)
            body = ttk.Frame(dialog, padding=(16, 14))
            body.pack(fill="both", expand=True)
            ttk.Label(body, text=title, style="Subtitle.TLabel").pack(anchor="w")
            if mode == "batch" and group_done is not None and group_total is not None:
                progress_line = (f"进度：已取出 {group_done}/{group_total} 个文件的反面，"
                                 f"剩余 {max(0, group_total - group_done)} 个文件的反面")
            else:
                progress_line = (f"进度：已完成 {done}/{total} 个文件，"
                                 f"剩余 {max(0, total - done)} 个")
            info = [
                f"文件：《{job['name']}》（第 {done}/{total} 个文件）",
                f"阶段：第二面（反面）已完成 · 本文件 {sheets} 张纸（{sides} 个输出面）",
                progress_line,
                "示例页序：" + self._duplex_sample_text(sides),
            ]
            if next_line:
                info.append(next_line)
            for line in info:
                ttk.Label(body, text=line, justify="left", wraplength=self._px(520),
                          foreground=self.colors["text"]).pack(anchor="w", pady=(6, 0))
            ttk.Label(body, text="未确认前不会继续打印；关闭窗口或按 Esc 将取消整批。",
                      style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

            def finish(action: str) -> None:
                result["action"] = action
                try:
                    dialog.grab_release()
                except tk.TclError:
                    pass
                dialog.destroy()

            buttons = ttk.Frame(body)
            buttons.pack(fill="x", pady=(14, 0))
            if mode == "batch":
                action_text = "已取出，继续打印反面"
            elif final:
                action_text = "已取出，全部完成"
            else:
                action_text = "已取出，打印下一个文件"
            ttk.Button(buttons, text=action_text, style="Accent.TButton",
                       command=lambda: finish("continue")).pack(side="left")
            ttk.Button(buttons, text="取消打印", style="Ghost.TButton",
                       command=lambda: finish("cancel")).pack(side="left", padx=(10, 0))
            dialog.protocol("WM_DELETE_WINDOW", lambda: finish("cancel"))
            dialog.bind("<Escape>", lambda _e: finish("cancel"))
            dialog.update_idletasks()
            self._place_toplevel_over_root(
                dialog, dialog.winfo_reqwidth() + self._px(8),
                dialog.winfo_reqheight() + self._px(8))
            dialog.deiconify()
            if getattr(self, "offscreen_dialog_alpha", None) is not None:
                dialog.attributes("-alpha", float(self.offscreen_dialog_alpha))
            dialog.lift()
            dialog.grab_set()
            dialog.focus_force()
            dialog.wait_window()
        except tk.TclError:
            return "cancel"
        return result["action"]

    def _batch_done_modal(self, done: int, total: int, sheets: int,
                          *, final: bool) -> str:
        """整批流程全部打印完成后的取纸提示（每份/整批）。

        返回 "continue"；Esc、关闭窗口或“取消打印”统一返回 "cancel"。
        """
        result = {"action": "cancel"}
        title = ("全部完成，请取出全部纸张" if final else
                 f"第 {done}/{total} 份已打完，请取出全部纸张")
        try:
            dialog = tk.Toplevel(self.root)
            dialog.withdraw()
            dialog.title(title)
            dialog.configure(background=self.colors["surface"])
            dialog.resizable(False, False)
            body = ttk.Frame(dialog, padding=(16, 14))
            body.pack(fill="both", expand=True)
            ttk.Label(body, text=title, style="Subtitle.TLabel").pack(anchor="w")
            info = [
                "该批所有文件的反面已打印完成。",
                f"张数：共 {sheets} 张纸（单面朝下时按打印机出纸方向整理）。",
                ("流程：整批翻面 · 打完全部再取纸（中途不停顿）"
                 if not final or done == total else
                 f"流程：整批翻面 · 逐文件取纸（本份已完成 {done}/{total}）"),
            ]
            for line in info:
                ttk.Label(body, text=line, justify="left", wraplength=self._px(520),
                          foreground=self.colors["text"]).pack(anchor="w", pady=(6, 0))
            ttk.Label(body, text="未确认前不会继续打印；关闭窗口或按 Esc 将取消整批。",
                      style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

            def finish(action: str) -> None:
                result["action"] = action
                try:
                    dialog.grab_release()
                except tk.TclError:
                    pass
                dialog.destroy()

            buttons = ttk.Frame(body)
            buttons.pack(fill="x", pady=(14, 0))
            ttk.Button(buttons, text="已取出，全部完成", style="Accent.TButton",
                       command=lambda: finish("continue")).pack(side="left")
            ttk.Button(buttons, text="取消打印", style="Ghost.TButton",
                       command=lambda: finish("cancel")).pack(side="left", padx=(10, 0))
            dialog.protocol("WM_DELETE_WINDOW", lambda: finish("cancel"))
            dialog.bind("<Escape>", lambda _e: finish("cancel"))
            dialog.update_idletasks()
            self._place_toplevel_over_root(
                dialog, dialog.winfo_reqwidth() + self._px(8),
                dialog.winfo_reqheight() + self._px(8))
            dialog.deiconify()
            if getattr(self, "offscreen_dialog_alpha", None) is not None:
                dialog.attributes("-alpha", float(self.offscreen_dialog_alpha))
            dialog.lift()
            dialog.grab_set()
            dialog.focus_force()
            dialog.wait_window()
        except tk.TclError:
            return "cancel"
        return result["action"]

    @staticmethod
    def _duplex_sample_text(sides: int) -> str:
        """示例页序：正面 1,3,5… → 反面（本文件范围内逆序）。"""
        fronts = list(range(1, sides + 1, 2))
        backs = list(range(2, sides + 1, 2))[::-1]
        return (f"正面 {format_page_list(fronts, limit=8)} → "
                f"反面 {format_page_list(backs, limit=8)}（共 {sides // 2} 张纸）")

    def _print_phase(self, pdf_path: str, order: List[int], job_name: str,
                     snapshot: dict, phase_text: str) -> None:
        """提交一个打印任务并按设置等待完成（队列监听超时自动降级）。"""
        if not order:
            return
        self._raise_if_stopped()
        self._post(("status", phase_text))
        self._post(("log", f"{phase_text}：共 {len(order)} 个输出面。"))
        started_at = time.time()

        def on_page(index: int, total: int, rel: int) -> None:
            self._post(("progress", index, total))
            self._post(("current", f"{phase_text} · {index}/{total} 面（PDF 第 {rel} 页）"))

        printer = GdiPrinter(snapshot["printer"], snapshot["printer_options"])
        try:
            job_id = printer.print_pages(
                pdf_path, order, job_name, self.stop_event,
                on_page=on_page, pause_event=self.pause_event)
        finally:
            self.current_job_id = printer.last_job_id
        if job_id <= 0:
            raise PrintError("打印任务提交失败。")

        if snapshot["wait_mode"] == "queue":
            try:
                finished = wait_for_print_job(
                    snapshot["printer"], job_id, self.stop_event,
                    log=self._log_worker, pause_event=self.pause_event)
                if not finished:
                    raise CancelledByUser("用户取消。")
            except PrintTimeoutError as exc:
                # 超时兜底：改用固定延时继续，避免整批中断
                self._post(("log", str(exc)))
                self._post(("status", f"{phase_text}：队列监听超时，改用固定延时继续…"))
                self._sleep_interruptible(snapshot["delay"])
        else:
            self._sleep_interruptible(snapshot["delay"])
        # 给机械出纸留少量余量
        self._sleep_interruptible(1.0)
        # 每个打印任务一行结果日志（jobId / 面数 / 耗时），便于失败后回看定位
        self._post(("log", f"{phase_text}：任务 #{job_id} 结束，{len(order)} 个输出面，"
                           f"用时 {time.time() - started_at:.1f} 秒。"))

    def _flip_modal(self, title: str, detail_lines: List[str], sample: str) -> bool:
        """翻面确认模态框：只有“已翻面，继续打印反面”返回 True。

        强约束：
          * 没有倒计时、没有 after 定时器、没有“按任意键继续”；
          * 关闭窗口 / Esc / Alt+F4 一律视为取消（result=False）；
          * 取消后由调用方抛 CancelledByUser，删除临时 PDF 并停止后续阶段。
        """
        result = {"ok": False}
        try:
            dialog = tk.Toplevel(self.root)
            dialog.withdraw()
            dialog.title(title)
            dialog.configure(background=self.colors["surface"])
            dialog.resizable(False, False)
            body = ttk.Frame(dialog, padding=(16, 14))
            body.pack(fill="both", expand=True)
            ttk.Label(body, text=title, style="Subtitle.TLabel").pack(anchor="w")
            for line in detail_lines:
                ttk.Label(body, text=line, justify="left", wraplength=self._px(540),
                          foreground=self.colors["text"]).pack(anchor="w", pady=(6, 0))
            ttk.Label(body, text=sample, font=_theme.FONT_SMALL,
                      foreground=self.colors["accent"], justify="left",
                      wraplength=self._px(540)).pack(anchor="w", pady=(10, 0))
            ttk.Label(body, text="关闭窗口或按 Esc 等同于取消；本工具不会自动继续。",
                      style="Muted.TLabel").pack(anchor="w", pady=(10, 0))

            def finish(ok: bool) -> None:
                result["ok"] = bool(ok)
                try:
                    dialog.grab_release()
                except tk.TclError:
                    pass
                dialog.destroy()

            buttons = ttk.Frame(body)
            buttons.pack(fill="x", pady=(14, 0))
            confirm = ttk.Button(buttons, text="已翻面，继续打印反面",
                                 style="Accent.TButton", command=lambda: finish(True))
            confirm.pack(side="left")
            ttk.Button(buttons, text="取消打印", style="Ghost.TButton",
                       command=lambda: finish(False)).pack(side="left", padx=(10, 0))
            dialog.protocol("WM_DELETE_WINDOW", lambda: finish(False))
            dialog.bind("<Escape>", lambda _e: finish(False))
            dialog.update_idletasks()
            self._place_toplevel_over_root(
                dialog, dialog.winfo_reqwidth() + self._px(8),
                dialog.winfo_reqheight() + self._px(8))
            dialog.deiconify()
            if getattr(self, "offscreen_dialog_alpha", None) is not None:
                dialog.attributes("-alpha", float(self.offscreen_dialog_alpha))
            dialog.lift()
            dialog.grab_set()
            dialog.focus_force()
            confirm.focus_set()
            dialog.wait_window()
        except tk.TclError:
            return False
        return result["ok"]

    def _ask_flip_batch(self, counts: List[int], fronts: List[int], backs: List[int],
                        ranges: List[Tuple[int, int]], prepared: List[dict],
                        round_no: str = "") -> bool:
        sheets = sum(counts) // 2
        groups = []
        for (start, end), item in zip(ranges, prepared):
            pages = [p for p in fronts if start <= p <= end]
            groups.append(f"《{item['job']['name']}》：{format_page_list(pages, limit=8)}")
        prefix = f"{round_no} " if round_no else ""
        detail = [
            f"{prefix}阶段：整批阶段一（全部正面）已完成",
            f"文件：{'、'.join('《%s》' % item['job']['name'] for item in prepared[:6])}"
            + (" 等" if len(prepared) > 6 else ""),
            f"张数：{sheets} 张纸（{sum(counts)} 个输出面）",
            "翻面方向：整叠纸【整体翻面】，保持上下顺序不变（通常为长边翻转）",
            "正面页序：\n" + "\n".join(groups),
            f"反面页序：{format_page_list(backs, limit=10)}（整批全局逆序）",
        ]
        sample = "示例：正面 1,3,5 … → 整叠翻面 → 反面 6,4,2 …（每文件已补空白页，不跨文件错位）"
        return bool(self._ui_call(lambda: self._flip_modal(
            "请翻面后继续（整批只翻 1 次）", detail, sample)))

    def _ask_flip_file(self, job: dict, sides: int, backs: List[int],
                       index: int, total_files: int) -> bool:
        sheets = sides // 2
        detail = [
            f"阶段：逐文件翻面 · 第一面（正面）已完成",
            f"文件：《{job['name']}》（第 {index}/{total_files} 个，剩余 {total_files - index} 个）",
            f"张数：{sheets} 张纸（{sides} 个输出面）",
            "翻面方向：整叠纸【整体翻面】，保持上下顺序不变（通常为长边翻转）",
            f"反面页序：{format_page_list(backs, limit=10)}（本文件范围内逆序）",
        ]
        sample = "示例页序：" + self._duplex_sample_text(sides)
        return bool(self._ui_call(lambda: self._flip_modal(
            "请翻面后继续（逐文件模式）", detail, sample)))




