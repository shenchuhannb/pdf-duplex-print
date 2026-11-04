# -*- coding: utf-8 -*-
"""主界面（tkinter/ttk）—— 组合骨架。

界面分为五个页签：文件、打印设置、版面、预览、日志；底部固定执行区。
打印任务全部在后台线程执行，通过队列回调更新界面，保证不冻结。

DuplexPrintApp 由下列 mixin 组合而成（纯机械拆分，逻辑不变）：
  * ui_files.py     —— 文件面板与文件操作、拖拽、Word 转换编排
  * ui_settings.py  —— 设置面板/高级设置/设置读写/打印机/窗口几何
  * ui_preview.py   —— 成纸预览
  * ui_log.py       —— 日志抽屉与日志操作
  * ui_worker.py    —— 打印执行、队列轮询、流程弹窗
共享常量集中在 ui_helpers.py，此处 re-export 保持入口兼容。
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


from ui_helpers import (
    APP_TITLE,
    PREVIEW_MAX_THUMBS,
    PREVIEW_LONG_EDGE,
    LIVE_REFRESH_DEBOUNCE_MS,
    GAP_S,
    GAP,
    GAP_CARD,
    CARD_PAD,
    LAYOUT_SCHEMES,
    DEFAULT_LAYOUT_SCHEME,
    WINDOW_DEFAULT_SIZE,
    SCALE_LABELS,
    SCALE_LABEL_TO_VALUE,
    RANGE_LABELS,
    RANGE_LABEL_TO_VALUE,
    PAGE_ORDER_LABELS,
    COLOR_LABELS,
    QUALITY_LABELS,
    WAIT_LABELS,
)


from ui_files import FilesPanelMixin
from ui_log import LogPanelMixin
from ui_preview import PreviewMixin
from ui_settings import SettingsPanelMixin
from ui_worker import WorkerMixin


_logger = logging.getLogger("pdf_duplex")


try:
    import tkinterdnd2 as _tkdnd
except ImportError:  # 可选依赖
    _tkdnd = None


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


class DuplexPrintApp(FilesPanelMixin, SettingsPanelMixin, PreviewMixin, LogPanelMixin, WorkerMixin):
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

    def _on_status_change(self, *_args) -> None:
        """状态文字变化：按语义着色，并按可用宽度省略，避免长文案挤压进度条。"""
        self._status_full = str(self.status_var.get())
        self._render_status_text()
        self._update_flow_steps()

    _STATUS_ERROR_WORDS = ("失败", "错误", "不可用", "未选择", "请先", "无法", "有误",
                           "未检测到", "不支持", "未就绪")
    _STATUS_WARN_WORDS = ("暂停", "取消")
    _STATUS_SUCCESS_WORDS = ("完成", "已清空", "已移除", "已继续", "准备就绪", "已就绪",
                             "已恢复", "已打开")

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
