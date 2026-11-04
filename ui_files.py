# -*- coding: utf-8 -*-
"""文件面板与文件操作、拖拽、Word 转换编排（自 ui_main.py 机械拆分）。"""


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


_logger = logging.getLogger("pdf_duplex")


try:
    import tkinterdnd2 as _tkdnd
except ImportError:  # 可选依赖
    _tkdnd = None


class FilesPanelMixin:
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


# _elide_filename 体内以 DuplexPrintApp._elide_text 形式引用同类静态方法；
# 拆分后该类名在本模块内别名到自身 mixin，保持方法体逐字节不变。
DuplexPrintApp = FilesPanelMixin
