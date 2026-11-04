# -*- coding: utf-8 -*-
"""设置面板/高级设置/设置读写/打印机/窗口几何（自 ui_main.py 机械拆分）。"""


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


class SettingsPanelMixin:
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
