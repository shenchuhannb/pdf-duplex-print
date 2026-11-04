# -*- coding: utf-8 -*-
"""成纸预览（自 ui_main.py 机械拆分）。"""


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


class PreviewMixin:
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
