# -*- coding: utf-8 -*-
"""日志抽屉与日志操作（自 ui_main.py 机械拆分）。"""


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
import win32_flicker
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


class LogPanelMixin:
    # ------------------------------------------------------------------ 底部常驻日志
    def _build_log_drawer(self, parent: tk.Frame) -> None:
        """底部常驻日志区（默认展开、始终可见）。

        结构：可拖动分隔条 → 标题条（日志 + 导出/清空/打开目录 + 页序详情 +
        最新一条摘要）→ 日志正文。高度按像素固定（默认 ≈6 行），拖动分隔条
        在最小值与窗口 1/3 之间调整；窗口变高时自动按比例增高，但不挤压
        工作区（工作区所在行仍是唯一 weight=1 的伸缩行）。

        拖动分隔条时只移动一个居中的浮动小手柄，松手后才应用高度：Windows 下内部
        拖动只产生局部脏区，Tk 来不及逐帧完整重绘，文件列表空状态和日志
        标题会残留旧像素、日志框会追赶鼠标抖动；窗口边框缩放是系统整窗
        失效所以正常。这里采用 Tk 经典 PanedWindow 的 opaqueresize=0 策略。
        """
        c = self.colors
        self._log_min_height = self._px(134)
        self._log_height = self._log_min_height
        self._log_user_sized = False
        self._log_drag = None
        self._log_drag_active = False
        self._log_pending_height = None
        self._log_settle_after = None
        self.log_drawer_open = True

        sash = tk.Frame(parent, background=c["bg"], height=self._px(6),
                        cursor="sb_v_double_arrow", highlightthickness=0)
        sash.grid(row=4, column=0, sticky="ew", pady=(self._px(GAP_S), 0))
        sash.bind("<Button-1>", self._on_log_sash_press)
        sash.bind("<B1-Motion>", self._on_log_sash_drag)
        sash.bind("<ButtonRelease-1>", self._on_log_sash_release)
        self.log_sash = sash
        grip = tk.Frame(sash, background=c["border_strong"], width=self._px(42),
                        height=self._px(2))
        grip.place(relx=0.5, rely=0.5, anchor="center")
        for widget in (sash, grip):
            widget.bind("<Button-1>", self._on_log_sash_press)
            widget.bind("<B1-Motion>", self._on_log_sash_drag)
            widget.bind("<ButtonRelease-1>", self._on_log_sash_release)

        drawer = tk.Frame(parent, background=c["surface"], highlightthickness=1,
                          highlightbackground=c["border"], height=self._log_height)
        drawer.grid(row=5, column=0, sticky="ew", pady=(self._px(GAP_S), 0))
        drawer.grid_propagate(False)
        drawer.columnconfigure(0, weight=1)
        drawer.rowconfigure(1, weight=1)
        self.log_drawer = drawer

        head = tk.Frame(drawer, background=c["surface_alt"],
                        padx=self._px(10), pady=self._px(4))
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(5, weight=1)
        tk.Label(head, text="运行日志", background=c["surface_alt"], foreground=c["text"],
                 font=_theme.FONT_SUBTITLE).grid(row=0, column=0, sticky="w")
        self.export_log_btn = ttk.Button(head, text="导出", style="ToolbarGhost.TButton",
                                         command=self.export_log)
        self.export_log_btn.grid(row=0, column=1, padx=(self._px(GAP), 0))
        ttk.Button(head, text="清空", style="ToolbarGhost.TButton",
                   command=self.clear_log).grid(row=0, column=2, padx=(self._px(GAP_S), 0))
        ttk.Button(head, text="打开目录", style="ToolbarGhost.TButton",
                   command=self.open_log_dir).grid(row=0, column=3, padx=(self._px(GAP_S), 0))
        self.log_detail_var = tk.StringVar(value="页序/缩略图 ▸")
        ttk.Button(head, textvariable=self.log_detail_var, style="ToolbarGhost.TButton",
                   command=self._toggle_log_detail).grid(row=0, column=4, padx=(self._px(GAP_S), 0))
        # 最新一条始终可见（按实际可用像素省略，不裁切半个字）
        self.log_summary_var = tk.StringVar(value="暂无日志")
        self._log_summary_full = "暂无日志"
        summary = tk.Label(head, textvariable=self.log_summary_var, background=c["surface_alt"],
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

        # 拖动只移动这个小手柄，任何布局都不动，从根上避免拖动过程中的残影。
        # 白底 + 强调色细描边 + 约 2px 握把，沿用分隔条的语言，不横贯整窗。
        self.log_drag_line = tk.Frame(parent, background=c["surface"],
                                      highlightthickness=1,
                                      highlightbackground=c["accent"],
                                      height=self._px(10), width=self._px(56))
        tk.Frame(self.log_drag_line, background=c["accent"],
                 width=self._px(28), height=max(2, self._px(2))).place(
            relx=0.5, rely=0.5, anchor="center")
        self.log_drag_line.place_forget()

    # ---------------------------------------------------------- 日志高度与详情
    def _on_log_sash_press(self, event) -> None:
        # 上一次松手的排版还没落定时先补完，再开始新的拖动。
        if self._log_settle_after is not None:
            self._cancel_log_settle()
            self._finish_log_resize()
        self._log_drag = (event.y_root, self._log_height,
                          self.log_sash.winfo_y() + self.log_sash.winfo_height() // 2)
        self._log_drag_active = True
        self._log_pending_height = self._log_height
        try:
            # 指示线可能远离 6px 分隔条，拖动期间保持缩放光标。
            self.root.configure(cursor="sb_v_double_arrow")
        except tk.TclError:
            pass
        self._show_log_drag_line(self._log_drag[2])

    def _on_log_sash_drag(self, event) -> None:
        if not self._log_drag:
            return
        start_y, start_height, start_center = self._log_drag
        height = self._clamp_log_height(start_height - (event.y_root - start_y))
        self._log_pending_height = height
        self._position_log_drag_line(start_center - (height - start_height))

    def _show_log_drag_line(self, center_y: int) -> None:
        self._position_log_drag_line(center_y)
        try:
            self.log_drag_line.lift()
        except tk.TclError:
            pass

    def _position_log_drag_line(self, center_y: int) -> None:
        try:
            parent = self.log_drag_line.master
            width, line_h = self._px(56), self._px(10)
            max_y = max(0, parent.winfo_height() - line_h)
            sash_w = self.log_sash.winfo_width()
            self.log_drag_line.place(
                x=self.log_sash.winfo_x() + max(0, (sash_w - width) // 2),
                y=max(0, min(int(center_y) - line_h // 2, max_y)),
                width=width, height=line_h)
        except tk.TclError:
            pass

    def _cancel_log_settle(self) -> None:
        if self._log_settle_after is not None:
            try:
                self.root.after_cancel(self._log_settle_after)
            except tk.TclError:
                pass
            self._log_settle_after = None

    def _on_log_sash_release(self, event=None) -> None:
        if not getattr(self, "_log_drag_active", False):
            return
        target = self._log_pending_height
        if target is None:
            target = self._log_height
        self._log_drag = None
        self._log_pending_height = None
        try:
            self.log_drag_line.place_forget()
            self.root.configure(cursor="")
        except tk.TclError:
            pass
        self._log_user_sized = True
        if target == self._log_height:
            # 只点了一下没有改变高度：不做任何排版与整树重绘。
            self._log_drag_active = False
            return
        # 先放开工作区内容的最小高度，再一次性落到目标高度。
        self._relax_workspace_for_log_drag()
        self._set_log_height(target)
        self._cancel_log_settle()
        self._log_settle_after = self.root.after_idle(self._finish_log_resize)

    def _finish_log_resize(self) -> None:
        self._log_settle_after = None
        if getattr(self, "shutting_down", False):
            self._log_drag_active = False
            return
        try:
            self.root.update_idletasks()
        except tk.TclError:
            self._log_drag_active = False
            return
        # 先把新高度对应的排版跑完（期间 Configure 仍被冻结跳过），
        # 再解除冻结并一次性补齐密度、预览和信息条，避免阈值反复切换。
        self._log_drag_active = False
        self._sync_file_panel_density()
        self._sync_workspace_density()
        self._sync_preview_info_visibility()
        self._on_preview_frame_configure()
        self._on_preview_canvas_configure()
        win32_flicker.invalidate_tree(self.root)

    def _relax_workspace_for_log_drag(self) -> None:
        """应用日志高度前放开工作区内容的最小高度，让收缩连续、不被卡住。"""
        try:
            if self.layout_scheme == "split" and hasattr(self, "_split_left"):
                floor = self._px(96)
                self._split_left.rowconfigure(0, minsize=floor)
                self._split_left.rowconfigure(1, minsize=floor)
            elif hasattr(self, "preview_frame"):
                self.preview_frame.rowconfigure(2, minsize=self._px(48))
            # 松手后的密度同步必须无条件恢复这些最小高度。
            self._split_density_key = None
            self._preview_canvas_min = None
        except tk.TclError:
            pass

    def _clamp_log_height(self, value: int) -> int:
        """把日志高度夹在“最小值 … 窗口 1/3”之间。"""
        try:
            window_h = max(1, self.root.winfo_height())
        except tk.TclError:
            window_h = self._log_min_height * 3
        limit = max(self._log_min_height, window_h // 3)
        return int(max(self._log_min_height, min(int(value), limit)))

    def _set_log_height(self, value: int) -> None:
        height = self._clamp_log_height(value)
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
