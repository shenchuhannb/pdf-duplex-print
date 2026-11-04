# -*- coding: utf-8 -*-
"""打印执行、队列轮询、流程弹窗（自 ui_main.py 机械拆分）。"""


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


class WorkerMixin:
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
