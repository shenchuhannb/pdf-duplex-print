# -*- coding: utf-8 -*-
"""js_api 桥（完整移植版）：对齐 Tk 版全部功能。

每个方法返回 {"ok": bool, ...}；耗时操作走线程池。
前端通过 window.pywebview.api.<method> 调用。
"""
import base64
import json
import os
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional, Tuple

import fitz

import printer_core
import settings_store
from core_utils import (
    CancelledByUser, PdfError, PrintError, PrinterOfflineError, PrintTimeoutError,
    PrintJobUnknownError, SettingsError, ToolError,
    format_page_list, is_pdf_path, is_word_path, is_wps_path,
    parse_page_range,
)
from pdf_layout import (
    CUSTOM_PAPER_NAME, NUP_CHOICES, PAPER_SIZES_MM,
    LayoutOptions, build_layout_plan, compose_layout_pdf, compute_duplex_orders,
    compute_merged_duplex_orders, compute_merged_per_copy_orders,
    compute_single_sided_orders, describe_plan, get_pdf_page_count,
    merge_layout_pdfs, merged_side_ranges, merged_sheet_count, plan_batch_summary,
    preview_metrics, resolve_range_pages,
)
from word_convert import (
    BatchCancelled, WordConvertError, available_converters,
    cleanup_stale_temp_dirs, convert_file, create_temp_dir,
)
from image_convert import image_to_pdf, is_image_path, supported_extensions_text
from printer_core import (
    GdiPrinter, PrinterOptions, cancel_job, check_printer_ready, get_default_printer,
    get_printable_insets_pt, get_printer_sources, get_printer_status, list_printers,
    open_directory, set_job_paused, wait_for_print_job,
)
from log_setup import get_logger


def _ok(**kw) -> Dict[str, Any]:
    return {"ok": True, **kw}


def _err(exc: Exception) -> Dict[str, Any]:
    return {"ok": False, "error": type(exc).__name__, "message": str(exc)}


class AppBridge:
    def __init__(self) -> None:
        self._window = None
        self._executor = ThreadPoolExecutor(max_workers=3)
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.current_job_id = 0
        self._printing = False
        self._paused = False
        self._flip_event = threading.Event()
        self._flip_result = False
        self._phase_confirm_event = threading.Event()
        self._phase_confirm_result = False
        # 文件表（iid 字符串 -> dict）
        self.files: Dict[str, Dict[str, Any]] = {}
        self._iid = 0
        # 当前生效设置（来自 settings.json，前端修改后同步回写）
        self.vars: Dict[str, Any] = dict(settings_store.load_settings())
        # 打印历史摘要
        self.last_print_summary: Dict[str, Any] = {}
        # 转换进度跟踪（前端轮询）
        self._converting: Dict[str, Dict[str, Any]] = {}

    # ============================================================ 基础设施
    def attach(self, window) -> None:
        self._window = window
        if os.environ.get("WV_SELFTEST") == "1":
            threading.Thread(target=self._selftest, daemon=True).start()

    def _selftest(self) -> None:
        import time
        time.sleep(4.0)
        checks = {}
        try:
            layout = self._window.evaluate_js(
                "JSON.stringify({metrics:document.querySelectorAll('.metric').length,"
                "topbar:document.querySelector('.topbar').offsetHeight,"
                "workspace:document.querySelector('.workspace').offsetHeight,"
                "printbar:document.querySelector('.printbar').offsetHeight,"
                "logpanel:document.querySelector('.logpanel').offsetHeight,"
                "vh:window.innerHeight,jsRan:!!window.__event,"
                "printers:(document.getElementById('printer-select')||{}).length,"
                "advBtn:!!document.getElementById('btn-advanced'),"
                "pauseBtn:!!document.getElementById('btn-pause'),"
                "cancelBtn:!!document.getElementById('btn-cancel'),"
                "planBtn:!!document.getElementById('btn-plan-detail'),"
                "progressBar:!!document.getElementById('progress-track')})")
            print("SELFTEST_LAYOUT:", layout, flush=True)
            checks["settings"] = self.get_settings().get("ok")
            checks["printers"] = self.list_printers().get("ok")
            checks["converters"] = self.converters_available().get("ok")
            checks["recent"] = self.get_recent().get("ok")
            checks["cleanup"] = self.cleanup_stale().get("ok")
            import tempfile as tf_mod
            test_path = os.path.join(tf_mod.mkdtemp(prefix="wv-selftest-"), "test.pdf")
            doc = fitz.open()
            for i in range(4):
                page = doc.new_page(width=595, height=842)
                page.insert_text((72, 72), f"SELFTEST_PAGE_{i + 1}")
            doc.save(test_path)
            doc.close()
            add_res = self.add_files([test_path])
            checks["add_file"] = add_res.get("ok") and len(add_res.get("added", [])) == 1
            if add_res.get("added"):
                rec = add_res["added"][0]
                checks["pages_correct"] = rec["pages"] == 4
                checks["set_options"] = self.set_file_options(rec["iid"], "1-3", 2).get("ok")
                prev = self.preview_page(rec["iid"], 1, {"copies": 1})
                checks["preview_render"] = prev.get("ok") and prev.get("total") == 4
                checks["plan_description"] = self.plan_description(rec["iid"], {"copies": 1}).get("ok")
                checks["compute_plan"] = self.compute_plan({"copies": 2, "manual_duplex": True}).get("ok")
                self.clear_files()
            os.unlink(test_path)
            print("SELFTEST_API:", json.dumps(checks), flush=True)
            print("SELFTEST_RESULT:", "PASS" if all(checks.values()) else "FAIL", flush=True)
        except Exception:
            print("SELFTEST_ERROR:\n" + traceback.format_exc(), flush=True)

    def _js(self, code: str) -> None:
        if self._window is not None:
            try:
                self._window.evaluate_js(code)
            except Exception:
                pass

    def _post(self, kind: str, payload: Any) -> None:
        data = json.dumps(payload, ensure_ascii=False)
        self._js(f"window.__event({json.dumps(kind)}, {data})")

    def _log(self, text: str, error: bool = False) -> None:
        import time
        stamp = time.strftime("%H:%M:%S")
        logger = get_logger()
        logger.info(text)
        self._post("log", {"stamp": stamp, "text": text, "error": error})

    # ============================================================ 设置
    def get_settings(self) -> Dict[str, Any]:
        return _ok(settings=self.vars, schemes={
            "nup": [{"value": n, "label": {1: "单页", 2: "2 合 1", 4: "4 合 1", 6: "6 合 1",
                                          8: "8 合 1", 9: "9 合 1"}.get(n, str(n))} for n in NUP_CHOICES],
            "range_modes": [
                {"value": "all", "label": "全部页面"},
                {"value": "custom", "label": "自定义范围"},
                {"value": "odd", "label": "仅奇数页"},
                {"value": "even", "label": "仅偶数页"},
                {"value": "reverse", "label": "逆序打印"},
            ],
            "scale_modes": [
                {"value": "shrink", "label": "适合可打印区"},
                {"value": "percent", "label": "自定义缩放"},
                {"value": "actual", "label": "实际大小 100%"},
            ],
            "orientations": [
                {"value": "auto", "label": "自动"},
                {"value": "portrait", "label": "纵向"},
                {"value": "landscape", "label": "横向"},
            ],
            "paper_names": list(PAPER_SIZES_MM.keys()) + [CUSTOM_PAPER_NAME],
            "page_orders": [{"value": "row", "label": "从左到右"},
                            {"value": "column", "label": "从上到下"}],
            "color_modes": [{"value": "color", "label": "彩色"},
                            {"value": "gray", "label": "灰度"},
                            {"value": "mono", "label": "黑白"}],
            "qualities": [{"value": "draft", "label": "草稿"},
                          {"value": "normal", "label": "标准"},
                          {"value": "high", "label": "高质量"}],
            "wait_modes": [{"value": "queue", "label": "监听打印队列（推荐）"},
                           {"value": "delay", "label": "固定延时（秒）"}],
            "flows": [{"value": "batch_all", "label": "整批，最后取纸"},
                      {"value": "batch_take", "label": "整批，逐文件取纸"},
                      {"value": "per_file", "label": "逐文件执行"}],
        })

    def save_settings(self, changes: dict) -> Dict[str, Any]:
        self.vars.update(changes)
        try:
            settings_store.save_settings(self.vars)
            return _ok()
        except Exception as exc:
            return _err(exc)

    def get_recent(self) -> Dict[str, Any]:
        return _ok(files=self.vars.get("recent_files") or [],
                   ranges=self.vars.get("recent_ranges") or [])

    # ============================================================ 打印机
    def list_printers(self) -> Dict[str, Any]:
        try:
            return _ok(printers=list_printers(), default=get_default_printer())
        except Exception as exc:
            return _err(exc)

    def printer_status(self, name: str) -> Dict[str, Any]:
        try:
            status = get_printer_status(name)
            check_printer_ready(name)
            return _ok(status=status, ready=True)
        except Exception as exc:
            return _ok(status={"name": name, "state": "异常"}, ready=False,
                       detail=str(exc))

    def paper_sources(self, printer: str) -> Dict[str, Any]:
        try:
            return _ok(sources=get_printer_sources(printer))
        except Exception as exc:
            return _ok(sources=[], error=str(exc))

    # ============================================================ 文件
    def add_files(self, paths: List[str]) -> Dict[str, Any]:
        added = []
        for path in paths:
            path = str(path)
            if not os.path.isfile(path):
                continue
            if is_pdf_path(path):
                try:
                    pages = get_pdf_page_count(path)
                except Exception as exc:
                    self._log(f"无法读取 {os.path.basename(path)}：{exc}", True)
                    continue
                self._iid += 1
                iid = str(self._iid)
                rec = {"path": path, "pages": pages, "status": "就绪",
                       "range_text": "", "copies": 1, "iid": iid}
                self.files[iid] = rec
                added.append(rec)
            else:
                self._iid += 1
                iid = str(self._iid)
                rec = {"path": path, "pages": 0, "status": "转换中",
                       "range_text": "", "copies": 1, "iid": iid, "source": path}
                self.files[iid] = rec
                self._converting[iid] = rec
                self._executor.submit(self._convert_one, iid, path)
                added.append(rec)
        if added:
            self._log(f"已添加 {len(added)} 个文件，当前共 {len(self.files)} 个待打印。")
            self._add_recent_files([r.get("source") or r["path"] for r in added])
        return _ok(added=added, total=len(self.files))

    def _convert_one(self, iid: str, path: str) -> None:
        name = os.path.basename(path)
        self._log(f"正在转换：{name}")
        rec = self.files[iid]
        try:
            with tempfile_tmp() as temp:
                out = convert_file(path, temp)
                rec["pages"] = get_pdf_page_count(out)
                rec["path"] = out
                rec["status"] = "就绪"
                self._log(f"转换完成：{name}（{rec['pages']} 页）")
        except (WordConvertError, BatchCancelled) as exc:
            rec["status"] = "失败"
            rec["error"] = str(exc)
            self._log(f"转换失败 {name}：{exc}", True)
        except Exception:
            rec["status"] = "失败"
            self._log(f"转换异常 {name}：{traceback.format_exc(limit=1)}", True)
        finally:
            self._converting.pop(iid, None)
            self._post("file-update", dict(rec))

    def _add_recent_files(self, paths: List[str]) -> None:
        files = list(self.vars.get("recent_files") or [])
        for p in reversed(paths):
            if p in files:
                files.remove(p)
            files.insert(0, p)
        self.vars["recent_files"] = files[:12]

    def add_recent(self, path: str) -> Dict[str, Any]:
        if not path:
            return _err(ValueError("路径为空"))
        if os.path.isdir(path):
            exts = supported_extensions_text()
            found = [os.path.join(path, f) for f in sorted(os.listdir(path))
                     if f.lower().endswith(exts)]
            return self.add_files(found)
        return self.add_files([path])

    def remove_files(self, iids: List[str]) -> Dict[str, Any]:
        for iid in iids:
            self.files.pop(iid, None)
        self._log("待打印列表已更新。")
        return _ok(total=len(self.files))

    def clear_files(self) -> Dict[str, Any]:
        self.files.clear()
        self._log("待打印列表已清空。")
        return _ok()

    def set_file_options(self, iid: str, range_text: str, copies: int) -> Dict[str, Any]:
        rec = self.files.get(iid)
        if rec is None:
            return _err(KeyError("文件不存在"))
        rec["range_text"] = range_text or ""
        rec["copies"] = max(1, min(9999, int(copies)))
        if range_text:
            self._add_recent_range(range_text)
        self._post("file-update", dict(rec))
        return _ok()

    def apply_range_to_all(self, range_text: str) -> Dict[str, Any]:
        for rec in self.files.values():
            rec["range_text"] = range_text
        if range_text:
            self._add_recent_range(range_text)
        return _ok()

    def clear_all_ranges(self) -> Dict[str, Any]:
        for rec in self.files.values():
            rec["range_text"] = ""
        return _ok()

    def _add_recent_range(self, text: str) -> None:
        ranges = list(self.vars.get("recent_ranges") or [])
        if text in ranges:
            ranges.remove(text)
        ranges.insert(0, text)
        self.vars["recent_ranges"] = ranges[:8]

    # ============================================================ 计划与预览
    def compute_plan(self, opts: dict) -> Dict[str, Any]:
        """计算批次计划：张数/翻面次数/每文件页数。"""
        try:
            side_counts = []
            layout = self._layout_options(opts)
            for rec in self.files.values():
                pages = rec["pages"]
                rt = (rec.get("range_text") or "").strip()
                if rt:
                    abs_pages, _ignored = resolve_range_pages("custom", rt, pages)
                    if abs_pages:
                        pages = len(abs_pages)
                # 排版后的真实面数（含双面/小册子补页），与 plan_batch_summary 对齐
                plan = build_layout_plan(max(1, pages), layout)
                side_counts.append(plan.output_count)
            copies = max(1, int(opts.get("copies", 1)))
            exec_mode = "per_file" if opts.get("print_flow") == "per_file" else "batch"
            copy_exec = "per_copy" if opts.get("copy_exec_on") else "all"
            summary = plan_batch_summary(side_counts, copies=copies,
                                         exec_mode=exec_mode, copy_exec=copy_exec,
                                         reverse=bool(opts.get("reverse", False)))
            total_pages = sum(side_counts)
            return _ok(files=len(self.files), total_pages=total_pages,
                       summary=summary)
        except Exception as exc:
            return _err(exc)

    def preview_page(self, iid: str, page_no: int, opts: dict,
                     max_width: int = 900) -> Dict[str, Any]:
        """渲染页面缩略预览（含排版后效果）。返回 dataURL。"""
        try:
            rec = self.files.get(iid)
            if rec is None:
                raise KeyError("文件不存在")
            layout = self._layout_options(opts)
            doc = fitz.open(rec["path"])
            if not doc.page_count:
                doc.close()
                return _err(PdfError("文件为空"))
            pno = max(1, min(page_no, doc.page_count))
            page = doc[pno - 1]
            zoom = max_width / max(1.0, page.rect.width)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=bool(opts.get("annots", True)))
            data = base64.b64encode(pix.tobytes("png")).decode("ascii")
            doc.close()
            return _ok(dataurl=f"data:image/png;base64,{data}",
                       page=pno, total=rec["pages"])
        except Exception as exc:
            return _err(exc)

    def plan_description(self, iid: str, opts: dict) -> Dict[str, Any]:
        """返回页序详情：整批页序摘要 + 双面示意文字 + 缩略图列表。"""
        try:
            rec = self.files.get(iid)
            if rec is None:
                raise KeyError("文件不存在")
            layout = self._layout_options(opts)
            pages = list(range(1, rec["pages"] + 1))
            plan = build_layout_plan(len(pages), layout)
            width, height = plan.paper_pt
            metrics = preview_metrics(plan, layout, width, height)
            return _ok(text=describe_plan(plan),
                       sample=self._duplex_sample_text(plan.output_count),
                       metrics=metrics)
        except Exception as exc:
            return _err(exc)

    def _duplex_sample_text(self, sides: int) -> str:
        if sides < 2:
            return "本文件只有 1 面内容，无需翻面。"
        if sides % 2:
            return f"奇数页在前，最后 1 张为空白背面；共 {sides} 面。"
        return f"正面与反面配对打印，共 {sides} 面。"

    # ============================================================ 打印
    def start_print(self, opts: dict) -> Dict[str, Any]:
        if self._printing:
            return _err(RuntimeError("已有打印任务在进行"))
        if not self.files:
            return _err(RuntimeError("没有待打印文件"))
        pending = [r for r in self.files.values() if r["status"] == "转换中"]
        if pending:
            return _err(RuntimeError("部分文件还在转换，请稍候"))
        try:
            snapshot = self._build_snapshot(opts)
        except Exception as exc:
            return _err(exc)
        self.stop_event.clear()
        self.pause_event.clear()
        self._paused = False
        self._printing = True
        self.current_job_id = 0
        self._executor.submit(self._worker_entry, snapshot)
        return _ok()

    def pause_toggle(self) -> Dict[str, Any]:
        if not self._printing:
            return _err(RuntimeError("当前没有进行中的打印任务"))
        if self._paused:
            self.pause_event.clear()
            self._paused = False
            self._log("已继续打印")
        else:
            self.pause_event.set()
            self._paused = True
            self._log("已暂停打印（当前页完成后暂停）")
            if self.current_job_id > 0:
                try:
                    set_job_paused(self.vars.get("printer", ""), self.current_job_id, True)
                except Exception:
                    pass
        return _ok(paused=self._paused)

    def cancel_print(self) -> Dict[str, Any]:
        self.stop_event.set()
        self.pause_event.clear()
        if self.current_job_id > 0:
            try:
                cancel_job(self.vars.get("printer", ""), self.current_job_id)
            except Exception:
                pass
        self._flip_event.set()
        self._phase_confirm_event.set()
        self._log("已请求取消打印。")
        return _ok()

    def confirm_flip(self, proceed: bool) -> Dict[str, Any]:
        """前端翻面模态框回调。"""
        self._flip_result = bool(proceed)
        self._flip_event.set()
        return _ok()

    def confirm_phase(self, proceed: bool) -> Dict[str, Any]:
        """不确定阶段确认回调。"""
        self._phase_confirm_result = bool(proceed)
        self._phase_confirm_event.set()
        return _ok()

    # ---------------------------------------------------- snapshot 构建
    def _layout_options(self, opts: dict) -> LayoutOptions:
        nup = int(opts.get("nup", 1))
        if nup not in NUP_CHOICES:
            nup = 1
        booklet = bool(opts.get("booklet", False))
        sources = self.vars.get("_paper_sources") or [("自动选择", 7)]
        source_code = 7
        for name, code in sources:
            if name == opts.get("source_name", ""):
                source_code = code
                break
        insets = None
        printer = opts.get("printer") or self.vars.get("printer", "")
        if printer:
            try:
                insets = get_printable_insets_pt(printer)
            except Exception:
                insets = None
        return LayoutOptions(
            paper_name=str(opts.get("paper_name", "A4")),
            paper_width_mm=float(opts.get("paper_width_mm", 210.0)),
            paper_height_mm=float(opts.get("paper_height_mm", 297.0)),
            orientation=str(opts.get("orientation", "auto")),
            source_name=str(opts.get("source_name", "")),
            source_code=source_code,
            auto_rotate=bool(opts.get("auto_rotate", True)),
            scale_mode=str(opts.get("scale_mode", "shrink")),
            scale_percent=int(opts.get("scale_percent", 100)),
            nup=2 if booklet else nup,
            nup_border=bool(opts.get("nup_border", False)),
            page_order=str(opts.get("page_order", "row")),
            booklet=booklet,
            margin_left_mm=float(opts.get("margin_left_mm", 10.0)),
            margin_top_mm=float(opts.get("margin_top_mm", 10.0)),
            margin_right_mm=float(opts.get("margin_right_mm", 10.0)),
            margin_bottom_mm=float(opts.get("margin_bottom_mm", 8.0)),
            crop_left_mm=float(opts.get("crop_left_mm", 0.0)),
            crop_top_mm=float(opts.get("crop_top_mm", 0.0)),
            crop_right_mm=float(opts.get("crop_right_mm", 0.0)),
            crop_bottom_mm=float(opts.get("crop_bottom_mm", 0.0)),
            manual_duplex=bool(opts.get("manual_duplex", False)),
            annots=bool(opts.get("annots", True)),
            skip_blank=bool(opts.get("skip_blank", False)),
            printable_insets_pt=insets,
        )

    def _build_snapshot(self, opts: dict) -> dict:
        printer = opts.get("printer") or self.vars.get("printer") or get_default_printer()
        jobs = []
        per_file_settings = {}
        for rec in self.files.values():
            pages = rec["pages"]
            rt = (rec.get("range_text") or "").strip()
            abs_pages = None
            if rt:
                abs_pages, _ignored = resolve_range_pages("custom", rt, pages)
                if abs_pages is None:
                    raise PdfError(f"《{os.path.basename(rec['path'])}》页码范围无效：{rt}")
            job = {"iid": rec["iid"], "path": rec["path"], "pages": pages,
                   "abs_pages": abs_pages, "copies": max(1, int(rec.get("copies", 1))),
                   "name": os.path.basename(rec.get("source") or rec["path"])}
            jobs.append(job)
            per_file_settings[rec["iid"]] = {"range_text": rt,
                                             "copies": job["copies"]}
        layout = self._layout_options(opts)
        copies = max(1, min(9999, int(opts.get("copies", 1))))
        return {
            "printer": printer,
            "jobs": jobs,
            "copies": copies,
            "per_file_settings": per_file_settings,
            "collate": opts.get("collate") == "collate",
            "manual_duplex": bool(opts.get("manual_duplex", False)),
            "print_flow": str(opts.get("print_flow", "batch_all")),
            "copy_exec_on": bool(opts.get("copy_exec_on", False)),
            "reverse": bool(opts.get("reverse", False)),
            "wait_mode": str(opts.get("wait_mode", "queue")),
            "delay": float(opts.get("delay", 20.0)),
            "keep_output": bool(opts.get("keep_output", False)),
            "auto_open": bool(opts.get("auto_open", False)),
            "output_dir": str(opts.get("output_dir", "")),
            "layout": layout,
            "printer_options": self._printer_options(opts, layout),
        }

    def _printer_options(self, opts: dict, layout: LayoutOptions) -> PrinterOptions:
        orientation = opts.get("orientation", "auto")
        if orientation == "landscape":
            landscape = True
        elif orientation == "portrait":
            landscape = False
        else:
            landscape = layout.paper_width_mm > layout.paper_height_mm
        return PrinterOptions(
            landscape=landscape,
            paper_name=layout.paper_name,
            paper_width_mm=layout.paper_width_mm,
            paper_height_mm=layout.paper_height_mm,
            source_code=layout.source_code,
            color_mode=str(opts.get("color_mode", "color")),
            print_quality=str(opts.get("print_quality", "normal")),
            dpi=int(opts.get("dpi", 300)),
            ink_saver=bool(opts.get("ink_saver", False)),
        )

    # ============================================================ 工作线程
    def _worker_entry(self, snapshot: dict) -> None:
        completed = cancelled = False
        error_text = ""
        try:
            self._worker_main(snapshot)
            completed = True
        except CancelledByUser:
            cancelled = True
            self._log("打印已取消。", True)
        except Exception as exc:
            error_text = str(exc)
            self._log(f"打印失败：{error_text}", True)
            traceback.print_exc()
        finally:
            self._printing = False
            self._post("print-done", {"completed": completed,
                                      "cancelled": cancelled, "error": error_text})

    def _worker_main(self, snapshot: dict) -> None:
        prepared = []
        stamp = ""
        with tempfile_tmp() as target_dir:
            stamp = os.path.basename(target_dir)
            for job in snapshot["jobs"]:
                self._raise_if_stopped()
                name = job["name"]
                self._post("rows", [(job["iid"], "正在排版")])
                self._log(f"正在排版：《{name}》")
                abs_pages = job["abs_pages"]
                if abs_pages is None:
                    abs_pages = list(range(1, job["pages"] + 1))
                out_path = os.path.join(target_dir, f"{self._safe_filename(name)}.pdf")
                result = compose_layout_pdf(
                    job["path"], abs_pages, out_path, snapshot["layout"])
                plan = build_layout_plan(len(abs_pages), snapshot["layout"])
                self._log(f"排版完成：《{name}》 → {result.selected_count} 面。")
                prepared.append({"job": job, "path": out_path,
                                 "sides": result.selected_count, "plan": plan})
            copies = snapshot["copies"]
            total_sides = sum(p["sides"] for p in prepared)
            manual = snapshot["manual_duplex"]
            if manual:
                self._log(
                    f"手动双面：整批 {total_sides} 面（{total_sides // 2} 张 × 份数 {copies}），"
                    "先打全部正面，取纸翻面后再打反面。")
            if manual and snapshot["print_flow"] == "per_file":
                self._run_by_file_duplex(snapshot, prepared, target_dir, stamp)
            elif manual:
                self._run_batch_duplex(snapshot, prepared, target_dir, stamp)
            else:
                self._run_single_sided(snapshot, prepared, target_dir, stamp)
        self._log("全部打印阶段已完成。")
        if snapshot.get("auto_open") and snapshot.get("output_dir"):
            open_directory(snapshot["output_dir"])

    def _safe_filename(self, name: str) -> str:
        keep = "".join(c if c not in r'\/:*?"<>|' else "_" for c in name)
        return keep or "未命名"

    def _raise_if_stopped(self) -> None:
        if self.stop_event.is_set():
            raise CancelledByUser("用户取消。")

    def _sleep_interruptible(self, seconds: float) -> None:
        import time
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self._raise_if_stopped()
            time.sleep(min(0.2, max(0.0, end - time.monotonic())))

    # ---------------------------------------------------- 单面
    def _run_single_sided(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
        for item in prepared:
            job = item["job"]
            sides = item["sides"]
            copies = item["job"]["copies"]
            order = compute_single_sided_orders(sides, copies,
                                                snapshot["collate"], snapshot["reverse"])
            self._print_phase(item["path"], order, f"PDF {job['name']} 单面",
                              snapshot, f"单面打印 · 《{job['name']}》")
            self._post("rows", [(job["iid"], "已完成")])
            self._log(f"单面打印完成：《{job['name']}》 {sides} 个输出面 × {copies} 份。")

    # ---------------------------------------------------- 手动双面
    def _run_batch_duplex(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
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

        # 单面特例
        if total == 1:
            order = compute_single_sided_orders(1, snapshot["copies"])
            self._print_phase(merged_path, order, "PDF 单面打印", snapshot,
                              "单面打印 · 正序")
            return

        fronts, backs = compute_merged_duplex_orders(total, snapshot["copies"],
                                                     snapshot["reverse"], ranges)
        self._post("rows", [(item["job"]["iid"], "正面打印中") for item in prepared])
        self._log(f"手动双面 · 正面：{len(fronts)} 个输出面（{sheets} 张），"
                  f"共 {snapshot['copies']} 份。")
        self._print_phase(merged_path, fronts, "PDF 双面正面", snapshot,
                          "手动双面 · 打印正面")
        if not self._ask_flip_batch(counts, fronts, backs, snapshot, sheets):
            raise CancelledByUser("用户取消翻面。")

        self._post("rows", [(item["job"]["iid"], "反面打印中") for item in prepared])
        if pause_between_files:
            # 反面按文件从后往前打，文件之间暂停取纸
            idx = 0
            for item in reversed(prepared):
                job = item["job"]
                sides = item["sides"]
                side_backs = backs[idx:idx + sides]
                idx += sides
                self._print_phase(item["path"], side_backs,
                                  f"PDF {job['name']} 反面", snapshot,
                                  f"手动双面 · 《{job['name']}》反面")
                self._post("rows", [(job["iid"], "已完成")])
                self._log(f"《{job['name']}》反面完成。")
        else:
            self._print_phase(merged_path, backs, "PDF 双面反面", snapshot,
                              "手动双面 · 打印反面")
        self._batch_done_modal(total, sheets, snapshot)

    def _run_by_file_duplex(self, snapshot: dict, prepared: List[dict]) -> None:
        for item in prepared:
            job = item["job"]
            sides = item["sides"]
            fronts, backs = compute_duplex_orders(sides, job["copies"],
                                                  snapshot["reverse"])
            self._post("rows", [(job["iid"], "正面打印中")])
            self._print_phase(item["path"], fronts, f"PDF {job['name']} 正面",
                              snapshot, f"逐文件双面 · 《{job['name']}》正面")
            if not self._ask_flip_file(job, sides, backs, snapshot):
                raise CancelledByUser("用户取消翻面。")
            self._post("rows", [(job["iid"], "反面打印中")])
            self._print_phase(item["path"], backs, f"PDF {job['name']} 反面",
                              snapshot, f"逐文件双面 · 《{job['name']}》反面")
            self._post("rows", [(job["iid"], "已完成")])
            self._file_done_modal(job, sides, snapshot)

    # ---------------------------------------------------- 打印阶段
    def _print_phase(self, pdf_path: str, order: List[int], job_name: str,
                     snapshot: dict, phase_text: str) -> None:
        if not order:
            return
        self._raise_if_stopped()
        self._post("status", phase_text)
        self._log(f"{phase_text}：共 {len(order)} 个输出面。")
        import time
        started_at = time.time()

        def on_page(index: int, total: int, rel: int) -> None:
            self._post("progress", {"index": index, "total": total,
                                    "text": f"{phase_text} · {index}/{total} 面"})

        printer = GdiPrinter(snapshot["printer"], snapshot["printer_options"])
        try:
            job_id = printer.print_pages(
                pdf_path, order, job_name, self.stop_event,
                on_page=on_page, pause_event=self.pause_event,
                on_job_start=lambda v: setattr(self, "current_job_id", v))
        finally:
            self.current_job_id = printer.last_job_id
        if job_id <= 0:
            raise PrintError("打印任务提交失败。")

        if snapshot["wait_mode"] == "queue":
            try:
                finished = wait_for_print_job(
                    snapshot["printer"], job_id, self.stop_event,
                    log=lambda t: self._log(t), pause_event=self.pause_event)
                if not finished:
                    raise CancelledByUser("用户取消。")
            except (PrintTimeoutError, PrintJobUnknownError) as exc:
                self._confirm_uncertain_phase(phase_text, str(exc))
        else:
            self._sleep_interruptible(snapshot["delay"])
        self._sleep_interruptible(2.0)  # POST_JOB_SETTLE_SECONDS
        self._log(f"{phase_text}：任务 #{job_id} 结束，{len(order)} 个输出面，"
                  f"用时 {time.time() - started_at:.1f} 秒。")

    # ---------------------------------------------------- 模态确认
    def _ask_flip_batch(self, counts: List[int], fronts: List[int], backs: List[int],
                        snapshot: dict, sheets: int) -> bool:
        self._raise_if_stopped()
        detail = [
            f"本批 {sheets} 张 × {snapshot['copies']} 份已打印正面完成。",
            "请取出打印纸，保持页序，整体翻面后重新放入纸盒。",
            "确认后点击「已翻面，继续打印反面」；取消将终止本次任务。",
        ]
        return self._flip_modal("手动双面 · 请翻面", detail,
                                self._duplex_sample_text(sum(counts)))

    def _ask_flip_file(self, job: dict, sides: int, backs: List[int],
                       snapshot: dict) -> bool:
        self._raise_if_stopped()
        detail = [
            f"《{job['name']}》正面已打印完成（{len(backs)} 个反面待打）。",
            "请取回该文件的打印纸，保持页序，翻面后放入纸盒。",
        ]
        return self._flip_modal("逐文件双面 · 请翻面", detail,
                                self._duplex_sample_text(sides))

    def _flip_modal(self, title: str, detail_lines: List[str], sample: str) -> bool:
        self._flip_event.clear()
        self._post("flip-modal", {"title": title, "detail": detail_lines,
                                  "sample": sample})
        self._flip_event.wait()
        return self._flip_result

    def _file_done_modal(self, job: dict, sides: int, snapshot: dict) -> None:
        self._log(f"《{job['name']}》已完成，{sides} 个输出面。")

    def _batch_done_modal(self, total: int, sheets: int, snapshot: dict) -> None:
        self._log(f"整批双面完成：{total} 个输出面，{sheets} 张。")

    def _confirm_uncertain_phase(self, phase_text: str, reason: str) -> None:
        self._phase_confirm_event.clear()
        self._post("phase-confirm", {"text": phase_text, "reason": reason})
        self._phase_confirm_event.wait()
        if not self._phase_confirm_result:
            raise CancelledByUser(f"{phase_text} 状态无法确认，用户取消。")

    # ============================================================ 杂项
    def cleanup_stale(self) -> Dict[str, Any]:
        removed = cleanup_stale_temp_dirs()
        if removed:
            self._log(f"已清理 {len(removed)} 个上次残留的临时目录。")
        return _ok(removed=len(removed))

    def converters_available(self) -> Dict[str, Any]:
        return _ok(converters=available_converters())

    def open_log_dir(self) -> Dict[str, Any]:
        try:
            open_directory(os.path.dirname(settings_store.settings_path()))
            return _ok()
        except Exception as exc:
            return _err(exc)


class tempfile_tmp:
    """上下文管理器：返回 create_temp_dir 路径，退出清理。"""

    def __enter__(self) -> str:
        self.path = create_temp_dir()
        return self.path

    def __exit__(self, *_exc) -> None:
        import shutil
        shutil.rmtree(self.path, ignore_errors=True)
