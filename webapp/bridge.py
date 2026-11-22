# -*- coding: utf-8 -*-
"""js_api 桥：前端 window.pywebview.api.* 的后端实现。

原则：UI 层（Tk）已删，业务逻辑直接调用既有核心模块。
所有方法返回 JSON 可序列化的 dict；耗时操作在线程池执行，
通过 window.evaluate_js 回推进度/日志，保持界面丝滑。
"""
import base64
import os
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional

import fitz

import printer_core
import settings_store
from core_utils import is_pdf_path, is_word_path, is_wps_path
from pdf_layout import (
    LayoutOptions, build_layout_plan, compose_layout_pdf, describe_plan,
    get_pdf_page_count, plan_batch_summary, preview_metrics,
    resolve_range_pages, build_range_pdf,
)
from word_convert import (
    BatchCancelled, WordConvertError, available_converters,
    cleanup_stale_temp_dirs, convert_file, create_temp_dir,
)
from image_convert import image_to_pdf, is_image_path

_logger_poll = None


def _ok(**kw) -> Dict[str, Any]:
    return {"ok": True, **kw}


def _err(exc: Exception) -> Dict[str, Any]:
    return {"ok": False, "error": type(exc).__name__, "message": str(exc)}


class AppBridge:
    """暴露给 JS 的 API；pywebview 自动把公开方法映射为
    window.pywebview.api.<method>。"""

    def __init__(self) -> None:
        self._window = None          # webview.Window，attach 后可用
        self._executor = ThreadPoolExecutor(max_workers=2)
        self._stop = threading.Event()
        self._paused = threading.Event()
        # 文件表：iid -> {path, pages, status, range_text, copies}
        self.files: Dict[str, Dict[str, Any]] = {}
        self._iid = 0
        # 打印中标记
        self._printing = False

    # ---------------------------------------------------------------- setup
    def attach(self, window) -> None:
        self._window = window

    def _js(self, code: str) -> None:
        if self._window is not None:
            try:
                self._window.evaluate_js(code)
            except Exception:
                pass

    def _log(self, text: str, error: bool = False) -> None:
        import time
        stamp = time.strftime("%H:%M:%S")
        payload = base64.b64encode(text.encode("utf-8")).decode("ascii")
        self._js(f"window.__pushLog('{stamp}', atob('{payload}'), {str(error).lower()})")

    # ------------------------------------------------------------ settings
    def get_settings(self) -> Dict[str, Any]:
        return _ok(settings=settings_store.load_settings())

    def save_settings(self, settings: dict) -> Dict[str, Any]:
        try:
            settings_store.save_settings(settings)
            return _ok()
        except Exception as exc:
            return _err(exc)

    # ------------------------------------------------------------ printers
    def list_printers(self) -> Dict[str, Any]:
        try:
            return _ok(printers=printer_core.list_printers(),
                       default=printer_core.get_default_printer())
        except Exception as exc:
            return _err(exc)

    def refresh_printers(self) -> Dict[str, Any]:
        return self.list_printers()

    def printer_status(self, name: str) -> Dict[str, Any]:
        try:
            status = printer_core.get_printer_status(name)
            ready, problem = printer_core.check_printer_ready(name), None
            return _ok(status=status, ready=True)
        except Exception as exc:
            return _ok(status={"name": name, "state": "异常", "detail": str(exc)},
                       ready=False)

    # --------------------------------------------------------------- files
    def add_files(self, paths: List[str]) -> Dict[str, Any]:
        """批量登记文件；Word/图片先转 PDF（返回转换任务句柄）。"""
        added = []
        to_convert = []
        for path in paths:
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
                self.files[iid] = {"path": path, "pages": pages,
                                   "status": "就绪", "range_text": "",
                                   "copies": 1}
                added.append(self.files[iid] | {"iid": iid})
            else:
                to_convert.append(path)
        result = _ok(added=added)
        if to_convert:
            self._executor.submit(self._convert_files, to_convert)
            result["converting"] = len(to_convert)
        return result

    def _convert_files(self, paths: List[str]) -> None:
        temp = create_temp_dir()
        try:
            for path in paths:
                if self._stop.is_set():
                    break
                name = os.path.basename(path)
                self._log(f"正在转换：{name}")
                try:
                    out = convert_file(path, temp)
                    pages = get_pdf_page_count(out)
                    self._iid += 1
                    iid = str(self._iid)
                    record = {"path": out, "pages": pages,
                              "status": "就绪", "range_text": "",
                              "copies": 1, "source": path}
                    self.files[iid] = record
                    payload = {"iid": iid, **record}
                    import json
                    self._js(f"window.__fileConverted({json.dumps(payload, ensure_ascii=False)})")
                    self._log(f"转换完成：{name}（{pages} 页）")
                except (WordConvertError, BatchCancelled) as exc:
                    self._log(f"转换失败 {name}：{exc}", True)
        except Exception:
            self._log("转换线程异常：" + traceback.format_exc(limit=1), True)
        finally:
            import shutil
            shutil.rmtree(temp, ignore_errors=True)

    def remove_files(self, iids: List[str]) -> Dict[str, Any]:
        for iid in iids:
            self.files.pop(iid, None)
        return _ok()

    def clear_files(self) -> Dict[str, Any]:
        self.files.clear()
        return _ok()

    def set_file_options(self, iid: str, range_text: str, copies: int) -> Dict[str, Any]:
        rec = self.files.get(iid)
        if rec is None:
            return _err(KeyError("文件不存在"))
        rec["range_text"] = range_text
        rec["copies"] = max(1, int(copies))
        return _ok()

    # -------------------------------------------------------------- plan
    def compute_plan(self, options: Dict[str, Any]) -> Dict[str, Any]:
        """按当前文件表与打印选项计算批次计划（张数/翻面次数/页序描述）。"""
        try:
            side_counts = []
            for rec in self.files.values():
                pages = rec["pages"]
                if rec["range_text"]:
                    abs_pages = resolve_range_pages(rec["range_text"], pages)
                    pages = len(abs_pages) if abs_pages is not None else pages
                side_counts.append(max(1, pages))
            copies = int(options.get("copies", 1))
            manual = bool(options.get("manual_duplex", False))
            summary = plan_batch_summary(side_counts, copies=copies,
                                         manual_duplex=manual)
            return _ok(summary=summary, files=len(self.files))
        except Exception as exc:
            return _err(exc)

    def preview_page(self, iid: str, page_no: int, options: Dict[str, Any],
                     max_width: int = 900) -> Dict[str, Any]:
        """渲染指定页为 PNG dataURL 供前端预览。"""
        try:
            rec = self.files.get(iid)
            if rec is None:
                raise KeyError("文件不存在")
            doc = fitz.open(rec["path"])
            if page_no < 1 or page_no > doc.page_count:
                page_no = 1
            page = doc[page_no - 1]
            zoom = max_width / max(1.0, page.rect.width)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            data = base64.b64encode(pix.tobytes("png")).decode("ascii")
            doc.close()
            return _ok(dataurl=f"data:image/png;base64,{data}",
                       page=page_no, total=rec["pages"])
        except Exception as exc:
            return _err(exc)

    # ------------------------------------------------------------- print
    def start_print(self, options: Dict[str, Any]) -> Dict[str, Any]:
        if self._printing:
            return _err(RuntimeError("已有打印任务在进行"))
        if not self.files:
            return _err(RuntimeError("没有待打印文件"))
        self._printing = True
        self._stop.clear()
        self._executor.submit(self._print_worker, dict(options))
        return _ok()

    def cancel_print(self) -> Dict[str, Any]:
        self._stop.set()
        return _ok()

    def _print_worker(self, options: Dict[str, Any]) -> None:
        try:
            self._do_print(options)
        except Exception as exc:
            self._log(f"打印失败：{exc}", True)
        finally:
            self._printing = False
            self._js("window.__printDone()")

    def _do_print(self, options: Dict[str, Any]) -> None:
        printer = options.get("printer") or printer_core.get_default_printer()
        copies = max(1, int(options.get("copies", 1)))
        manual = bool(options.get("manual_duplex", False))
        lo = LayoutOptions(manual_duplex=manual)
        po = printer_core.PrinterOptions(copies=copies)
        self._log(f"开始打印：{len(self.files)} 个文件 → {printer}")
        temp = create_temp_dir()
        try:
            composed_paths: List[str] = []
            side_counts: List[int] = []
            for rec in self.files.values():
                if self._stop.is_set():
                    self._log("已取消打印", True)
                    return
                src = rec["path"]
                name = os.path.basename(rec.get("source") or src)
                self._log(f"排版：{name}")
                out = os.path.join(temp, f"composed-{len(composed_paths)}.pdf")
                result = compose_layout_pdf(src, [i + 1 for i in range(rec["pages"])],
                                            out, lo)
                composed_paths.append(out)
                side_counts.append(result.selected_count)
                self._log(f"排版完成：{name} → {result.selected_count} 面")
            with printer_core.GdiPrinter(printer, po) as gp:
                for i, path in enumerate(composed_paths):
                    if self._stop.is_set():
                        self._log("已取消打印", True)
                        return
                    doc = fitz.open(path)
                    for pno in range(doc.page_count):
                        if self._stop.is_set():
                            doc.close()
                            self._log("已取消打印", True)
                            return
                        gp.print_page(doc[pno])
                    doc.close()
                    self._log(f"已送出 {i + 1}/{len(composed_paths)} 个文件")
            self._log("打印完成" + ("，请按提示翻面后继续第二面" if manual else ""))
        finally:
            import shutil
            shutil.rmtree(temp, ignore_errors=True)

    # ------------------------------------------------------------- misc
    def cleanup_stale(self) -> Dict[str, Any]:
        removed = cleanup_stale_temp_dirs()
        if removed:
            self._log(f"已清理 {len(removed)} 个上次残留的临时目录")
        return _ok(removed=len(removed))

    def converters_available(self) -> Dict[str, Any]:
        return _ok(converters=available_converters())

    def open_log_dir(self) -> Dict[str, Any]:
        try:
            printer_core.open_directory(os.path.dirname(settings_store.settings_path()))
            return _ok()
        except Exception as exc:
            return _err(exc)
