# -*- coding: utf-8 -*-
"""js_api 桥（完整移植版）：对齐 Tk 版全部功能。

每个方法返回 {"ok": bool, ...}；耗时操作走线程池。
前端通过 window.pywebview.api.<method> 调用。
"""
import base64
import json
import os
import queue
import re
import shutil
import subprocess
import threading
import time
import traceback
from concurrent.futures import Future
from dataclasses import replace as dataclass_replace
from typing import Any, Dict, List, Optional, Sequence, Tuple

import fitz

import printer_core
import settings_store
import window_geometry
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
    merge_layout_pdfs, merged_side_ranges, merged_sheet_count, open_pdf,
    plan_batch_summary, preview_metrics, resolve_range_pages,
)
from word_convert import (
    BatchCancelled, WordConvertError, available_converters,
    cleanup_stale_temp_dirs, convert_file, create_temp_dir,
)
from image_convert import ImageConvertError, image_to_pdf, is_image_path
from printer_core import (
    GdiPrinter, PrinterOptions, cancel_job, check_printer_ready, get_default_printer,
    get_printable_insets_pt, get_printer_sources, get_printer_status, list_printers,
    open_directory, set_job_paused, wait_for_print_job,
)
from log_setup import get_logger


def _ok(**kw) -> Dict[str, Any]:
    return {"ok": True, **kw}


_PY_EXC_PREFIX = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception|Warning):\s*")


def _describe_error(exc: BaseException, fallback: str = "操作未完成，请稍后重试。") -> str:
    """把异常转成可直接展示给用户的中文文本。

    各业务异常（printer_core / pdf_layout）的消息本身已是中文且可操作；
    这里只去掉可能出现的 “TypeError:” 之类类名前缀，并给空消息兜底，
    避免界面出现裸露的 Python 异常名或“未知错误”。
    """
    text = str(exc or "").strip()
    text = _PY_EXC_PREFIX.sub("", text).strip()
    if not text or text.startswith("Traceback"):
        return fallback
    return text


def _err(exc: Exception) -> Dict[str, Any]:
    return {"ok": False, "error": type(exc).__name__, "message": str(exc)}


# 执行流程标签：与界面三个单选值一一对应，只用于日志展示，绝不改写流程。
PRINT_FLOW_LABELS = {
    "batch_all": "整批-最后取纸",
    "batch_take": "整批-逐文件取纸",
    "per_file": "逐文件执行",
}


def print_flow_of(opts: dict) -> str:
    """返回界面选择的执行流程，非法/缺失值退回 batch_all（不推断、不改写）。"""
    flow = str(opts.get("print_flow") or "")
    return flow if flow in PRINT_FLOW_LABELS else "batch_all"


def _uses_per_file_copies(opts: dict) -> bool:
    """有效副本数规则：只有“逐文件执行”流程才采用文件级份数。

    整批流程（batch_all/batch_take）一次任务打完所有文件，份数只能取全局
    copies；此时文件行残留的旧份数必须忽略，避免界面显示与实际打印不一致。
    """
    return print_flow_of(opts) == "per_file"


# WeType（微信输入法）是纯 TSF 输入法：实测 IMM32 的 open 状态在中文/英文下
# 都是 0，HIMC 也恒为 0。因此不能"读状态再决定是否切换"，只能在窗口稳定后
# 盲发一次 Shift；重复校验会像 InputTip 一样把刚切好的中文再切回英文。
IME_RESTORE_DELAY_SECONDS = 0.4


def _is_own_window_visible(window) -> bool:
    """本应用窗口是否可见且未最小化。

    不能用 GetForegroundWindow 判断：本机壁纸程序的 BkShadowWndClass
    会长期占据前景，应用窗口明明可见也会被误判成“不在前台”。启动场景下
    “窗口可见且未最小化”已经足够表达“用户刚打开本应用”。
    """
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        user32.IsWindowVisible.argtypes = [wintypes.HWND]
        user32.IsWindowVisible.restype = wintypes.BOOL
        user32.IsIconic.argtypes = [wintypes.HWND]
        user32.IsIconic.restype = wintypes.BOOL
        hwnd = int(window.native.Handle.ToInt64())
        return bool(user32.IsWindowVisible(hwnd)) and not bool(user32.IsIconic(hwnd))
    except Exception:
        return False


def _send_single_shift() -> None:
    """发送一次真实 Shift（按下 + 抬起），不做任何状态回读。"""
    import ctypes
    user32 = ctypes.windll.user32
    user32.keybd_event(0x10, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(0x10, 0, 2, 0)


def _is_supported_source(path: str) -> bool:
    return bool(is_pdf_path(path) or is_word_path(path) or is_wps_path(path)
                or is_image_path(path))


_FILE_DIALOG_TYPES = (
    "支持的文件 (*.pdf;*.doc;*.docx;*.rtf;*.txt;*.odt;*.wps;"
    "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.webp)",
    "PDF 文件 (*.pdf)",
    "所有文件 (*.*)",
)

# 翻面/状态确认弹窗的默认等待上限（秒）。超时按“用户取消”处理，
# 避免前端关闭窗口或 JS 失效时工作线程永久阻塞、永远走不到汇总。
MODAL_WAIT_TIMEOUT_SECONDS = 300.0

# webapp 窗口默认/最小尺寸：逻辑像素（96 DPI 基准），与 pywebview
# create_window/resize 使用的单位一致。物理像素换算见 window_geometry。
WINDOW_DEFAULT_SIZE = (1280, 860)
WINDOW_MIN_SIZE = (1024, 680)

# 窗口移动/缩放停止后延迟落盘的秒数：拖动期间事件密集，避免频繁写盘；
# 关闭时还会保存一次最终几何作为兜底。
GEOMETRY_SAVE_DEBOUNCE_SECONDS = 0.6


class DaemonThreadPool:
    """线程池子集实现：工作线程为 daemon，窗口关闭后不阻塞进程退出。

    CPython 3.9+ 的 ``ThreadPoolExecutor`` 使用非守护线程，并在解释器退出时
    由 atexit 钩子 join 所有工作线程。打印任务（排版、等待打印队列）可能持续
    数十秒，会让“关闭窗口后进程退出”被明显拖慢。这里只实现本项目用到的
    ``submit`` / ``shutdown`` 语义，用 daemon 线程执行任务，关闭窗口后可随
    主线程立即退出；未完成任务的异常仍落在 ``Future`` 上，不影响排查。
    """

    def __init__(self, max_workers: int = 3,
                 thread_name_prefix: str = "pdf-duplex") -> None:
        self._max_workers = max(1, int(max_workers))
        self._thread_name_prefix = thread_name_prefix
        self._queue: "queue.Queue" = queue.Queue()
        self._threads: List[threading.Thread] = []
        self._lock = threading.Lock()
        self._closed = False

    # ------------------------------------------------------------ 对外接口
    def submit(self, fn, *args, **kwargs) -> "Future":
        future: "Future" = Future()
        with self._lock:
            if self._closed:
                raise RuntimeError("线程池已关闭")
            self._queue.put((future, fn, args, kwargs))
            self._spawn_worker_locked()
        return future

    def shutdown(self, wait: bool = True, cancel_futures: bool = False) -> None:
        """停止接收新任务；cancel_futures 时丢弃排队任务，wait 时等待收尾。"""
        with self._lock:
            self._closed = True
            threads = list(self._threads)
        if cancel_futures:
            while True:
                try:
                    future, _fn, _args, _kwargs = self._queue.get_nowait()
                except queue.Empty:
                    break
                future.cancel()
                self._queue.task_done()
        for _thread in threads:
            self._queue.put(None)  # 哨兵：让每个工作线程在处理完手头任务后退出
        if wait:
            for thread in threads:
                thread.join()

    # ------------------------------------------------------------ 内部实现
    def _spawn_worker_locked(self) -> None:
        if len(self._threads) >= self._max_workers:
            return
        thread = threading.Thread(
            target=self._run_worker,
            name=f"{self._thread_name_prefix}_{len(self._threads)}",
            daemon=True)
        self._threads.append(thread)
        thread.start()

    def _run_worker(self) -> None:
        while True:
            item = self._queue.get()
            if item is None:
                self._queue.task_done()
                return
            try:
                future, fn, args, kwargs = item
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    future.set_result(fn(*args, **kwargs))
                except BaseException as exc:  # 任务异常交给 Future 保存，不丢失
                    future.set_exception(exc)
            finally:
                self._queue.task_done()


class AppBridge:
    def __init__(self) -> None:
        self._window = None
        # 工作线程全部为 daemon：窗口关闭后不再被 atexit join 拖住退出。
        self._executor = DaemonThreadPool(max_workers=3)
        self._shutdown_event = threading.Event()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.current_job_id = 0
        # 打印任务占位锁：start_print 在锁内检查并立即置位，杜绝并发双启动。
        self._print_lock = threading.Lock()
        self._printing = False
        self._paused = False
        self._flip_event = threading.Event()
        self._flip_result = False
        self._phase_confirm_event = threading.Event()
        self._phase_confirm_result = False
        # 模态框等待上限，测试可 monkeypatch 成很小的值。
        self._modal_timeout = MODAL_WAIT_TIMEOUT_SECONDS
        # 文件表（iid 字符串 -> dict）
        self.files: Dict[str, Dict[str, Any]] = {}
        self._iid = 0
        # 当前生效设置（来自 settings.json，前端修改后同步回写）
        self.vars: Dict[str, Any] = dict(settings_store.load_settings())
        # 打印历史摘要
        self.last_print_summary: Dict[str, Any] = {}
        # 转换进度跟踪（前端轮询）
        self._converting: Dict[str, Dict[str, Any]] = {}
        # 窗口关闭已确认后置位：禁止再从任意线程调用 evaluate_js。
        # evaluate_js 会阻塞等待 GUI 同步上下文回调，在关闭流程中投递 JS
        # 会与 GUI 消息循环互等导致死锁。
        self._suppress_js = False
        # 窗口几何记忆：moved/resized 只在 GUI/事件线程标记“待保存”，
        # Win32 读取与设置落盘交给 daemon 线程，避免拖动时卡住消息循环。
        self._geometry_lock = threading.Lock()
        self._geometry_write_lock = threading.Lock()
        self._geometry_save_pending = False
        self._geometry_save_delay = GEOMETRY_SAVE_DEBOUNCE_SECONDS
        self._geometry_closed = False

    # ============================================================ 基础设施
    def attach(self, window) -> None:
        self._window = window
        # 原生拖放只需启用 pywebview 的路径收集通道。不要在 webview.start()
        # 之前访问 window.dom.body：DOM API 会等待尚未创建的窗口并固定阻塞 20 秒。
        # 前端已有真实 drop 监听；这里保证 on_script_notify 始终接收 FilesDropped。
        try:
            from webview.dom import _dnd_state
            _dnd_state["num_listeners"] = max(1, _dnd_state["num_listeners"])
        except Exception:
            pass
        self._executor.submit(self._drop_pump)
        # 窗口图标（Win32 层）：用项目自带 app_icon.ico 替换默认 python 图标
        self._executor.submit(self._apply_window_icon)
        # 微信输入法在本应用激活时会跳到英文；窗口稳定后补一次 Shift。
        self._executor.submit(self._restore_ime)
        window.events.closed += self._on_window_closed
        # 打印中关闭窗口先确认。pywebview 6.x 以 closing.set() 的返回值表示
        # 是否取消关闭（任一处理器返回 False 即取消），且为同步回调。
        closing_event = getattr(window.events, "closing", None)
        if closing_event is not None:
            closing_event += self._on_window_closing
        self._bind_geometry_events(window)
        if os.environ.get("WV_SELFTEST") == "1":
            threading.Thread(target=self._selftest, daemon=True).start()

    def _restore_ime(self) -> None:
        """补偿微信输入法在窗口激活时跳到英文的问题。

        前提：WeType 的 IMM32 open 在中文/英文下均为 0，无法可靠读状态；
        只在前台确实是本窗口时发送一次 Shift，且不回读、不重试。
        PDF_DUPLEX_RESTORE_IME=0 可关闭；自检模式不发送。
        """
        if os.environ.get("PDF_DUPLEX_RESTORE_IME") == "0":
            return
        if os.environ.get("WV_SELFTEST") == "1":
            return
        window = getattr(self, "_window", None)
        if window is None:
            return
        try:
            shown = getattr(window.events, "shown", None)
            if shown is not None and not shown.wait(15):
                return
            time.sleep(IME_RESTORE_DELAY_SECONDS)
            if not _is_own_window_visible(window):
                self._log("启动恢复输入法：窗口不可见，跳过")
                return
            _send_single_shift()
            self._log("启动恢复输入法：已发送 Shift")
        except Exception:
            get_logger().exception("启动恢复输入法失败")

    def _bind_geometry_events(self, window) -> None:
        """订阅 moved/resized，移动或缩放结束后记忆窗口几何。"""
        events = getattr(window, "events", None)
        for name in ("moved", "resized"):
            event = getattr(events, name, None)
            if event is None:
                continue
            try:
                event += self._on_window_geometry_event
            except Exception:
                get_logger().exception(f"订阅窗口 {name} 事件失败")

    def _on_window_geometry_event(self, *_args) -> None:
        """窗口移动/缩放后请求保存几何（在 pywebview 事件线程执行）。

        pywebview 的 moved/resized 是非锁事件，处理器本身就在独立线程；
        这里仍只做标记与投递：Win32 查询和设置落盘交给 daemon 线程，
        既避免拖动时阻塞消息循环，也绝不在此调用 evaluate_js。
        """
        if self._geometry_closed or self._suppress_js:
            return
        with self._geometry_lock:
            if self._geometry_save_pending:
                return
            self._geometry_save_pending = True
        try:
            self._executor.submit(self._flush_geometry_save)
        except Exception:
            # 线程池已关闭（退出流程）：最终几何由 closing 回调保存。
            with self._geometry_lock:
                self._geometry_save_pending = False

    def _flush_geometry_save(self) -> None:
        """去抖后保存窗口几何（daemon 线程）。"""
        try:
            delay = float(self._geometry_save_delay)
            if delay > 0:
                time.sleep(delay)
            self.save_window_geometry()
        finally:
            with self._geometry_lock:
                self._geometry_save_pending = False

    def _window_handle(self) -> int:
        """返回 WebView2 顶层窗口句柄；窗口未就绪时返回 0。"""
        native = getattr(self._window, "native", None)
        handle = getattr(native, "Handle", None)
        try:
            return int(handle.ToInt64())
        except Exception:
            return 0

    def _capture_window_geometry(
            self) -> Optional[Tuple[Tuple[int, int, int, int], float]]:
        """读取当前窗口几何（逻辑像素）与窗口 DPI 缩放。

        最大化/最小化由 window_geometry 取还原态，避免保存整屏尺寸或
        最小化时的 -32000 离屏坐标。
        """
        hwnd = self._window_handle()
        if not hwnd:
            return None
        geometry = window_geometry.get_window_geometry(hwnd)
        if geometry is None:
            return None
        return geometry, window_geometry.window_dpi_scale(hwnd)

    @staticmethod
    def compose_window_geometry(
            geometry: Sequence[int], work_areas, scale: float,
            default_size=WINDOW_DEFAULT_SIZE,
            min_size=WINDOW_MIN_SIZE) -> str:
        """把逻辑几何整理成可保存的 Tk 风格几何串（纯函数，便于测试）。

        几何落在某块工作区内时恢复（越界部分夹回工作区）；否则使用默认
        尺寸在主屏居中，覆盖外接屏断开后窗口跑到屏幕外的情况。
        """
        chosen = window_geometry.choose_logical_geometry(
            window_geometry.format_geometry(*geometry), work_areas, scale,
            default_size, min_size)
        return window_geometry.format_geometry(*chosen)

    def save_window_geometry(self) -> bool:
        """保存当前窗口几何到 settings.json；失败只记录日志。

        返回 True 表示已保存或几何未变化，False 表示窗口不可用/保存失败。
        该方法可在任意线程调用，内部不做任何 JS 投递。
        """
        # 串行化：移动/缩放去抖线程与关闭回调用同一把锁，避免并发改写
        # self.vars 或互相覆盖刚写入的几何字符串。
        with self._geometry_write_lock:
            return self._save_window_geometry_locked()

    def _save_window_geometry_locked(self) -> bool:
        try:
            captured = self._capture_window_geometry()
            if captured is None:
                return False
            geometry, scale = captured
            try:
                work_areas = window_geometry.list_work_areas()
            except Exception:
                get_logger().exception("枚举显示器工作区失败，按默认居中回退")
                work_areas = []
            saved = self.compose_window_geometry(geometry, work_areas, scale)
            previous = str(self.vars.get("window_geometry") or "")
            if saved == previous:
                return True
            self.vars["window_geometry"] = saved
            try:
                settings_store.save_settings(self.vars)
            except Exception:
                # 回滚内存值，关闭时还能再尝试一次，也避免用假值误判“已保存”。
                self.vars["window_geometry"] = previous
                raise
            return True
        except Exception:
            get_logger().exception("保存窗口几何失败（不影响窗口关闭）")
            return False

    def _on_window_closing(self) -> bool:
        """打印中关闭窗口前的确认。

        pywebview 6.x 的 closing 事件按 `should_cancel = closing.set()` 处理：
        任一处理器返回 False 即取消关闭，因此返回 True 才表示允许关闭。
        该回调运行在 GUI 线程，必须立即返回，不能等待工作线程，否则窗口会卡死；
        用户确认取消后由 cancel_print 通知工作线程退出，随后放行关闭。
        """
        # 关闭前再保存一次最终几何：只做 Win32 读取与设置落盘（不投递 JS），
        # 最大化/最小化时保存还原态；失败仅记录日志，绝不阻止窗口关闭。
        self.save_window_geometry()
        if not self._printing:
            return True
        try:
            proceed = self._window.create_confirmation_dialog(
                "取消打印并退出",
                "打印任务仍在进行。\n\n"
                "退出会立即取消当前打印任务，已送出的纸张无法收回。\n"
                "确定要取消打印并退出吗？")
        except Exception:
            get_logger().exception("关闭窗口确认对话框失败")
            return True  # 对话框不可用时允许退出，避免窗口无法关闭
        if proceed is None:
            get_logger().warning("无法显示关闭确认对话框，已阻止关闭以保护打印任务。")
            return False
        if not proceed:
            return False
        # 已确认退出：先屏蔽 JS 投递（避免 GUI 线程 evaluate_js 自锁），再取消任务。
        self._suppress_js = True
        self.cancel_print()
        return True

    def _on_window_closed(self, *_args) -> None:
        """停止拖放轮询并唤醒等待中的任务，确保关闭窗口后进程能够退出。"""
        self._suppress_js = True
        self._geometry_closed = True
        self._shutdown_event.set()
        self.stop_event.set()
        self.pause_event.set()
        self._flip_event.set()
        self._phase_confirm_event.set()
        self._executor.shutdown(wait=False, cancel_futures=True)

    def _drop_pump(self) -> None:
        """轮询 pywebview 原生拖放通道，拿到真实文件路径后入队。"""
        from webview.dom import _dnd_state
        if os.environ.get("WV_SELFTEST") == "1":
            print(f"DROP_CHANNEL num_listeners={_dnd_state['num_listeners']}", flush=True)
        seen = 0
        while not self._shutdown_event.wait(0.25):
            try:
                paths = list(_dnd_state.get("paths", []))
                if len(paths) > seen:
                    new_items = paths[seen:]
                    seen = len(paths)
                    real = [p for _name, p in new_items if os.path.isfile(p)]
                    if real:
                        self._log(f"拖入 {len(real)} 个文件")
                        self._executor.submit(self._add_dropped, real)
            except Exception:
                pass

    def _add_dropped(self, paths: List[str]) -> None:
        res = self.add_files(paths)
        self._post("files-added", {"added": res.get("added", []),
                                   "total": len(self.files)})

    def _apply_window_icon(self) -> None:
        """给当前 WebView2 顶层窗口设置完整 Win32 图标，不按标题查找窗口。"""
        try:
            from app_identity import apply_native_window_icon

            if not self._window.events.shown.wait(10):
                raise RuntimeError("等待窗口显示超时")
            native = self._window.native
            hwnd = int(native.Handle.ToInt64())
            icon_path = os.path.join(os.path.dirname(os.path.dirname(
                os.path.abspath(__file__))), "app_icon.ico")
            handles = apply_native_window_icon(hwnd, icon_path)
            if set(handles) != {"small", "big"}:
                raise RuntimeError("应用图标加载不完整")
            self._window_icon_handles = handles
        except Exception:
            get_logger().exception("设置 WebView2 窗口图标失败")

    def _selftest(self) -> None:
        import time
        time.sleep(4.0)
        checks = {}
        try:
            # 在最小支持尺寸下验证按钮、标题指标和整个页面均不横向溢出。
            self._window.resize(1024, 680)
            time.sleep(0.5)
            layout = self._window.evaluate_js(
                "JSON.stringify({metrics:document.querySelectorAll('.metric').length,"
                "brandIcon:(()=>{const i=document.querySelector('.brand-icon');"
                "return i && i.complete && i.naturalWidth>0})(),"
                "topbar:document.querySelector('.topbar').offsetHeight,"
                "workspace:document.querySelector('.workspace').offsetHeight,"
                "printbar:document.querySelector('.printbar').offsetHeight,"
                "logpanel:document.querySelector('.logpanel').offsetHeight,"
                "vh:window.innerHeight,vw:window.innerWidth,"
                "pageOverflow:document.documentElement.scrollWidth>window.innerWidth,"
                "buttonsOkay:Array.from(document.querySelectorAll('.btn'))"
                ".filter(b=>b.offsetParent!==null).every(b=>{const r=b.getBoundingClientRect();"
                "return r.left>=0 && r.right<=window.innerWidth && r.top>=0 &&"
                "r.bottom<=window.innerHeight && r.height>=28}),"
                "fieldsOkay:Array.from(document.querySelectorAll('.field,.radio'))"
                ".filter(e=>e.offsetParent!==null).every(e=>e.scrollWidth<=e.clientWidth+1),"
                "fileButtonsSameRow:(()=>{const b=['btn-add','btn-add-dir','btn-remove','btn-clear']"
                ".map(id=>document.getElementById(id).getBoundingClientRect());"
                "return b.every(r=>r.top===b[0].top)})(),"
                "selectionBarHidden:(document.getElementById('selection-bar')||{}).hidden===true,"
                "emptyBoxFits:(()=>{const list=document.getElementById('file-list');"
                "const box=document.getElementById('empty-state');"
                "if(!list||!box) return false;"
                "const lr=list.getBoundingClientRect(),br=box.getBoundingClientRect();"
                "return br.top>=lr.top-0.5 && br.bottom<=lr.bottom+0.5"
                " && box.scrollHeight<=box.clientHeight+1 && br.height>60})(),"
                "printerSelectWidth:Math.round("
                "document.getElementById('printer-select').getBoundingClientRect().width),"
                "plusStroke:(()=>{const p=document.querySelector('#btn-empty-add svg path');"
                "if(!p) return false;"
                "const d=(p.getAttribute('d')||'').replace(/\\s+/g,'');"
                "return d==='M125v14M512h14'"
                " && Math.abs(parseFloat(p.getAttribute('stroke-width'))-3.4)<0.001"
                " && p.getAttribute('stroke-linecap')==='round'"
                " && (p.getAttribute('stroke')||'').toLowerCase()==='#fff'})(),"
                "jsRan:!!window.__event,"
                "printers:(document.getElementById('printer-select')||{}).length,"
                "advBtn:!!document.getElementById('btn-advanced'),"
                "pauseBtn:!!document.getElementById('btn-pause'),"
                "cancelBtn:!!document.getElementById('btn-cancel'),"
                "planBtn:!!document.getElementById('btn-plan-detail'),"
                "progressBar:!!document.getElementById('progress-track')})")
            print("SELFTEST_LAYOUT:", layout, flush=True)
            layout_data = json.loads(layout)
            printer_width = int(layout_data.get("printerSelectWidth") or 0)
            checks["layout"] = (
                layout_data.get("metrics") == 4 and layout_data.get("brandIcon") is True
                and layout_data.get("pageOverflow") is False
                and layout_data.get("buttonsOkay") is True
                and layout_data.get("fieldsOkay") is True
                and layout_data.get("fileButtonsSameRow") is True
            )
            checks["printer_width"] = 190 <= printer_width <= 245
            checks["plus_icon_style"] = bool(layout_data.get("plusStroke"))
            checks["empty_selection_bar_hidden"] = bool(layout_data.get("selectionBarHidden"))
            checks["empty_state_fits"] = bool(layout_data.get("emptyBoxFits"))
            # 窗口几何记忆：真实 resize 后按 DPI 换算保存逻辑像素几何，
            # 150% DPI 下必须仍是 1024x680（而不是 1536x1020）。
            checks["geometry_save"] = bool(self.save_window_geometry())
            parsed_geometry = window_geometry.parse_geometry(
                str(self.vars.get("window_geometry") or ""))
            print("SELFTEST_GEOMETRY:", self.vars.get("window_geometry"), flush=True)
            checks["geometry_parse"] = parsed_geometry is not None
            checks["geometry_logical_size"] = bool(
                parsed_geometry is not None
                and abs(parsed_geometry[0] - 1024) <= 4
                and abs(parsed_geometry[1] - 680) <= 4)
            from app_identity import APP_USER_MODEL_ID, get_process_app_user_model_id
            checks["aumid"] = get_process_app_user_model_id() == APP_USER_MODEL_ID
            icon_handles = getattr(self, "_window_icon_handles", {})
            if icon_handles:
                import ctypes
                user32 = ctypes.windll.user32
                user32.SendMessageW.restype = ctypes.c_ssize_t
                hwnd = int(self._window.native.Handle.ToInt64())
                big = user32.SendMessageW(hwnd, 0x007F, 1, 0)
                small2 = user32.SendMessageW(hwnd, 0x007F, 2, 0)
                checks["window_icon"] = (
                    int(big) == int(icon_handles["big"])
                    and int(small2) == int(icon_handles["small"])
                )
            else:
                checks["window_icon"] = False
            checks["settings"] = self.get_settings().get("ok")
            checks["printers"] = self.list_printers().get("ok")
            checks["converters"] = self.converters_available().get("ok")
            checks["recent"] = self.get_recent().get("ok")
            checks["cleanup"] = self.cleanup_stale().get("ok")
            # 通过 WebView2 的真实 on_script_notify 分支注入 CoreWebView2File，
            # 不再直接往 _dnd_state['paths'] 写数据来冒充拖放成功。
            import tempfile as tf_mod
            from webview.platforms.edgechromium import EdgeChrome
            temp_dir = tf_mod.mkdtemp(prefix="wv-drop-")
            drop_paths = []
            for index in (1, 2):
                drop_path = os.path.join(temp_dir, f"dropped-{index}.pdf")
                doc = fitz.open()
                # 4 页：让“全部/奇数页/偶数页/自定义”范围标签都能被真实校验，
                # 1 页文件无法区分这些文案（都会被收敛为“全部”）。
                for page_no in range(1, 5):
                    page = doc.new_page(width=595, height=842)
                    page.insert_text((72, 72), f"DROPPED_FILE_{index}_P{page_no}")
                doc.save(drop_path)
                doc.close()
                drop_paths.append(drop_path)

            class CoreWebView2File:
                def __init__(self, path):
                    self.Path = path

            class DropArgs:
                def get_WebMessageAsJson(self):
                    return '"FilesDropped"'

                def get_AdditionalObjects(self):
                    return [CoreWebView2File(path) for path in drop_paths]

            def js(code):
                return self._window.evaluate_js(code)

            # 原生弹窗探针：若任何交互仍走系统 confirm，计数会 > 0。
            js("window.__confirmCalls = 0;"
               "window.confirm = () => { window.__confirmCalls += 1; return true; }")

            # 空状态圆形加号必须是真实可点击按钮、几何完整、不被选中操作条遮挡，
            # 且中心点命中的就是按钮（或按钮内 SVG）。先断言空状态再投递拖放，
            # 避免拖放线程提前写入文件，把“空状态”断言变成竞态。
            checks["empty_add_button"] = js(
                "(()=>{const b=document.getElementById('btn-empty-add');"
                "if(!b) return false;"
                "const r=b.getBoundingClientRect();"
                "const list=document.getElementById('file-list').getBoundingClientRect();"
                "const bar=document.getElementById('selection-bar');"
                "const visible=Math.max(0,Math.min(r.bottom,list.bottom)"
                "-Math.max(r.top,list.top));"
                "const ratio=r.height>0?visible/r.height:0;"
                "const hit=document.elementFromPoint(r.left+r.width/2,r.top+r.height/2);"
                "return b.tagName==='BUTTON' && b.offsetParent!==null"
                " && Math.round(r.width)===30 && Math.round(r.height)===30"
                " && typeof b.onclick==='function' && b.onclick.name==='addFiles'"
                " && !!bar && bar.hidden===true && ratio>=0.9"
                " && (hit===b || b.contains(hit))})()")

            # 纸张方向选项只保留“自动/竖版/横版”，不再带括号说明。
            checks["orientation_labels_short"] = js(
                "(()=>{const o=document.getElementById('preview-orientation');"
                "return !!o && [...o.options].map(x=>x.textContent).join('|')"
                " === '自动|竖版|横版'})()")

            # 拖入高亮只能一层：空状态时只给 #empty-state 实线框+浅蓝底，
            # #file-list 不得再出现第二圈虚线（outline 必须为空）。
            checks["drop_hint_empty_single_layer"] = js(
                "(()=>{const ev=new Event('dragenter',{bubbles:true});"
                "ev.dataTransfer={types:['Files']};document.dispatchEvent(ev);"
                "const list=document.getElementById('file-list');"
                "const empty=document.getElementById('empty-state');"
                "const cs=getComputedStyle(empty);"
                "const ok=empty.classList.contains('drag-over')"
                " && !list.classList.contains('drag-over')"
                " && cs.borderStyle==='solid' && !list.style.outline;"
                "const lv=new Event('dragleave');lv.relatedTarget=null;"
                "document.dispatchEvent(lv);"
                "return ok && !empty.classList.contains('drag-over')"
                " && !list.classList.contains('drag-over')})()")

            EdgeChrome.on_script_notify(None, None, DropArgs())

            # 任务结束汇总事件必须弹出 modal-summary，并逐行写入中文信息。
            js("window.__event('summary',{title:'打印完成',"
               "message:'共 1 个文件、约 1 张纸的打印任务已提交完成。\\n请核对页序。'})")
            checks["summary_modal"] = js(
                "document.getElementById('modal-summary').style.display==='flex' && "
                "document.getElementById('summary-title').textContent==='打印完成' && "
                "document.getElementById('summary-body').children.length===2")
            js("hideModal('modal-summary')")

            # 纸张不一致确认复用翻面模态框，标题与继续按钮文案必须正确。
            js("window.__event('flip-modal',{title:'纸张尺寸/方向不一致',"
               "detail:['检测到纸张不一致'],sample:'仍要继续整批打印吗？',"
               "proceed_label:'仍要继续整批打印'})")
            checks["paper_confirm_modal"] = js(
                "document.getElementById('modal-flip').style.display==='flex' && "
                "document.getElementById('flip-title').textContent==='纸张尺寸/方向不一致' && "
                "document.getElementById('btn-flip-confirm').textContent==='仍要继续整批打印'")
            js("hideModal('modal-flip')")

            def wait_js(code, timeout=4.0):
                deadline = time.monotonic() + timeout
                while time.monotonic() < deadline:
                    if js(code):
                        return True
                    time.sleep(0.1)
                return False

            checks["drop_pump"] = wait_js("state.files.length === 2")
            paths = sorted(rec["path"] for rec in self.files.values())
            checks["drop_path_correct"] = paths == sorted(drop_paths)
            checks["drop_preview_ready"] = wait_js(
                "state.previewFile && state.previewFile.iid === state.files[1].iid && "
                "!!document.querySelector('#preview-stage img')")
            # 有文件时只高亮 #file-list 内侧，空状态不得再叠加虚线框。
            checks["drop_hint_files_single_layer"] = js(
                "(()=>{const ev=new Event('dragenter',{bubbles:true});"
                "ev.dataTransfer={types:['Files']};document.dispatchEvent(ev);"
                "const list=document.getElementById('file-list');"
                "const empty=document.getElementById('empty-state');"
                "const ok=list.classList.contains('drag-over')"
                " && !empty.classList.contains('drag-over')"
                " && !list.style.outline;"
                "const lv=new Event('dragleave');lv.relatedTarget=null;"
                "document.dispatchEvent(lv);"
                "return ok && !list.classList.contains('drag-over')"
                " && !empty.classList.contains('drag-over')})()")
            # 添加 / 拖放的 2 个文件必须默认全部选中；取消勾选一个后只剩 1 个。
            checks["auto_select_added"] = js(
                "state.selected.size === 2 && "
                "state.files.every(f => state.selected.has(f.iid))")
            checks["deselect_one_click"] = js(
                "(()=>{const cb=document.querySelector('.file-row input[type=checkbox]');"
                "if(!cb) return false;"
                "cb.onclick({stopPropagation(){}});"
                "return state.selected.size === 1 && !state.selected.has(state.files[0].iid)"
                " && document.querySelectorAll('.file-row.selected').length === 1})()")
            # 行单击只切换预览，不改变选中；已选中行再点也不能取消选中。
            js("document.querySelectorAll('.file-row')[0].onclick()")
            checks["row_click_does_not_select"] = wait_js(
                "state.selected.size === 1 && !state.selected.has(state.files[0].iid) && "
                "state.previewFile && state.previewFile.iid === state.files[0].iid")
            js("state.selected = new Set(state.files.map(f => f.iid)); renderFiles()")
            js("document.querySelectorAll('.file-row')[1].onclick()")
            checks["row_click_keeps_selection"] = wait_js(
                "state.selected.size === 2 && "
                "state.files.every(f => state.selected.has(f.iid)) && "
                "state.previewFile && state.previewFile.iid === state.files[1].iid")
            # 有文件时选中操作条必须重新显示，且 4 个按钮仍是单横排、在条内不溢出。
            checks["selection_buttons_single_row"] = js(
                "(()=>{const bar=document.getElementById('selection-bar');"
                "if(!bar || bar.hidden) return false;"
                "const bs=[...bar.querySelectorAll('.selection-buttons .btn')]"
                ".map(b=>b.getBoundingClientRect());"
                "if(bs.length!==4) return false;"
                "const barRect=bar.getBoundingClientRect();"
                "return bs.every(r=>Math.abs(r.top-bs[0].top)<0.5 && r.left>=barRect.left-0.5"
                " && r.right<=barRect.right+0.5 && r.left>=0 && r.right<=window.innerWidth"
                " && Math.round(r.height)>=28)})()")
            # 输出单选项顺序：手动双面在前、单面在后，默认仍为手动双面。
            checks["duplex_option_order"] = js(
                "(()=>{const rs=[...document.querySelectorAll('input[name=duplex]')];"
                "return rs.length===2 && rs[0].value==='manual' && rs[1].value==='simplex'"
                " && rs[0].checked===true})()")
            # 四个选中操作按钮必须等宽铺满操作条（不只是同一行）。
            checks["selection_buttons_equal_width"] = js(
                "(()=>{const bar=document.getElementById('selection-bar');"
                "const btns=[...bar.querySelectorAll('.selection-buttons .btn')];"
                "if(btns.length!==4) return false;"
                "const ws=btns.map(b=>b.getBoundingClientRect().width);"
                "const inner=bar.clientWidth-20;"
                "const rects=btns.map(b=>b.getBoundingClientRect());"
                "const span=rects[3].right-rects[0].left;"
                "return Math.max(...ws)-Math.min(...ws)<=1"
                " && Math.abs(span-inner)<=2"
                " && btns.every(b=>b.scrollWidth<=b.clientWidth+1)})()")

            # 文件行显示“当前实际生效范围”，并随全局范围设置即时刷新。
            checks["range_label_default_all"] = js(
                "[...document.querySelectorAll('.file-row .pages')]"
                ".every(el=>el.textContent.includes('全部'))")
            # #range-text 在非自定义模式只显示说明文案；custom/recent 才显示原文。
            checks["range_text_title"] = js(
                "document.getElementById('range-text').title"
                " === '全局范围仅对未单独设置的文件生效；各文件可独立设置范围与份数'")
            checks["range_text_default_all"] = js(
                "document.getElementById('range-text').value === '全部'")
            js("document.getElementById('range-mode').value = 'odd'; onSettingChange()")
            checks["range_label_follows_odd"] = wait_js(
                "(()=>{const rows=[...document.querySelectorAll('.file-row .pages')];"
                "return rows.length===2 && rows.every(r=>r.textContent.includes('奇数页'))})()")
            checks["range_text_follows_odd"] = wait_js(
                "document.getElementById('range-text').value === '仅奇数页'")
            js("document.getElementById('range-mode').value = 'even'; onSettingChange()")
            checks["range_text_follows_even"] = wait_js(
                "document.getElementById('range-text').value === '仅偶数页'")
            js("document.getElementById('range-mode').value = 'current'; onSettingChange()")
            checks["range_text_follows_current"] = wait_js(
                "document.getElementById('range-text').value === '当前页'")
            js("document.getElementById('range-mode').value = 'custom';"
               "document.getElementById('range-text').value = '1-2'; onSettingChange()")
            checks["range_label_follows_custom"] = wait_js(
                "(()=>{const rows=[...document.querySelectorAll('.file-row .pages')];"
                "return rows.length===2 && rows.every(r=>r.textContent.includes('1-2'))})()")
            checks["range_text_follows_custom"] = wait_js(
                "document.getElementById('range-text').value === '1-2'")
            js("document.getElementById('range-mode').value = 'all'; onSettingChange()")
            checks["range_label_back_to_all"] = wait_js(
                "[...document.querySelectorAll('.file-row .pages')]"
                ".every(el=>el.textContent.includes('全部'))")
            checks["range_text_back_to_all"] = js(
                "document.getElementById('range-text').value === '全部'")
            # 切回指定范围必须恢复用户原文，不能被“全部”提示文案覆盖。
            js("document.getElementById('range-mode').value = 'custom'; onSettingChange()")
            checks["range_text_custom_restored"] = wait_js(
                "document.getElementById('range-text').value === '1-2'")
            js("document.getElementById('range-mode').value = 'all'; onSettingChange()")

            # 范围无效时“开始打印”必须真的呈灰色禁用态：它是 accent 主按钮，
            # disabled 变体必须靠层叠顺序覆盖 --accent，否则用户会误以为仍可打印。
            js("document.getElementById('range-mode').value = 'custom';"
               "document.getElementById('range-text').value = '9-10'; onSettingChange()")
            checks["print_disabled_on_range_error"] = wait_js(
                "document.getElementById('btn-print').disabled === true && "
                "state.files.length > 0 && state.files.some((f) => !!f.range_error)")
            # 背景/边框带 .12s 过渡：等过渡结束再取计算样式，避免读到蓝→灰的中间帧。
            checks["print_disabled_grey_settled"] = wait_js(
                "getComputedStyle(document.getElementById('btn-print'))"
                ".backgroundColor === 'rgb(238, 240, 243)'", timeout=3.0)
            time.sleep(0.1)
            print_button = {}
            try:
                raw_style = js(
                    "(() => { const el = document.getElementById('btn-print');"
                    "const cs = getComputedStyle(el);"
                    "return JSON.stringify({disabled: el.disabled, bg: cs.backgroundColor,"
                    "color: cs.color, opacity: cs.opacity, border: cs.borderColor}); })()")
                if isinstance(raw_style, str):
                    print_button = json.loads(raw_style)
            except (TypeError, ValueError):
                print_button = {}
            print("SELFTEST_PRINT_DISABLED_STYLE:", json.dumps(print_button), flush=True)
            checks["print_disabled_not_accent_blue"] = (
                print_button.get("bg") != "rgb(37, 99, 235)"
                and print_button.get("color") != "rgb(255, 255, 255)")
            checks["print_disabled_uses_grey_palette"] = (
                print_button.get("bg") == "rgb(238, 240, 243)"
                and print_button.get("color") == "rgb(98, 109, 123)"
                and print_button.get("opacity") == "1")
            js("document.getElementById('range-mode').value = 'all'; onSettingChange()")
            checks["print_enabled_after_range_fixed"] = wait_js(
                "document.getElementById('btn-print').disabled === false && "
                "state.files.every((f) => !f.range_error)")

            # 打印机失败统一显示“（脱机）”，并用不存在的名称触发真实失败路径。
            js("void updatePrinterStatus('SELFTEST-OFFLINE-01')")
            checks["printer_offline_text"] = wait_js(
                "document.getElementById('printer-name').textContent.endsWith('（脱机）')"
                " && document.getElementById('printer-dot').className.indexOf('bad') >= 0")
            js("void updatePrinterStatus(state.settings.printer)")

            checks["recent_select"] = wait_js(
                "document.getElementById('recent-select').options.length >= 2 && "
                "!document.getElementById('btn-add-recent').disabled")

            # 真实调用文件行 onclick，必须切换当前预览并更新图片。
            second_preview = js(
                "(document.querySelector('#preview-stage img')||{}).src || ''")
            js("document.querySelectorAll('.file-row')[0].onclick()")
            checks["preview_switches"] = wait_js(
                "state.previewFile && state.previewFile.iid === state.files[0].iid && "
                "!!document.querySelector('#preview-stage img')")
            first_preview = js(
                "(document.querySelector('#preview-stage img')||{}).src || ''")
            checks["preview_image_changes"] = bool(
                first_preview and second_preview and first_preview != second_preview)

            # 移除当前预览文件：只选中它，避免多选触发确认框（此处专测预览回退）。
            js("state.selected = new Set([state.previewFile.iid]); renderFiles()")
            js("document.getElementById('btn-remove').click()")
            checks["preview_after_remove"] = wait_js(
                "state.files.length === 1 && state.previewFile && "
                "state.previewFile.iid === state.files[0].iid")

            # Esc 只关闭确认框、不执行清空（防止键盘误删）。
            js("document.getElementById('btn-clear').click()")
            checks["confirm_modal_clear"] = wait_js(
                "document.getElementById('modal-confirm').style.display === 'flex'")
            js("document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape'}))")
            checks["confirm_escape_cancels"] = wait_js(
                "document.getElementById('modal-confirm').style.display === 'none'"
                " && state.files.length === 1")

            # 再次点击清空 → 应用内确认框 → 确认后才真正清空。
            js("document.getElementById('btn-clear').click()")
            wait_js("document.getElementById('modal-confirm').style.display === 'flex'")
            js("answerConfirm(true)")
            checks["preview_after_clear"] = wait_js(
                "state.files.length === 0 && state.previewFile === null && "
                "!document.querySelector('#preview-stage img') && "
                "!!document.querySelector('#preview-stage .preview-empty')")

            # 最近记录在清空后仍可重新添加，并复用同一套文件/预览状态链。
            js("document.getElementById('btn-add-recent').click()")
            checks["recent_add"] = wait_js(
                "state.files.length === 1 && state.previewFile !== null")

            js("document.querySelector('input[name=preview-view][value=spread]').checked = true;"
               "onSettingChange()")
            checks["preview_spread"] = wait_js(
                "document.querySelectorAll('#preview-stage.spread img').length === 2")
            js("document.getElementById('preview-nup').value = '2'; onSettingChange()")
            checks["preview_nup"] = wait_js(
                "document.getElementById('preview-status').textContent.includes('2 合 1')")
            # 横版 + 2 合 1：纸张必须真的变成 297×210（宽>高），预览图片也是横向，
            # 说明方向交换、槽位比例、预览与真实合成三者一致。
            js("document.getElementById('preview-orientation').value = 'landscape';"
               "onSettingChange()")
            checks["preview_landscape"] = wait_js(
                "(()=>{const s=document.getElementById('preview-status').textContent;"
                "const img=document.querySelector('#preview-stage img');"
                "return s.includes('297×210 mm') && s.includes('2 合 1') && !!img"
                " && img.naturalWidth>0 && img.naturalWidth>img.naturalHeight})()", timeout=6.0)
            time.sleep(0.4)
            checks["settings_persist"] = (
                self.vars.get("nup") == 2 and self.vars.get("preview_view") == "spread")

            # B：1 页文件使用 “1,3,5,…,31” 这类含越界页码的请求时，标签只能显示
            # 有效页（收敛为“全部”），绝不能回显越界页码。
            overflow_text = "1,3,5,7,9,11,13,15,17,19,21,23,25,27,29,31"
            valid_one, dropped_one = resolve_range_pages("custom", overflow_text, 1)
            label_one = self._range_label(
                "custom", overflow_text, valid=valid_one, total_pages=1)
            print("SELFTEST_RANGE_LABEL:", label_one, flush=True)
            checks["range_label_single_page_no_overflow"] = (
                label_one == "全部" and bool(dropped_one)
                and not any(str(p) in label_one for p in dropped_one))

            # 追加第 2 个文件，供“更多”菜单与“应用对象”校验使用。
            extra_path = os.path.join(temp_dir, "dropped-extra.pdf")
            doc = fitz.open()
            for page_no in range(1, 5):
                page = doc.new_page(width=595, height=842)
                page.insert_text((72, 72), f"DROPPED_EXTRA_P{page_no}")
            doc.save(extra_path)
            doc.close()
            drop_paths.append(extra_path)
            extra_added = self.add_files([extra_path]).get("added", [])
            js(f"window.__event('files-added', {json.dumps({'added': extra_added, 'total': len(self.files)}, ensure_ascii=False)})")
            checks["extra_file_added"] = wait_js("state.files.length === 2")

            # A：上移/下移只移动当前预览文件；默认全选时两个方向也不会互相阻挡。
            js("window.__firstBefore = state.files[0].iid;"
               "window.__secondBefore = state.files[1].iid;"
               "state.selected = new Set(state.files.map(f => f.iid)); renderFiles();"
               "selectPreviewFile(state.files[1])")
            time.sleep(0.2)
            js("document.getElementById('btn-move-up').click()")
            checks["more_move_up_single_file"] = wait_js(
                "state.files[0].iid === window.__secondBefore && "
                "state.files[1].iid === window.__firstBefore && "
                "state.selected.size === 2 && "
                "state.previewFile && state.previewFile.iid === window.__secondBefore")
            js("document.getElementById('btn-move-down').click()")
            checks["more_move_down_single_file"] = wait_js(
                "state.files[0].iid === window.__firstBefore && "
                "state.files[1].iid === window.__secondBefore && state.selected.size === 2")

            # A：去重无重复时必须写可读日志，而不是静默无反应。
            js("document.getElementById('log-body').innerHTML = ''")
            js("document.getElementById('btn-dedupe').click()")
            checks["dedupe_no_duplicate_feedback"] = wait_js(
                "document.getElementById('log-body').innerText.includes('未发现重复文件')")

            # C：单独设置默认作用对象 = 当前文件，另一文件必须保持不变。
            js("state.selected = new Set(state.files.map(f => f.iid)); renderFiles();"
               "selectPreviewFile(state.files[1]);"
               "openFileOpt(state.files[1].iid)")
            checks["fileopt_scope_default_current"] = js(
                "document.getElementById('fileopt-scope-current').checked === true"
                " && document.getElementById('fileopt-title').textContent.includes('将应用到')"
                " && document.getElementById('fileopt-preview').textContent.includes('将应用到')")
            js("document.getElementById('fileopt-range').value = '2';"
               "document.getElementById('fileopt-copies').value = '3'; applyFileOpt()")
            checks["fileopt_current_only"] = wait_js(
                "state.files[1].range_text === '2' && state.files[1].copies === 3 && "
                "!state.files[0].range_text && (state.files[0].copies || 1) === 1")

            # C：scope=selected 命中全部已选文件。
            js("state.selected = new Set(state.files.map(f => f.iid)); renderFiles();"
               "openFileOpt(state.files[0].iid)")
            checks["fileopt_scope_count"] = js(
                "document.getElementById('fileopt-scope-count').textContent === '2'")
            js("setFileOptScope('selected')")
            checks["fileopt_scope_selected_label"] = js(
                "document.getElementById('fileopt-scope-selected').checked === true"
                " && document.getElementById('fileopt-preview').textContent"
                ".includes('已选的 2 个文件')")
            js("document.getElementById('fileopt-range').value = '1-2';"
               "document.getElementById('fileopt-copies').value = '5'; applyFileOpt()")
            checks["fileopt_scope_selected"] = wait_js(
                "state.files.length === 2 && "
                "state.files.every(f => f.range_text === '1-2' && f.copies === 5)")

            # C：“应用到全部”按钮必须预选“全部文件”并作用于全部。
            js("state.selected = new Set([state.files[0].iid]); renderFiles();"
               "document.getElementById('btn-file-apply-all').click()")
            checks["fileopt_apply_all_preselects_all"] = js(
                "document.getElementById('fileopt-scope-all').checked === true"
                " && document.getElementById('fileopt-preview').textContent"
                ".includes('全部 2 个文件')")
            js("document.getElementById('fileopt-range').value = '';"
               "document.getElementById('fileopt-copies').value = '2'; applyFileOpt()")
            checks["fileopt_scope_all"] = wait_js(
                "state.files.every(f => !f.range_text && f.copies === 2)")

            js("openFileOpt(state.files[0].iid);"
               "document.getElementById('fileopt-range').value = '1';"
               "document.getElementById('fileopt-copies').value = '2';"
               "applyFileOpt()")
            checks["file_options"] = wait_js(
                "state.files[0].range_text === '1' && state.files[0].copies === 2 && "
                "state.advanced.per_file === true")
            # 有逐文件自定义范围时，全局“全部”只作用于未单独设置的文件，
            # 输入框显示“各文件独立”（不再写“逐文件（见左侧）”）。
            checks["range_text_per_file_hint"] = wait_js(
                "document.getElementById('range-mode').value === 'all' && "
                "document.getElementById('range-text').value === '各文件独立'")
            js("saveFileOptions({applyAll:true})")
            checks["file_apply_all"] = wait_js(
                "state.files.every(f => f.range_text === '1' && f.copies === 2)")
            js("saveFileOptions({inherit:true, applyAll:true})")
            checks["file_inherit"] = wait_js(
                "state.files.every(f => !f.range_text && f.copies === null) && "
                "state.advanced.per_file === false")
            # 清除逐文件范围后，全局“全部”重新显示为“全部”。
            checks["range_text_back_to_all_after_inherit"] = wait_js(
                "document.getElementById('range-mode').value === 'all' && "
                "document.getElementById('range-text').value === '全部'")

            js("openAdvanced(); advancedDraft.nup = 4; hideModal('modal-advanced')")
            checks["advanced_cancel"] = js("state.advanced.nup === 2")
            js("openAdvanced(); document.getElementById('adv-nup').value = '4'; applyAdvanced()")
            checks["advanced_apply"] = wait_js(
                "state.advanced.nup === 4 && document.getElementById('preview-nup').value === '4'")
            # “恢复 Acrobat 默认”必须走后端 restore_defaults(acrobat=True)：
            # 单面、份数 1、范围 all、高级版式全部回基线，并自动关闭弹窗。
            js("openAdvanced()")
            js("void restoreAcrobatDefaults()")
            checks["acrobat_defaults"] = wait_js(
                "document.getElementById('modal-advanced').style.display === 'none' && "
                "document.getElementById('advanced-summary').textContent === '高级设置：Acrobat 默认' && "
                "document.querySelector('input[name=duplex]:checked').value === 'simplex' && "
                "collectOptions().manual_duplex === false && "
                "collectOptions().print_flow === 'batch_all' && "
                "document.getElementById('copies').value === '1' && "
                "document.getElementById('range-mode').value === 'all' && "
                "state.advanced.nup === 1 && state.advanced.annots === false", timeout=6.0)
            js("toggleLogDetail()")
            checks["log_thumbnails"] = wait_js(
                "!document.getElementById('log-detail').hidden && "
                "document.getElementById('thumb-strip').children.length > 0")
            # 先把设置改脏，确保“恢复默认”确实走了回写而不是本来就等于默认值。
            js("document.getElementById('preview-nup').value = '4';"
               "document.getElementById('copies').value = '2'; onSettingChange()")
            checks["restore_needs_dirty_state"] = js(
                "state.advanced.nup === 4 && document.getElementById('copies').value === '2'")
            js("document.getElementById('btn-restore-defaults').click()")
            checks["confirm_modal_restore"] = wait_js(
                "document.getElementById('modal-confirm').style.display === 'flex'")
            js("answerConfirm(true)")
            checks["restore_defaults"] = wait_js(
                "state.advanced.nup === 1 && document.getElementById('copies').value === '1'")
            checks["restore_keeps_recent"] = js(
                "document.getElementById('recent-select').options.length >= 2")
            # 全程（清空/移除/恢复默认）都不得再走系统原生确认框。
            checks["no_native_confirm"] = js("window.__confirmCalls === 0")

            self.clear_files()
            js("clearPreview(); renderFiles()")
            for drop_path in drop_paths:
                os.unlink(drop_path)
            os.rmdir(temp_dir)
            print("SELFTEST_API:", json.dumps(checks), flush=True)
            print("SELFTEST_RESULT:", "PASS" if all(checks.values()) else "FAIL", flush=True)
        except Exception:
            print("SELFTEST_ERROR:\n" + traceback.format_exc(), flush=True)
        if os.environ.get("WV_SELFTEST_CLOSE") == "1":
            self._window.destroy()

    def _js(self, code: str) -> None:
        if self._window is not None and not self._suppress_js:
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
                # 与 settings_store / pdf_layout 契约一致：自定义百分比即 custom。
                {"value": "custom", "label": "自定义缩放"},
                {"value": "actual", "label": "实际大小 100%"},
            ],
            "orientations": [
                {"value": "auto", "label": "自动"},
                {"value": "portrait", "label": "竖版"},
                {"value": "landscape", "label": "横版"},
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
        changes = dict(changes or {})
        if "delay" in changes and "delay_seconds" not in changes:
            changes["delay_seconds"] = changes.pop("delay")
        self.vars.update(changes)
        try:
            settings_store.save_settings(self.vars)
            return _ok()
        except Exception as exc:
            return _err(exc)

    def restore_defaults(self, acrobat: bool = False) -> Dict[str, Any]:
        """恢复应用默认或 Acrobat 等效默认，保留最近记录与窗口几何。"""
        recent = self.vars.get("recent_files") or []
        ranges = self.vars.get("recent_ranges") or []
        geometry = self.vars.get("window_geometry") or ""
        self.vars = settings_store.acrobat_defaults() if acrobat else settings_store.default_settings()
        self.vars["recent_files"] = recent
        self.vars["recent_ranges"] = ranges
        self.vars["window_geometry"] = geometry
        try:
            settings_store.save_settings(self.vars)
            return _ok(settings=self.vars)
        except Exception as exc:
            return _err(exc)

    def choose_output_directory(self) -> Dict[str, Any]:
        try:
            import webview
            if self._window is None:
                raise RuntimeError("窗口尚未就绪")
            result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
            paths = self._dialog_result_paths(result)
            return _ok(path=paths[0] if paths else "")
        except Exception as exc:
            return _err(exc)

    def clear_recent(self) -> Dict[str, Any]:
        self.vars["recent_files"] = []
        try:
            settings_store.save_settings(self.vars)
        except Exception as exc:
            return _err(exc)
        self._post("recent-updated", [])
        return _ok()

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
            return _ok(status={"name": name, "state": "脱机"}, ready=False,
                       detail=_describe_error(
                           exc, "请检查打印机电源、USB/网络连接与纸张后重试。"))

    def paper_sources(self, printer: str) -> Dict[str, Any]:
        try:
            sources = get_printer_sources(printer)
            self.vars["_paper_sources"] = sources
            return _ok(sources=sources)
        except Exception as exc:
            return _ok(sources=[], error=str(exc))

    # ============================================================ 文件
    @staticmethod
    def _dialog_result_paths(result) -> List[str]:
        if not result:
            return []
        if isinstance(result, (str, os.PathLike)):
            return [os.fspath(result)]
        return [os.fspath(path) for path in result]

    def choose_files(self) -> Dict[str, Any]:
        """使用 Windows 原生多选对话框，直接返回真实本地路径。

        WebView2 的 File.path 按浏览器安全规范恒为空，不能用于后端读文件。
        """
        try:
            import webview
            if self._window is None:
                raise RuntimeError("窗口尚未就绪，请稍后重试")
            result = self._window.create_file_dialog(
                webview.FileDialog.OPEN, allow_multiple=True,
                file_types=_FILE_DIALOG_TYPES)
            return _ok(paths=self._dialog_result_paths(result))
        except Exception as exc:
            return _err(exc)

    def choose_folder(self) -> Dict[str, Any]:
        """原生选择文件夹，只枚举其中受支持的文件，不触发目录上传确认。"""
        try:
            import webview
            if self._window is None:
                raise RuntimeError("窗口尚未就绪，请稍后重试")
            result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
            folders = self._dialog_result_paths(result)
            if not folders:
                return _ok(paths=[])
            with os.scandir(folders[0]) as entries:
                paths = [entry.path for entry in entries
                         if entry.is_file() and _is_supported_source(entry.name)]
            return _ok(paths=sorted(paths, key=os.path.normcase))
        except Exception as exc:
            return _err(exc)

    def add_files(self, paths: List[str]) -> Dict[str, Any]:
        added = []
        ignored = 0
        duplicates = 0
        existing = {os.path.normcase(os.path.abspath(rec.get("source") or rec["path"]))
                    for rec in self.files.values()}
        for path in paths:
            path = os.path.abspath(str(path).strip().strip('"'))
            if not path or not os.path.isfile(path) or not _is_supported_source(path):
                ignored += 1
                continue
            key = os.path.normcase(path)
            if key in existing:
                duplicates += 1
                continue
            if is_pdf_path(path):
                try:
                    pages = get_pdf_page_count(path)
                except Exception as exc:
                    self._log(
                        f"无法读取 {os.path.basename(path)}：{_describe_error(exc, '文件可能损坏或被占用。')}", True)
                    ignored += 1
                    continue
                self._iid += 1
                iid = str(self._iid)
                rec = {"path": path, "pages": pages, "status": "就绪", "kind": "pdf",
                       "range_text": "", "copies": 1, "iid": iid}
                self.files[iid] = rec
                added.append(rec)
            else:
                self._iid += 1
                iid = str(self._iid)
                kind = "image" if is_image_path(path) else "office"
                rec = {"path": path, "pages": 0, "status": "转换中", "kind": kind,
                       "range_text": "", "copies": 1, "iid": iid, "source": path}
                self.files[iid] = rec
                self._converting[iid] = rec
                self._executor.submit(self._convert_one, iid, path)
                added.append(rec)
            existing.add(key)
        if added:
            self._log(f"已添加 {len(added)} 个文件，当前共 {len(self.files)} 个待打印。")
            self._add_recent_files([r.get("source") or r["path"] for r in added])
        if ignored:
            self._log(f"已忽略 {ignored} 个不存在或不受支持的文件。", True)
        return _ok(added=added, total=len(self.files), ignored=ignored,
                   duplicates=duplicates)

    def _convert_one(self, iid: str, path: str) -> None:
        name = os.path.basename(path)
        self._log(f"正在转换：{name}")
        rec = self.files.get(iid)
        if rec is None:
            self._converting.pop(iid, None)
            return
        temp_dir = create_temp_dir()
        out = os.path.join(temp_dir, os.path.splitext(name)[0] + ".pdf")
        try:
            if is_image_path(path):
                info = image_to_pdf(path, out)
            else:
                info = convert_file(path, out)
            # 转换后的 PDF 必须在预览/打印期间持续存在，不能离开函数就删除。
            rec = self.files.get(iid)
            if rec is None:
                shutil.rmtree(temp_dir, ignore_errors=True)
                return
            rec["pages"] = int(info.get("pages") or get_pdf_page_count(out))
            rec["path"] = out
            rec["status"] = "就绪"
            rec["temp_dir"] = temp_dir
            self._log(f"转换完成：{name}（{rec['pages']} 页）")
        except (WordConvertError, ImageConvertError, BatchCancelled) as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            rec["status"] = "失败"
            rec["error"] = _describe_error(exc, "请检查文件是否损坏或受密码保护。")
            self._log(f"文件转换失败 {name}：{rec['error']}", True)
        except Exception as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            rec["status"] = "失败"
            rec["error"] = _describe_error(exc, "请检查文件是否损坏或受密码保护。")
            self._log(f"文件转换失败 {name}：{rec['error']}", True)
            # 详细堆栈只写日志文件，界面不再展示 Python 调用栈。
            get_logger().exception("文件转换异常：%s", name)
        finally:
            self._converting.pop(iid, None)
            if iid in self.files:
                self._post("file-update", dict(rec))

    def _add_recent_files(self, paths: List[str]) -> None:
        files = list(self.vars.get("recent_files") or [])
        for p in reversed(paths):
            if p in files:
                files.remove(p)
            files.insert(0, p)
        self.vars["recent_files"] = files[:12]
        try:
            settings_store.save_settings(self.vars)
        except Exception as exc:
            self._log(f"保存最近文件记录失败：{exc}", True)
        self._post("recent-updated", self.vars["recent_files"])

    def add_recent(self, path: str) -> Dict[str, Any]:
        if not path:
            return _err(ValueError("路径为空"))
        if os.path.isdir(path):
            with os.scandir(path) as entries:
                found = [entry.path for entry in entries
                         if entry.is_file() and _is_supported_source(entry.name)]
            return self.add_files(sorted(found, key=os.path.normcase))
        return self.add_files([path])

    @staticmethod
    def _release_conversion_temp(rec: Dict[str, Any]) -> None:
        temp_dir = rec.pop("temp_dir", "")
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def remove_files(self, iids: List[str]) -> Dict[str, Any]:
        for iid in iids:
            rec = self.files.pop(iid, None)
            if rec is not None:
                self._release_conversion_temp(rec)
        self._log("待打印列表已更新。")
        return _ok(total=len(self.files))

    def clear_files(self) -> Dict[str, Any]:
        for rec in self.files.values():
            self._release_conversion_temp(rec)
        self.files.clear()
        self._log("待打印列表已清空。")
        return _ok()

    def move_files(self, iids: List[str], direction: int) -> Dict[str, Any]:
        ids = list(self.files)
        selected = set(iids)
        indexes = [ids.index(iid) for iid in ids if iid in selected]
        if not indexes:
            return _ok(files=list(self.files.values()))
        ordered = sorted(indexes, reverse=direction > 0)
        for index in ordered:
            target = index + (1 if direction > 0 else -1)
            if target < 0 or target >= len(ids) or ids[target] in selected:
                continue
            ids[index], ids[target] = ids[target], ids[index]
        self.files = {iid: self.files[iid] for iid in ids}
        return _ok(files=list(self.files.values()))

    def dedupe_files(self) -> Dict[str, Any]:
        seen = set()
        removed = []
        for iid, rec in list(self.files.items()):
            key = os.path.normcase(os.path.abspath(rec.get("source") or rec["path"]))
            if key in seen:
                removed.append(iid)
                self.files.pop(iid, None)
                self._release_conversion_temp(rec)
            else:
                seen.add(key)
        if removed:
            # 用户反馈由前端按 removed 数量写日志（含 removed=0 的“未发现重复文件”），
            # 这里只落文件日志，避免同一条结果在界面上出现两次。
            get_logger().info(f"去重完成：移除 {len(removed)} 个重复文件。")
        return _ok(files=list(self.files.values()), removed=removed)

    def show_in_explorer(self, iid: str) -> Dict[str, Any]:
        rec = self.files.get(iid)
        if rec is None:
            return _err(KeyError("文件不存在"))
        try:
            path = os.path.normpath(rec.get("source") or rec["path"])
            subprocess.Popen(["explorer", "/select,", path])
            return _ok()
        except Exception as exc:
            return _err(exc)

    def refresh_files(self) -> Dict[str, Any]:
        """F5：重新读取源 PDF 页数并标记丢失/损坏文件。"""
        for rec in self.files.values():
            source = rec.get("source") or rec["path"]
            if rec.get("kind") != "pdf":
                continue
            if not os.path.isfile(source):
                rec["status"] = "失败"
                rec["error"] = "文件不存在"
                continue
            try:
                rec["pages"] = get_pdf_page_count(source)
                rec["path"] = source
                rec["status"] = "就绪"
                rec.pop("error", None)
            except Exception as exc:
                rec["status"] = "失败"
                rec["error"] = str(exc)
        self._post("files-replaced", list(self.files.values()))
        return _ok(files=list(self.files.values()))

    def set_file_options(self, iid: str, range_text: str, copies: int) -> Dict[str, Any]:
        return self.set_files_options([iid], range_text, copies)

    @staticmethod
    def _validate_range_for_record(rec: Dict[str, Any], range_text: str) -> Optional[str]:
        """校验 custom 范围文本对单个文件是否至少命中 1 页。

        返回 None 表示通过；否则返回可直接展示的中文错误：
          - 语法错误（parse_page_range 抛 PdfError）→ “范围格式无效：…”
          - 页码全部越界 → “《文件名》不包含第 2 页（共 1 页）”

        空文本等价“全部”，直接放行；转换中/失败/页数未知的文件无从判断，也
        放行并交给 compute_plan 逐文件兜底，避免把“暂时没有页数”或“本就不可
        打印”的文件误判成“范围无效”而拒绝保存整批。
        """
        text = (range_text or "").strip()
        if not text:
            return None
        if rec.get("status") in ("转换中", "失败"):
            return None
        total = int(rec.get("pages") or 0)
        if total < 1:
            return None
        name = os.path.basename(rec.get("source") or rec.get("path") or "")
        try:
            valid, dropped = resolve_range_pages("custom", text, total)
        except PdfError as exc:
            return f"范围格式无效：{_describe_error(exc)}"
        if not valid:
            pages_text = AppBridge._format_range_pages(dropped) if dropped else text
            return f"《{name}》不包含第 {pages_text} 页（共 {total} 页）"
        return None

    def _check_range_for_records(self, records: Sequence[Dict[str, Any]],
                                 range_text: str) -> None:
        """对所有目标文件做一次范围校验，任一无效即抛错（调用方不得写入）。

        先整批校验再统一写入，保证“一个文件范围无效”时不会留下“前几个文件已
        改、后面报错”的半完成状态。
        """
        text = (range_text or "").strip()
        if not text:
            return
        try:
            parse_page_range(text)
        except PdfError as exc:
            raise PdfError(f"范围格式无效：{_describe_error(exc)}") from exc
        for rec in records:
            message = self._validate_range_for_record(rec, text)
            if message:
                raise PdfError(message)

    def set_files_options(self, iids: List[str], range_text: str, copies: int,
                          inherit: bool = False, apply_all: bool = False) -> Dict[str, Any]:
        try:
            copies = 1 if inherit else int(copies)
            if not 1 <= copies <= 99:
                raise ValueError("单独份数需为 1-99 的整数")
            targets = list(self.files) if apply_all else list(iids)
            if not targets:
                raise ValueError("请先选择文件")
            # 目标记录（已不存在的 iid 按“跳过”处理）先全部校验，再统一写入，
            # 任一文件范围无效就整体拒绝，绝不部分写入。
            records = [self.files[iid] for iid in targets if iid in self.files]
            if not inherit:
                self._check_range_for_records(records, range_text)
            for rec in records:
                rec["range_text"] = "" if inherit else (range_text or "")
                rec["copies"] = None if inherit else copies
                # 范围已重新落定：旧的失败标记立即失效，等 compute_plan 重新判定。
                rec.pop("range_error", None)
            self.vars["per_file"] = not inherit
            settings_store.save_settings(self.vars)
            if range_text and not inherit:
                self._add_recent_range(range_text)
            self._post("files-replaced", list(self.files.values()))
            return _ok(files=list(self.files.values()))
        except Exception as exc:
            return _err(exc)

    def apply_range_to_all(self, range_text: str) -> Dict[str, Any]:
        """把范围应用到全部文件；任一文件无有效页则整体不修改。"""
        try:
            self._check_range_for_records(list(self.files.values()), range_text)
            for rec in self.files.values():
                rec["range_text"] = range_text
                rec.pop("range_error", None)
            if range_text:
                self._add_recent_range(range_text)
            return _ok()
        except Exception as exc:
            return _err(exc)

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
    @staticmethod
    def _format_range_pages(pages: Sequence[int], limit: int = 8) -> str:
        """把有效页码压缩成 1-5,8,10-12 这种可读写法。

        连续页码合并成区间；区间段数超过 limit 时以“…”收尾，避免一页文件的
        越界请求把 1,3,5,…,31 整串铺到文件行里。
        """
        ordered = sorted({int(p) for p in pages if int(p) > 0})
        if not ordered:
            return "全部"
        segments: List[str] = []
        start = prev = ordered[0]
        for page in ordered[1:]:
            if page == prev + 1:
                prev = page
                continue
            segments.append(str(start) if start == prev else f"{start}-{prev}")
            start = prev = page
        segments.append(str(start) if start == prev else f"{start}-{prev}")
        if len(segments) > limit:
            return ",".join(segments[:limit]) + "…"
        return ",".join(segments)

    @staticmethod
    def _range_label(mode: str, text: str, current_page: int = 1,
                     valid: Optional[Sequence[int]] = None,
                     total_pages: Optional[int] = None) -> str:
        """把范围模式翻译成文件行可读的“实际生效范围”。

        valid/total_pages 来自 resolve_range_pages：custom 文本里超出文件页数的
        页码已被丢弃，标签必须按有效页展示，绝不能回显原文中的越界页（曾出现
        1 页文件显示 “1,3,5,…,31”）。全部页 / 全部奇数页 / 全部偶数页收敛成
        简短文案，其余情况才列出实际页码。
        """
        pages = max(0, int(total_pages or 0))
        effective = sorted({int(p) for p in (valid or []) if int(p) > 0})
        if valid is not None and not effective:
            # 请求里的页码全部越界：_record_pages 会先抛错，这里只做防御。
            return "无有效页"
        if effective and pages:
            if effective == list(range(1, pages + 1)):
                return "全部"
            if effective == list(range(1, pages + 1, 2)):
                return "奇数页"
            if effective == list(range(2, pages + 1, 2)):
                return "偶数页"
            if mode in ("custom", "recent"):
                return AppBridge._format_range_pages(effective)
        if mode == "current":
            return f"当前页（第 {max(1, int(current_page or 1))} 页）"
        if mode == "odd":
            return "奇数页"
        if mode == "even":
            return "偶数页"
        if mode in ("custom", "recent"):
            return text.strip() or "全部"
        if effective:
            return AppBridge._format_range_pages(effective)
        return "全部"

    def _record_pages(self, rec: Dict[str, Any], opts: dict) -> Dict[str, Any]:
        """统一解析单文件范围、排版计划与有效份数，供预览/指标/打印复用。"""
        if not rec.get("pages"):
            raise PdfError(f"《{os.path.basename(rec.get('source') or rec['path'])}》没有可用页面")
        per_file = bool(opts.get("per_file", False))
        custom = bool(per_file and (rec.get("range_text") or "").strip())
        mode = "custom" if custom else str(opts.get("range_mode") or "all")
        text = str(rec.get("range_text") or "") if custom else str(opts.get("range_text") or "")
        if mode == "recent" and not text.strip():
            text = str(self.vars.get("last_range") or self.vars.get("range_text") or "")
        valid, dropped = resolve_range_pages(
            mode, text, int(rec["pages"]), int(opts.get("current_page") or 1))
        if not valid:
            raise PdfError(f"《{os.path.basename(rec.get('source') or rec['path'])}》范围没有有效页面")
        layout = self._layout_options(opts)
        # 计划必须与 compose_layout_pdf 口径一致：auto 方向跟随“选中范围第一页”
        # 的显示尺寸（page.rect 已计入 /Rotate），否则横版首页会算出竖版 plan，
        # 令 preview/plan_description/persist 的纸张方向与真实合成页冲突。
        first_landscape = False
        src = open_pdf(rec["path"])
        try:
            first_rect = src[valid[0] - 1].rect
            first_landscape = first_rect.width > first_rect.height
        finally:
            src.close()
        plan = build_layout_plan(len(valid), layout, first_landscape)
        # 有效副本数与 compute_plan/_build_snapshot 同一规则：整批流程取全局
        # 份数，逐文件流程才取文件级份数；文件行显示的份数据此同步。
        if _uses_per_file_copies(opts):
            copies = int(rec.get("copies") or 1)
        else:
            copies = int(opts.get("copies") or 1)
        return {"valid": valid, "dropped": dropped, "plan": plan,
                "copies": max(1, copies), "layout": layout,
                "range_mode": mode, "range_text": text,
                "range_label": self._range_label(
                    mode, text, int(opts.get("current_page") or 1),
                    valid=valid, total_pages=int(rec["pages"]))}

    def compute_plan(self, opts: dict) -> Dict[str, Any]:
        """按真实范围和逐文件份数计算批次指标。

        单个文件范围无效（例如对 1 页文件设了第 2 页）不再拖垮整批：该文件行
        标记 range_error/“范围无效”后跳过其计划，其余文件照常计算并回传 rows，
        前端因此可以继续刷新合法文件行。只要还有 1 个有效文件就返回 ok=true 并
        附带 errors；全部无效时返回 ok=false，但同样回传 rows/errors，前端仍需
        重绘这些行（显示“范围无效”）而不是把整批留在旧值。
        """
        try:
            selected_pages = 0
            side_counts = []
            copy_counts = []
            content_sides = 0
            rows = []
            errors = []
            for rec in self.files.values():
                if rec.get("status") in ("转换中", "失败"):
                    continue
                try:
                    resolved = self._record_pages(rec, opts)
                except PdfError as exc:
                    # 逐文件容错：只标记该文件，批次继续；错误原文保留给前端 tooltip。
                    message = str(exc)
                    rec["range_error"] = message
                    rec["effective_range"] = "范围无效"
                    rec["selected_pages"] = 0
                    rec["effective_copies"] = None
                    rec["plan_text"] = "—"
                    rows.append({
                        "iid": rec["iid"],
                        "effective_range": rec["effective_range"],
                        "selected_pages": rec["selected_pages"],
                        "effective_copies": rec["effective_copies"],
                        "plan_text": rec["plan_text"],
                        "range_error": message,
                    })
                    errors.append({
                        "iid": rec["iid"],
                        "name": os.path.basename(rec.get("source") or rec["path"]),
                        "message": message,
                    })
                    continue
                # 有效文件必须清除上一次的失败标记，否则修正范围后行仍显示“范围无效”。
                rec.pop("range_error", None)
                plan = resolved["plan"]
                blanks = plan.source_blank_added + plan.output_blank_added
                rec["selected_pages"] = len(resolved["valid"])
                # 文件行显示“当前实际生效范围”：全局范围变化后随指标刷新即时更新。
                rec["effective_range"] = resolved["range_label"]
                rec["effective_copies"] = resolved["copies"]
                rec["plan_text"] = (f"补{blanks} → " if blanks else "") + f"{plan.output_count} 面"
                # 前端 state.files 是反序列化副本，必须随响应把每文件字段回传，
                # 否则文件行会一直显示旧范围（例如永远是“全部”）。
                rows.append({
                    "iid": rec["iid"],
                    "effective_range": rec["effective_range"],
                    "selected_pages": rec["selected_pages"],
                    "effective_copies": rec["effective_copies"],
                    "plan_text": rec["plan_text"],
                    "range_error": None,
                })
                selected_pages += len(resolved["valid"])
                side_counts.append(plan.output_count)
                copy_counts.append(resolved["copies"])
                content_sides += sum(any(rel is not None for rel in side) for side in plan.sides)
            if errors and not side_counts:
                return {
                    "ok": False,
                    "error": "PdfError",
                    "message": "所有文件的范围都无效，请调整页码范围后重试。",
                    "rows": rows,
                    "errors": errors,
                }
            copies = max(1, int(opts.get("copies", 1)))
            # 执行模式严格跟随界面流程：绝不因“各文件份数不同/勾了逐文件设置”
            # 等数据差异静默改成逐文件，否则界面选“整批”却按逐文件打印。
            exec_mode = "per_file" if _uses_per_file_copies(opts) else "batch"
            copy_exec = "per_copy" if opts.get("copy_exec_on") else "all"
            if not opts.get("manual_duplex", False):
                # 单面输出绝不进入双面页序：plan_batch_summary 及其底层的
                # compute_merged_duplex_orders 会要求“每个文件面数为偶数”，
                # 单面 3 页文件本应完全合法，这里直接构造摘要即可。
                summary = {
                    "total_pages": sum(side_counts),
                    "sheets": sum(count * copies_used
                                  for count, copies_used in zip(side_counts, copy_counts)),
                    "fronts": [],
                    "backs": [],
                    "flip_count": 0,
                    "note": "单面输出：每个输出面单独打印，无需翻面",
                }
            else:
                summary = plan_batch_summary(
                    side_counts, copies=copies, exec_mode=exec_mode, copy_exec=copy_exec,
                    reverse=bool(opts.get("reverse", False)))
                if exec_mode == "per_file":
                    summary["sheets"] = sum((count // 2) * copies_used
                                            for count, copies_used in zip(side_counts, copy_counts))
                    summary["flip_count"] = sum(1 for count in side_counts if count > 1)
                else:
                    summary["sheets"] = (sum(side_counts) // 2) * copies
                    if content_sides <= 1:
                        summary["flip_count"] = 0
            summary["copies"] = copies
            return _ok(files=len(side_counts), total_pages=selected_pages,
                       summary=summary, rows=rows, errors=errors)
        except Exception as exc:
            return _err(exc)

    def preview_page(self, iid: str, page_no: int, opts: dict,
                     max_width: int = 900) -> Dict[str, Any]:
        """渲染应用当前范围、纸张、缩放与 N-up 后的真实输出面。"""
        try:
            rec = self.files.get(iid)
            if rec is None:
                raise KeyError("文件不存在")
            resolved = self._record_pages(rec, opts)
            with tempfile_tmp() as temp_dir:
                out = os.path.join(temp_dir, "preview.pdf")
                result = compose_layout_pdf(rec["path"], resolved["valid"], out,
                                            resolved["layout"])
                doc = fitz.open(out)
                try:
                    total = doc.page_count
                    if not total:
                        raise PdfError("排版结果为空")
                    pno = max(1, min(int(page_no), total))
                    page = doc[pno - 1]
                    # 纸张文案直接取真实合成页面的尺寸，确保“预览文案 == 实际页面”。
                    page_w, page_h = page.rect.width, page.rect.height
                    zoom = max_width / max(1.0, page_w)
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=False)
                    data = base64.b64encode(pix.tobytes("png")).decode("ascii")
                    plan = result.plan
                    # 缩放比例取排版时实测值（含裁切/自动旋转），而不是按纸张尺寸估算。
                    side_scales = result.side_scales or []
                    scale, overflow = (side_scales[pno - 1] if pno <= len(side_scales)
                                       else (0.0, False))
                    return _ok(dataurl=f"data:image/png;base64,{data}",
                               page=pno, total=total, plan=describe_plan(plan),
                               paper=f"{page_w / 72 * 25.4:.0f}×{page_h / 72 * 25.4:.0f} mm",
                               paper_pt=[round(page_w, 2), round(page_h, 2)],
                               landscape=page_w > page_h,
                               nup=plan.nup,
                               overflow=bool(overflow),
                               scale_percent=round(float(scale)))
                finally:
                    doc.close()
        except Exception as exc:
            return _err(exc)

    def plan_description(self, iid: str, opts: dict) -> Dict[str, Any]:
        """返回页序详情：整批页序摘要 + 双面示意文字 + 缩略图列表。"""
        try:
            rec = self.files.get(iid)
            if rec is None:
                raise KeyError("文件不存在")
            resolved = self._record_pages(rec, opts)
            plan = resolved["plan"]
            src_doc = open_pdf(rec["path"])
            try:
                first_rect = src_doc[resolved["valid"][0] - 1].rect
                src_w, src_h = first_rect.width, first_rect.height
            finally:
                src_doc.close()
            # 缩放百分比必须以源页尺寸为基准计算；传纸张尺寸会与真实排版不符。
            metrics = preview_metrics(plan, resolved["layout"], src_w, src_h)
            range_text = "、".join(str(p) for p in resolved["valid"][:24])
            if len(resolved["valid"]) > 24:
                range_text += f" …（共 {len(resolved['valid'])} 页）"
            return _ok(text=f"选取原页：{range_text}\n{describe_plan(plan)}",
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
        # 锁内只做“检查 + 立即占位”，保证两个并发调用只有一个能进入；
        # 参数校验/打印机探测等耗时操作一律在锁外完成，避免 GUI 调用被锁阻塞。
        with self._print_lock:
            if self._printing:
                return _err(RuntimeError("已有打印任务在进行"))
            self._printing = True
        try:
            if not self.files:
                raise RuntimeError("没有待打印文件")
            pending = [r for r in self.files.values() if r["status"] == "转换中"]
            if pending:
                raise RuntimeError("部分文件还在转换，请稍候")
            snapshot = self._build_snapshot(opts)
            check_printer_ready(snapshot["printer"])
        except Exception as exc:
            # 任何校验/快照/打印机探测失败都必须释放占位，否则界面永久禁用打印。
            self._release_print_slot()
            return _err(exc)
        self._active_printer = snapshot["printer"]
        self.stop_event.clear()
        self.pause_event.clear()
        self._paused = False
        self.current_job_id = 0
        try:
            self._executor.submit(self._worker_entry, snapshot)
        except Exception as exc:
            self._release_print_slot()
            return _err(exc)
        return _ok()

    def _release_print_slot(self) -> None:
        """释放打印占位（start_print 异常路径与工作线程 finally 共用）。"""
        with self._print_lock:
            self._printing = False

    def pause_toggle(self) -> Dict[str, Any]:
        if not self._printing:
            return _err(RuntimeError("当前没有进行中的打印任务"))
        if self._paused:
            if self.current_job_id > 0:
                try:
                    set_job_paused(self._active_printer, self.current_job_id, False)
                except Exception as exc:
                    return _err(exc)
            self.pause_event.clear()
            self._paused = False
            self._log("已继续打印")
        else:
            self.pause_event.set()
            self._paused = True
            self._log("已暂停打印（当前页完成后暂停）")
            if self.current_job_id > 0:
                try:
                    set_job_paused(self._active_printer, self.current_job_id, True)
                except Exception:
                    pass
        return _ok(paused=self._paused)

    def cancel_print(self) -> Dict[str, Any]:
        self.stop_event.set()
        self.pause_event.clear()
        if self.current_job_id > 0:
            try:
                cancel_job(self._active_printer, self.current_job_id)
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
        sources = opts.get("_paper_sources") or self.vars.get("_paper_sources") or [("自动选择", 7)]
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
        paper_name = str(opts.get("paper_name", "A4"))
        default_width, default_height = PAPER_SIZES_MM.get(paper_name, (210.0, 297.0))
        return LayoutOptions(
            paper_name=paper_name,
            paper_width_mm=float(opts.get("paper_width_mm", default_width)),
            paper_height_mm=float(opts.get("paper_height_mm", default_height)),
            orientation=str(opts.get("orientation", "auto")),
            source_name=str(opts.get("source_name", "")),
            source_code=source_code,
            auto_rotate=bool(opts.get("auto_rotate", False)),
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
        if not printer:
            raise PrintError("请先选择目标打印机")
        try:
            copies_number = float(str(opts.get("copies", 1)).strip())
            if not copies_number.is_integer():
                raise ValueError
            copies = int(copies_number)
        except (ValueError, OverflowError):
            raise ValueError("打印份数需为 1-99 的整数")
        if not 1 <= copies <= 99:
            raise ValueError("打印份数需为 1-99 的整数")
        percent = int(opts.get("scale_percent", 100))
        if not 10 <= percent <= 500:
            raise ValueError("缩放比例需为 10%-500%")
        for key in ("margin_left_mm", "margin_top_mm", "margin_right_mm", "margin_bottom_mm",
                    "crop_left_mm", "crop_top_mm", "crop_right_mm", "crop_bottom_mm"):
            value = float(opts.get(key, 0))
            if not 0 <= value <= 100:
                raise ValueError("边距与裁切需为 0-100 毫米")
        if str(opts.get("paper_name", "A4")) == CUSTOM_PAPER_NAME:
            width, height = float(opts.get("paper_width_mm", 0)), float(opts.get("paper_height_mm", 0))
            if not (10 <= width <= 2000 and 10 <= height <= 2000):
                raise ValueError("自定义纸张宽/高需为 10-2000 毫米")
        delay = float(opts.get("delay", opts.get("delay_seconds", 15.0)))
        if not 0 <= delay <= 600:
            raise ValueError("固定延时需为 0-600 秒")
        keep_output = bool(opts.get("keep_output", False))
        output_dir = str(opts.get("output_dir", "")).strip()
        if keep_output and not output_dir:
            raise ValueError("已选择保留排版 PDF，请先设置输出目录")
        if keep_output:
            os.makedirs(output_dir, exist_ok=True)

        # 执行流程只认界面选择（batch_all/batch_take/per_file），不做任何静默改写。
        print_flow = print_flow_of(opts)
        per_file_flow = print_flow == "per_file"
        jobs = []
        per_file_settings = {}
        for rec in self.files.values():
            if rec.get("status") in ("失败", "转换中"):
                continue
            resolved = self._record_pages(rec, opts)
            # 整批流程：忽略文件残留的旧份数，强制用全局 copies（一次任务打完）；
            # 逐文件流程：沿用该文件的实际生效份数。
            job_copies = resolved["copies"] if per_file_flow else copies
            job = {"iid": rec["iid"], "path": rec["path"], "pages": rec["pages"],
                   "abs_pages": resolved["valid"], "dropped": len(resolved["dropped"]),
                   "copies": job_copies,
                   # 排版后的实际纸面（已含 auto 首页/小册子发生的方向交换）。
                   "paper": resolved["plan"].paper_pt,
                   "name": os.path.basename(rec.get("source") or rec["path"])}
            jobs.append(job)
            # 逐文件设置只在逐文件流程生效：整批流程若保留它，单面分支会退化成
            # “每文件单独提交”，与界面“整批”矛盾。整批时保持空字典（假值）。
            if per_file_flow:
                per_file_settings[rec["iid"]] = {
                    "range_text": rec.get("range_text") or "", "copies": job_copies}
        if not jobs:
            raise PdfError("列表中没有可用于打印的页面，请检查文件与打印范围")
        layout = self._layout_options(opts)
        mode = str(opts.get("range_mode") or "all")
        if mode in ("custom", "recent"):
            self.vars["last_range"] = str(opts.get("range_text") or "")
            self._add_recent_range(self.vars["last_range"])
            settings_store.save_settings(self.vars)
        return {
            "printer": printer,
            "jobs": jobs,
            "copies": copies,
            # 逐文件设置仅在逐文件流程生效；整批流程为空字典（假值）。
            "per_file_settings": per_file_settings,
            "collate": opts.get("collate") in (True, "collate"),
            "manual_duplex": bool(opts.get("manual_duplex", False)),
            "print_flow": print_flow,
            "copy_exec_on": bool(opts.get("copy_exec_on", False)),
            "reverse": bool(opts.get("reverse", False)),
            "wait_mode": str(opts.get("wait_mode", "queue")),
            "delay": delay,
            "keep_output": keep_output,
            "auto_open": bool(opts.get("auto_open", False)),
            "output_dir": output_dir,
            "layout": layout,
            # 打印纸向必须取实际排版纸面：auto+横版首页、小册子 auto 的 plan 是
            # 横版，但 LayoutOptions 仍是原始 210×297，只看它会把 DEVMODE 判成竖版。
            # 批量混向时以首个文件为准（runtime 会先弹确认并要求逐文件流程），
            # 逐文件流程会在 _worker_main 中按各文件 paper 覆盖 landscape。
            "printer_options": self._printer_options(opts, layout, jobs[0]["paper"]),
        }

    def _printer_options(self, opts: dict, layout: LayoutOptions,
                         paper_pt: Optional[Sequence[float]] = None) -> PrinterOptions:
        """构造打印选项；paper_pt 为排版后的实际纸面（点）。

        传 paper_pt 时只按“合成纸宽>高”决定 DEVMODE 方向，保证打印方向与
        预览/合成页一致；不传时退回旧逻辑（显式方向或原始纸张宽高）。
        """
        orientation = opts.get("orientation", "auto")
        if paper_pt is not None:
            landscape = float(paper_pt[0]) > float(paper_pt[1])
        elif orientation == "landscape":
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

    def _printer_options_for_paper(self, base: PrinterOptions,
                                   paper_pt: Sequence[float]) -> PrinterOptions:
        """按单个文件的实际排版纸面复制打印选项（仅切换 DEVMODE 方向）。"""
        if not isinstance(base, PrinterOptions):
            # 兼容测试替身/旧快照：无法复制时保持原对象，由 _print_phase 回退。
            return base
        landscape = float(paper_pt[0]) > float(paper_pt[1])
        return dataclass_replace(base, landscape=landscape)

    # ============================================================ 工作线程
    def _worker_entry(self, snapshot: dict) -> None:
        completed = cancelled = False
        error_text = ""
        self._last_prepared = []
        try:
            self._worker_main(snapshot)
            completed = True
        except CancelledByUser:
            cancelled = True
            self._log("打印已取消。", True)
        except Exception as exc:
            error_text = _describe_error(exc, "请检查打印机与文件后重试。")
            self._log(f"打印失败：{error_text}", True)
            # 打包后（pythonw / 无控制台）stderr 可能不存在，堆栈必须进日志文件。
            get_logger().exception("打印任务异常")
        finally:
            self._release_print_slot()
            self._post("print-done", {"completed": completed,
                                      "cancelled": cancelled, "error": error_text})
            # 最终汇总：单面/双面任务结束后统一弹窗，由用户手动关闭。
            self._post("summary", self._summary_payload(
                completed, cancelled, error_text, snapshot))

    def _summary_payload(self, completed: bool, cancelled: bool,
                         error_text: str, snapshot: dict) -> Dict[str, Any]:
        """构造任务结束汇总（完成/取消/失败）的简洁中文信息。"""
        prepared = list(getattr(self, "_last_prepared", []) or [])
        copies = int(snapshot.get("copies") or 1)
        manual = bool(snapshot.get("manual_duplex", False))
        sheets = 0
        total_sides = 0
        for item in prepared:
            sides = int(item.get("sides") or 0)
            item_copies = int((item.get("job") or {}).get("copies") or copies)
            total_sides += sides
            # 手动双面按每 2 面 1 张计，单面输出按每面 1 张计。
            sheets += (sides // 2 if manual else sides) * item_copies
        if completed:
            title = "打印完成"
            message = (
                f"共 {len(prepared)} 个文件、约 {sheets} 张纸的打印任务已提交完成。\n"
                "请按提示完成翻面与取纸，并核对页序。")
        elif cancelled:
            title = "已取消"
            message = "打印已取消，剩余的临时文件已清理。"
        else:
            title = "打印失败"
            message = f"打印失败：{error_text or '请检查打印机与文件后重试。'}"
        return {"completed": bool(completed), "cancelled": bool(cancelled),
                "error": error_text, "title": title, "message": message,
                "files": len(prepared), "sheets": sheets,
                "total_sides": total_sides}

    def _worker_main(self, snapshot: dict) -> None:
        prepared = []
        # 供结束汇总统计（列表在同一线程内原地追加，取消时也能拿到已完成部分）。
        self._last_prepared = prepared
        # 流程只取界面选择并原样执行；日志先明确本次流程，便于对照界面排查。
        flow = print_flow_of(snapshot)
        snapshot["print_flow"] = flow
        self._log(f"执行流程：{PRINT_FLOW_LABELS[flow]}（与界面选择一致）")
        stamp = ""
        with tempfile_tmp() as target_root:
            stamp = os.path.basename(target_root)
            target_dir = snapshot["output_dir"] if snapshot["keep_output"] else target_root
            for job_index, job in enumerate(snapshot["jobs"], start=1):
                self._raise_if_stopped()
                name = job["name"]
                self._post("rows", [(job["iid"], "正在排版")])
                self._log(f"正在排版：《{name}》")
                abs_pages = job["abs_pages"]
                if abs_pages is None:
                    abs_pages = list(range(1, job["pages"] + 1))
                safe_name = self._safe_filename(os.path.splitext(name)[0])
                out_path = os.path.join(target_dir, f"{job_index:02d}_{safe_name}_{stamp}.pdf")
                # 排版单文件可能耗时较久：进入前再确认一次取消/暂停，避免关闭窗口后
                # 仍然继续排版，拖慢退出。
                self._raise_if_stopped()
                result = compose_layout_pdf(
                    job["path"], job["abs_pages"], out_path, snapshot["layout"])
                plan = result.plan
                self._log(f"排版完成：《{name}》 → {plan.output_count} 个输出面。")
                prepared.append({"job": job, "path": out_path,
                                 "sides": plan.output_count, "plan": plan,
                                 "paper": plan.paper_pt,
                                 # 逐文件流程按各自纸向打印（仅切换 landscape）。
                                 "printer_options": self._printer_options_for_paper(
                                     snapshot["printer_options"], plan.paper_pt)})
            copies = snapshot["copies"]
            total_sides = sum(p["sides"] for p in prepared)
            # 批量流程所有文件纸向一致时，用“排版后实际纸面”校正 DEVMODE 方向，
            # 避免 auto/小册子/跳过空白页等运行时变化让打印方向与合成页脱节。
            directions = {round(item["paper"][0], 3) > round(item["paper"][1], 3)
                          for item in prepared}
            if len(directions) == 1 and isinstance(
                    snapshot.get("printer_options"), PrinterOptions):
                actual_landscape = directions.pop()
                base_options = snapshot["printer_options"]
                if base_options.landscape != actual_landscape:
                    snapshot["printer_options"] = dataclass_replace(
                        base_options, landscape=actual_landscape)
                    self._log("打印纸向已按实际排版纸面校正为"
                              f"{'横版' if actual_landscape else '竖版'}。")
            # 整批流程若各文件排版后的纸张尺寸/方向不一致，合并打印会请求不同
            # 纸张，部分驱动会拒绝或触发换纸盒，先请用户确认（与 Tk 版行为一致）。
            if snapshot["print_flow"] in ("batch_all", "batch_take") and len(prepared) > 1:
                papers = {(round(item["paper"][0], 1), round(item["paper"][1], 1))
                          for item in prepared}
                if len(papers) > 1:
                    detail = "、".join(
                        f"《{item['job']['name']}》{item['paper'][0] / 72 * 25.4:.0f}×"
                        f"{item['paper'][1] / 72 * 25.4:.0f}mm" for item in prepared[:4])
                    if not self._flip_modal(
                            "纸张尺寸/方向不一致",
                            ["检测到各文件排版后的纸张尺寸或方向不一致：", detail,
                             "整批流程会在一次任务中请求不同纸张，部分驱动会拒绝或触发换纸盒。",
                             "建议取消后改用「逐文件执行」流程。"],
                            "仍要继续整批打印吗？", "仍要继续整批打印"):
                        raise CancelledByUser("用户取消：纸张不一致，建议切换逐文件流程。")
                    self._log("用户选择继续整批打印：按首个文件纸向整批打印"
                              "（其余文件的纸张方向不会单独生效）。")
            manual = snapshot["manual_duplex"]
            if manual:
                self._log(
                    f"手动双面：整批 {total_sides} 面（{total_sides // 2} 张 × 份数 {copies}），"
                    "先打全部正面，取纸翻面后再打反面。")
            if manual and snapshot["print_flow"] == "per_file":
                # 逐文件流程直接复用已排版好的单个 PDF，无需 target_dir/stamp。
                self._run_by_file_duplex(snapshot, prepared)
            elif manual:
                self._run_batch_duplex(snapshot, prepared, target_dir, stamp)
            else:
                self._run_single_sided(snapshot, prepared, target_dir, stamp)
        self._log("全部打印阶段已完成。")
        if snapshot.get("auto_open") and snapshot.get("keep_output") and snapshot.get("output_dir"):
            open_directory(snapshot["output_dir"])

    def _safe_filename(self, name: str) -> str:
        keep = "".join(c if c not in r'\/:*?"<>|' else "_" for c in name)
        return keep or "未命名"

    def _raise_if_stopped(self) -> None:
        if self.stop_event.is_set():
            raise CancelledByUser("用户取消。")
        while self.pause_event.is_set():
            if self.stop_event.is_set():
                raise CancelledByUser("用户取消。")
            import time
            time.sleep(0.2)

    def _sleep_interruptible(self, seconds: float) -> None:
        import time
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self._raise_if_stopped()
            time.sleep(min(0.2, max(0.0, end - time.monotonic())))

    # ---------------------------------------------------- 单面
    def _run_single_sided(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
        if snapshot.get("per_file_settings"):
            for item in prepared:
                job = item["job"]
                copies = int(job.get("copies") or snapshot["copies"])
                order = compute_single_sided_orders(
                    item["sides"], copies, snapshot["collate"], snapshot["reverse"])
                self._post("rows", [(job["iid"], "单面打印中")])
                self._print_phase(item["path"], order, f"PDF {job['name']} 单面",
                                  snapshot, f"单面打印 · 《{job['name']}》",
                                  printer_options=item.get("printer_options"))
                self._post("rows", [(job["iid"], "已完成")])
                self._log(f"单面打印完成：《{job['name']}》{item['sides']} 面 × {copies} 份。")
            return
        paths = [item["path"] for item in prepared]
        total = sum(item["sides"] for item in prepared)
        print_path = paths[0]
        if len(paths) > 1:
            print_path = os.path.join(target_dir, f"单面合并_{stamp}.pdf")
            merged = fitz.open()
            try:
                for path in paths:
                    src = fitz.open(path)
                    try:
                        merged.insert_pdf(src)
                    finally:
                        src.close()
                merged.save(print_path, garbage=3, deflate=True)
            finally:
                merged.close()
        order = compute_single_sided_orders(
            total, snapshot["copies"], snapshot["collate"], snapshot["reverse"])
        self._post("rows", [(item["job"]["iid"], "单面打印中") for item in prepared])
        self._print_phase(print_path, order, "PDF 单面打印", snapshot, "单面打印 · 正序")
        self._post("rows", [(item["job"]["iid"], "已完成") for item in prepared])
        self._log(f"单面打印完成：{total} 个输出面 × {snapshot['copies']} 份。")

    # ---------------------------------------------------- 手动双面
    def _run_batch_duplex(self, snapshot: dict, prepared: List[dict],
                          target_dir: str, stamp: str) -> None:
        flow = str(snapshot.get("print_flow") or "batch_all")
        pause_between_files = flow == "batch_take"
        counts = [item["sides"] for item in prepared]
        total = sum(counts)
        sheets = total // 2
        merged_path = os.path.join(target_dir, f"整批合并_{stamp}.pdf")
        merged_pages, _ = merge_layout_pdfs([item["path"] for item in prepared], merged_path)
        if merged_pages != total:
            raise PdfError("合并任务页数与预期不一致，请重试。")
        ranges = merged_side_ranges(counts)
        content_total = sum(
            any(rel is not None for rel in side)
            for item in prepared for side in item["plan"].sides)
        if content_total == 1:
            order = compute_single_sided_orders(1, snapshot["copies"])
            self._print_phase(merged_path, order, "PDF 单面打印", snapshot,
                              "单面打印 · 仅 1 页内容，无需翻面")
            self._post("rows", [(item["job"]["iid"], "已完成") for item in prepared])
            return

        copies = snapshot["copies"]
        if copies > 1 and snapshot["copy_exec_on"]:
            items = compute_merged_per_copy_orders(counts, copies, snapshot["reverse"])
            for copy_no, fronts, backs in items:
                prefix = f"第 {copy_no}/{copies} 份 · "
                self._print_phase(merged_path, fronts, f"PDF 第{copy_no}份正面", snapshot,
                                  prefix + "整批正面")
                if not self._ask_flip_batch(counts, fronts, backs, snapshot, sheets,
                                           copy_label=prefix):
                    raise CancelledByUser("用户取消翻面")
                self._print_batch_backs(snapshot, prepared, ranges, backs, merged_path,
                                        pause_between_files=pause_between_files, prefix=prefix)
                # 每份完成时只弹一次：非最后一份提示取走当前份，最后一份提示全部完成。
                self._batch_done_modal(total, sheets, snapshot, done=copy_no,
                                       total=copies, final=copy_no >= copies)
            return

        fronts, backs = compute_merged_duplex_orders(
            counts, copies, collate=snapshot["collate"], reverse=snapshot["reverse"])
        self._post("rows", [(item["job"]["iid"], "正面打印中") for item in prepared])
        self._print_phase(merged_path, fronts, "PDF 双面正面", snapshot,
                          "手动双面 · 整批正面（正序）")
        if not self._ask_flip_batch(counts, fronts, backs, snapshot, sheets):
            raise CancelledByUser("用户取消翻面")
        self._post("rows", [(item["job"]["iid"], "反面打印中") for item in prepared])
        self._print_batch_backs(snapshot, prepared, ranges, backs, merged_path,
                                pause_between_files=pause_between_files)
        self._batch_done_modal(total, sheets * copies, snapshot)

    def _print_batch_backs(self, snapshot: dict, prepared: List[dict],
                           ranges: List[Tuple[int, int]], backs: List[int],
                           merged_path: str, pause_between_files: bool = False,
                           prefix: str = "") -> None:
        groups = []
        for (start, end), item in zip(ranges, prepared):
            chunk = [page for page in backs if start <= page <= end]
            if chunk:
                groups.append((item, chunk))
        if not pause_between_files or len(groups) <= 1:
            self._print_phase(merged_path, backs, f"PDF {prefix}双面反面", snapshot,
                              prefix + "整批反面（全局逆序）")
            return
        ordered = list(reversed(groups))
        for index, (item, chunk) in enumerate(ordered, start=1):
            job = item["job"]
            self._print_phase(merged_path, chunk, f"PDF {prefix}{job['name']}反面", snapshot,
                              f"{prefix}逐文件取纸 · 《{job['name']}》反面")
            self._post("rows", [(job["iid"], "已完成")])
            if index < len(ordered):
                next_name = ordered[index][0]["job"]["name"]
                self._file_done_modal(job, len(chunk) * 2, snapshot, next_name=next_name,
                                      done=index, total=len(ordered))

    def _run_by_file_duplex(self, snapshot: dict, prepared: List[dict]) -> None:
        total_files = len(prepared)
        for index, item in enumerate(prepared, start=1):
            job = item["job"]
            sides = item["sides"]
            copies = int(job.get("copies") or snapshot["copies"])
            content_total = sum(any(rel is not None for rel in side)
                                for side in item["plan"].sides)
            if content_total == 1:
                fronts = compute_single_sided_orders(1, copies, snapshot["collate"],
                                                     snapshot["reverse"])
                backs = []
            else:
                fronts, backs = compute_duplex_orders(
                    sides, copies, collate=snapshot["collate"], reverse=snapshot["reverse"])
            self._post("rows", [(job["iid"], "正面打印中")])
            self._print_phase(item["path"], fronts, f"PDF {job['name']} 正面", snapshot,
                              f"逐文件 {index}/{total_files} · 《{job['name']}》正面",
                              printer_options=item.get("printer_options"))
            if backs:
                if not self._ask_flip_file(job, sides * copies, backs, snapshot):
                    raise CancelledByUser("用户取消翻面")
                self._post("rows", [(job["iid"], "反面打印中")])
                self._print_phase(item["path"], backs, f"PDF {job['name']} 反面", snapshot,
                                  f"逐文件 {index}/{total_files} · 《{job['name']}》反面",
                                  printer_options=item.get("printer_options"))
            self._post("rows", [(job["iid"], "已完成")])
            next_name = prepared[index]["job"]["name"] if index < total_files else ""
            self._file_done_modal(job, sides * copies, snapshot, next_name=next_name,
                                  done=index, total=total_files)

    # ---------------------------------------------------- 打印阶段
    def _print_phase(self, pdf_path: str, order: List[int], job_name: str,
                     snapshot: dict, phase_text: str,
                     printer_options: Optional[PrinterOptions] = None) -> None:
        """提交一个打印阶段。

        printer_options 用于逐文件流程按该文件实际排版纸面设置 DEVMODE 方向；
        未提供时回退到整批快照的方向（批量流程所有文件共用一种纸向）。
        """
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

        options = (printer_options if printer_options is not None
                   else snapshot["printer_options"])
        printer = GdiPrinter(snapshot["printer"], options)
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
                        snapshot: dict, sheets: int, copy_label: str = "") -> bool:
        self._raise_if_stopped()
        label = (copy_label or "").strip()
        # copy_label 形如“第 2/3 份 · ”，用于多份逐份执行时区分当前进度。
        prefix = f"{label} " if label else ""
        detail = [
            f"{prefix}本批 {sheets} 张 × {snapshot['copies']} 份已打印正面完成。",
            "请取出打印纸，保持页序，整体翻面后重新放入纸盒。",
            "确认后点击「已翻面，继续打印反面」；取消将终止本次任务。",
        ]
        title = "手动双面 · 请翻面"
        if label:
            title = f"{title}（{label.rstrip('· ').strip()}）"
        return self._flip_modal(title, detail,
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

    def _flip_modal(self, title: str, detail_lines: List[str], sample: str,
                    proceed_label: str = "已翻面，继续打印反面") -> bool:
        """弹出确认框并等待前端回调，超时按取消处理。

        每次调用必须先清空事件并复位结果，否则上一轮的“继续”可能被误判为
        本次确认。等待带上限：前端被关闭、JS 失效或用户直接退出时，工作线程
        仍能走到 ``_worker_entry`` 的 finally，保证最终汇总不被吞掉。
        """
        self._flip_event.clear()
        self._flip_result = False
        self._post("flip-modal", {"title": title, "detail": detail_lines,
                                   "sample": sample, "proceed_label": proceed_label})
        if not self._flip_event.wait(self._modal_timeout):
            self._flip_result = False
            self._log(f"等待确认超时（{self._modal_timeout:.0f} 秒）：“{title}”，"
                      "按取消处理。", True)
            return False
        return self._flip_result

    def _file_done_modal(self, job: dict, sides: int, snapshot: dict,
                         next_name: str = "", done: int = 1, total: int = 1) -> None:
        title = "全部完成，请取出最后一叠纸" if not next_name else f"请取出《{job['name']}》"
        detail = [
            f"文件：《{job['name']}》（进度 {done}/{total}）",
            f"本文件/分组已完成 {sides} 个输出面，约 {max(1, sides // 2)} 张纸。",
            "未确认前不会继续后续打印；取消将终止整批任务。",
        ]
        if next_name:
            detail.append(f"下一个：《{next_name}》")
        if not self._flip_modal(title, detail, "请确认纸张已取出后再继续。",
                                "已取出，全部完成" if not next_name else "已取出，继续打印"):
            raise CancelledByUser("用户在取纸提示中取消")
        self._log(f"《{job['name']}》已完成，{sides} 个输出面。")

    def _batch_done_modal(self, sides: int, sheets: int, snapshot: dict,
                          done: int = 1, total: int = 1, final: bool = True) -> None:
        """整批双面完成后的取纸提示（每份只弹一次）。

        final=True：本次是最后一份，提示“全部完成，请取出全部纸张”；
        final=False：仍有后续份，提示“第 X/Y 份已打完，请取出”后继续。
        """
        if final:
            title = "全部完成，请取出全部纸张"
            detail = [f"本批 {sides} 个输出面已完成，约 {sheets} 张纸。",
                      "请整理并核对纸叠；未确认前不会继续其他阶段。"]
            sample = "整批打印流程已完成。"
            proceed_label = "已取出，全部完成"
            cancel_message = "用户在完成提示中取消"
        else:
            title = f"第 {done}/{total} 份已打完，请取出"
            detail = [f"第 {done}/{total} 份已完成：{sides} 个输出面，约 {sheets} 张纸。",
                      "请取出这一份并整理好；确认后继续打印下一份。"]
            sample = f"剩余 {max(0, total - done)} 份待打印。"
            proceed_label = "已取出，继续打印下一份"
            cancel_message = "用户在份间取纸提示中取消"
        if not self._flip_modal(title, detail, sample, proceed_label):
            raise CancelledByUser(cancel_message)
        if final:
            self._log(f"整批双面完成：{sides} 个输出面，{sheets} 张。")
        else:
            self._log(f"第 {done}/{total} 份完成：{sides} 个输出面，{sheets} 张。")


    def _confirm_uncertain_phase(self, phase_text: str, reason: str) -> None:
        """队列状态无法确认时请用户裁决；无响应（超时）一律按取消处理。"""
        # 弹窗前必须先写日志：否则界面只剩一个模态框，日志里完全看不出原因
        # （是队列读取失败还是超时），用户与现场支持都无法复盘。
        self._log(f"{phase_text}：打印状态待确认——{reason}", True)
        self._phase_confirm_event.clear()
        self._phase_confirm_result = False
        self._post("phase-confirm", {"text": phase_text, "reason": reason})
        if not self._phase_confirm_event.wait(self._modal_timeout):
            self._log(f"{phase_text} 状态确认超时（{self._modal_timeout:.0f} 秒），"
                      "按取消处理。", True)
            raise CancelledByUser(f"{phase_text} 状态无法确认（等待确认超时），已取消。")
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

    def export_log(self, text: str) -> Dict[str, Any]:
        try:
            import webview
            if self._window is None:
                raise RuntimeError("窗口尚未就绪")
            result = self._window.create_file_dialog(
                webview.FileDialog.SAVE, save_filename="PDF打印日志.txt",
                file_types=("文本文件 (*.txt)",))
            paths = self._dialog_result_paths(result)
            if not paths:
                return _ok(cancelled=True)
            with open(paths[0], "w", encoding="utf-8-sig", newline="\n") as handle:
                handle.write(str(text or ""))
            return _ok(path=paths[0])
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
