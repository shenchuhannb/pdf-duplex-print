# -*- coding: utf-8 -*-
"""旧 Tk → WebView 完整功能移植的行为回归。"""
import os
import sys
from types import SimpleNamespace

import fitz
import pytest

import settings_store
from core_utils import CancelledByUser, PrintError, PrintTimeoutError
from pdf_layout import CUSTOM_PAPER_NAME, LayoutOptions, build_layout_plan
from printer_core import PrinterOptions, devmode_field_updates
from webapp.bridge import AppBridge


@pytest.fixture
def bridge(monkeypatch):
    monkeypatch.setattr(settings_store, "load_settings", lambda: {})
    monkeypatch.setattr(settings_store, "save_settings", lambda *args, **kwargs: None)
    instance = AppBridge()
    instance._executor.shutdown(wait=True)
    instance._log = lambda *_args, **_kwargs: None
    instance._post = lambda *_args, **_kwargs: None
    return instance


def make_pdf(path, pages=3, prefix="PAGE"):
    doc = fitz.open()
    for index in range(pages):
        page = doc.new_page(width=400, height=600)
        page.insert_text((40, 70), f"{prefix}_{index + 1}", fontsize=20)
    doc.save(path)
    doc.close()


def add_pdf(bridge, path, pages=3, prefix="PAGE"):
    make_pdf(path, pages, prefix)
    return bridge.add_files([str(path)])["added"][0]


def opts(**changes):
    result = settings_store.default_settings()
    result.update(printer="SIMULATED", current_page=1, preview_view="single")
    result.update(changes)
    return result


def test_preview_is_composed_output_not_original_pdf(bridge, tmp_path):
    rec = add_pdf(bridge, tmp_path / "three.pdf", pages=3)
    result = bridge.preview_page(rec["iid"], 1, opts(nup=2, manual_duplex=True))
    assert result["ok"] is True
    assert result["total"] == 2  # 3 页补为 4 页，再按 2 合 1 输出 2 面
    assert result["nup"] == 2
    assert result["dataurl"].startswith("data:image/png;base64,")


def test_preview_honors_custom_paper_and_range(bridge, tmp_path):
    rec = add_pdf(bridge, tmp_path / "five.pdf", pages=5)
    result = bridge.preview_page(rec["iid"], 1, opts(
        range_mode="custom", range_text="1-3", paper_name=CUSTOM_PAPER_NAME,
        paper_width_mm=148, paper_height_mm=210))
    assert result["ok"] is True
    assert result["paper"] == "148×210 mm"
    assert "1、2、3" in bridge.plan_description(rec["iid"], opts(
        range_mode="custom", range_text="1-3"))["text"]


def test_metrics_include_ranges_and_per_file_copies(bridge, tmp_path):
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=4)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=3)
    bridge.set_files_options([first["iid"]], "1-2", 2)
    # 逐文件份数只在“逐文件执行”流程生效（整批流程一律用全局份数）。
    options = opts(per_file=True, manual_duplex=True, copies=3, print_flow="per_file")
    result = bridge.compute_plan(options)
    assert result["ok"] is True
    # 第一个文件选 2 页，第二个文件默认 3 页，共 5 页。
    assert result["total_pages"] == 5
    assert first["effective_copies"] == 2
    assert second["effective_copies"] == 1
    assert first["selected_pages"] == 2


def test_file_move_dedupe_and_inherit(bridge, tmp_path):
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=1)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=1)
    moved = bridge.move_files([first["iid"]], 1)
    assert [rec["iid"] for rec in moved["files"]] == [second["iid"], first["iid"]]
    bridge.files["duplicate"] = dict(first, iid="duplicate")
    result = bridge.dedupe_files()
    assert result["removed"] == ["duplicate"]
    bridge.set_files_options([first["iid"]], "1", 4)
    inherited = bridge.set_files_options([first["iid"]], "", 1, inherit=True)
    target = next(rec for rec in inherited["files"] if rec["iid"] == first["iid"])
    assert target["range_text"] == "" and target["copies"] is None


def test_snapshot_global_copies_and_keep_output_validation(bridge, tmp_path):
    add_pdf(bridge, tmp_path / "a.pdf", pages=2)
    snapshot = bridge._build_snapshot(opts(copies=3, manual_duplex=False))
    assert snapshot["copies"] == 3
    assert snapshot["jobs"][0]["copies"] == 3
    with pytest.raises(ValueError, match="输出目录"):
        bridge._build_snapshot(opts(keep_output=True, output_dir=""))
    snapshot = bridge._build_snapshot(opts(
        keep_output=True, output_dir=str(tmp_path / "out"), copies=2))
    assert snapshot["keep_output"] is True
    assert snapshot["output_dir"] == str(tmp_path / "out")


def prepared(bridge, tmp_path, count=2):
    result = []
    for index in range(count):
        path = tmp_path / f"prepared-{index}.pdf"
        make_pdf(path, pages=2, prefix=f"FILE{index}")
        result.append({
            "path": str(path), "sides": 2,
            "plan": build_layout_plan(2, LayoutOptions(manual_duplex=True)),
            "job": {"iid": str(index), "name": f"{index}.pdf", "copies": 1},
        })
    return result


def test_per_copy_batch_executes_each_copy_with_matching_backs(bridge, tmp_path, monkeypatch):
    items = prepared(bridge, tmp_path)
    calls = []
    flip_labels = []
    monkeypatch.setattr(bridge, "_print_phase",
                        lambda path, order, *args: calls.append((path, list(order))))
    monkeypatch.setattr(bridge, "_ask_flip_batch",
                        lambda *args, **kwargs: flip_labels.append(kwargs.get("copy_label")) or True)
    completed = []
    monkeypatch.setattr(bridge, "_batch_done_modal",
                        lambda *args, **kwargs: completed.append((args, kwargs)))
    snapshot = {"copies": 2, "copy_exec_on": True, "print_flow": "batch_all",
                "collate": True, "reverse": False}
    bridge._run_batch_duplex(snapshot, items, str(tmp_path), "test")
    assert [order for _path, order in calls] == [[1, 3], [4, 2], [1, 3], [4, 2]]
    assert len(completed) == 2
    assert all(path.endswith("整批合并_test.pdf") for path, _order in calls)
    # 逐份执行：翻面提示必须带“第 X/Y 份”标签。
    assert flip_labels == ["第 1/2 份 · ", "第 2/2 份 · "]
    # 每份完成只弹一次：非最后一份 final=False，最后一份 final=True。
    assert [kwargs["done"] for _args, kwargs in completed] == [1, 2]
    assert [kwargs["total"] for _args, kwargs in completed] == [2, 2]
    assert [kwargs["final"] for _args, kwargs in completed] == [False, True]


def test_batch_take_uses_global_pages_on_merged_pdf(bridge, tmp_path, monkeypatch):
    items = prepared(bridge, tmp_path)
    calls = []
    monkeypatch.setattr(bridge, "_print_phase",
                        lambda path, order, *args: calls.append((path, list(order))))
    monkeypatch.setattr(bridge, "_ask_flip_batch", lambda *args, **kwargs: True)
    monkeypatch.setattr(bridge, "_batch_done_modal", lambda *args, **kwargs: None)
    monkeypatch.setattr(bridge, "_file_done_modal", lambda *args, **kwargs: None)
    snapshot = {"copies": 1, "copy_exec_on": False, "print_flow": "batch_take",
                "collate": True, "reverse": False}
    bridge._run_batch_duplex(snapshot, items, str(tmp_path), "test")
    assert [order for _path, order in calls] == [[1, 3], [4], [2]]
    assert all(path.endswith("整批合并_test.pdf") for path, _order in calls)


def test_single_sided_global_copies_are_not_lost(bridge, tmp_path, monkeypatch):
    items = prepared(bridge, tmp_path)
    calls = []
    monkeypatch.setattr(bridge, "_print_phase",
                        lambda path, order, *args: calls.append(list(order)))
    snapshot = {"copies": 3, "per_file_settings": False, "collate": True, "reverse": False}
    bridge._run_single_sided(snapshot, items, str(tmp_path), "test")
    assert calls == [[1, 2, 3, 4, 1, 2, 3, 4, 1, 2, 3, 4]]


def test_restore_defaults_preserves_recent_records(bridge, monkeypatch):
    bridge.vars["recent_files"] = ["a.pdf"]
    bridge.vars["recent_ranges"] = ["1-2"]
    bridge.vars["copies"] = 7
    monkeypatch.setattr(settings_store, "save_settings", lambda *args, **kwargs: None)
    result = bridge.restore_defaults()
    assert result["ok"] is True
    assert result["settings"]["copies"] == 1
    assert result["settings"]["recent_files"] == ["a.pdf"]
    assert result["settings"]["recent_ranges"] == ["1-2"]


def test_restore_acrobat_defaults_returns_single_sided_baseline(bridge, monkeypatch):
    """restore_defaults(acrobat=True) 必须整份切到 Acrobat 基线。

    该分支是“恢复 Acrobat 默认”按钮的真实后端入口：单面输出、全部页面、
    份数 1、整批流程，并清掉批注/边距/版式等自定义项。
    """
    bridge.vars.update({
        "manual_duplex": True,
        "annots": True,
        "margin_left_mm": 12.0,
        "range_mode": "custom",
        "range_text": "3-9",
        "copies": 5,
        "print_flow": "per_file",
        "per_file": True,
        "reverse": True,
        "nup": 4,
    })
    bridge.vars["recent_files"] = ["a.pdf"]
    bridge.vars["recent_ranges"] = ["1-2"]
    monkeypatch.setattr(settings_store, "save_settings", lambda *args, **kwargs: None)

    result = bridge.restore_defaults(True)

    assert result["ok"] is True
    settings = result["settings"]
    assert settings["manual_duplex"] is False
    assert settings["annots"] is False
    assert settings["margin_left_mm"] == 0.0
    assert settings["range_mode"] == "all"
    assert settings["copies"] == 1
    assert settings["print_flow"] == "batch_all"
    assert settings["per_file"] is False
    assert settings["reverse"] is False
    assert settings["nup"] == 1
    assert settings["auto_rotate"] is True
    assert settings["auto_rotate_choice_version"] == 1
    # Acrobat 基线同样不能丢最近记录。
    assert settings["recent_files"] == ["a.pdf"]
    assert settings["recent_ranges"] == ["1-2"]
    assert bridge.vars["manual_duplex"] is False


# ============================================================ 打印控制
def snapshot_minimal(**changes):
    result = {"printer": "SIMULATED", "jobs": [], "copies": 1,
              "per_file_settings": False, "collate": True, "manual_duplex": False,
              "print_flow": "batch_all", "copy_exec_on": False, "reverse": False,
              "wait_mode": "queue", "delay": 0, "keep_output": False,
              "auto_open": False, "output_dir": "", "layout": None,
              "printer_options": None}
    result.update(changes)
    return result


def answering_post(bridge, posts, answer):
    """构造 _post 替身：弹窗事件按 answer 立即回调 confirm_*。"""

    def post(kind, payload):
        posts.append((kind, payload))
        if kind == "flip-modal":
            bridge.confirm_flip(answer["proceed"])
        elif kind == "phase-confirm":
            bridge.confirm_phase(answer["proceed"])

    return post


def test_start_print_guards_pause_cancel_and_confirm_events(bridge, tmp_path, monkeypatch):
    # 1) 无文件 / 转换中 / 已有任务，都必须拒绝启动。
    assert bridge.start_print({})["ok"] is False
    bridge.files["1"] = {"iid": "1", "path": "x.pdf", "status": "转换中"}
    assert "转换" in bridge.start_print({})["message"]
    bridge.files.clear()
    bridge._printing = True
    assert "已有打印任务" in bridge.start_print({})["message"]
    bridge._printing = False

    # 2) 参数校验失败必须原样返回错误，而不是进入打印。
    rec = add_pdf(bridge, tmp_path / "guard.pdf")
    monkeypatch.setattr("webapp.bridge.check_printer_ready", lambda name: None)
    bad = bridge.start_print(opts(copies=0))
    assert bad["ok"] is False and "份数" in bad["message"]

    # 3) 正常启动：提交线程池并置打印态。
    submitted = []
    bridge._executor = SimpleNamespace(submit=lambda fn, *args: submitted.append((fn, args)))
    result = bridge.start_print(opts(copies=2))
    assert result["ok"] is True
    assert bridge._printing is True
    assert bridge.stop_event.is_set() is False
    assert submitted[0][0].__name__ == "_worker_entry"
    _ = rec

    # 4) 暂停/继续：首次 set_job_paused 失败被吞掉，但状态仍要正确翻转。
    bridge.current_job_id = 7
    pause_calls = []

    def flaky_set_job_paused(*args):
        pause_calls.append(args)
        if len(pause_calls) == 1:
            raise RuntimeError("spooler")

    monkeypatch.setattr("webapp.bridge.set_job_paused", flaky_set_job_paused)
    assert bridge.pause_toggle()["paused"] is True
    assert bridge.pause_event.is_set() is True
    assert bridge.pause_toggle()["paused"] is False
    assert bridge.pause_event.is_set() is False

    # 5) 取消打印：cancel_job 失败也要保证事件置位。
    monkeypatch.setattr("webapp.bridge.cancel_job",
                        lambda *args: (_ for _ in ()).throw(RuntimeError("gone")))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge.cancel_print()
    assert bridge.stop_event.is_set() is True
    assert bridge._flip_event.is_set() is True
    assert bridge._phase_confirm_event.is_set() is True

    # 6) 前端回调只记录结果并唤醒等待。
    assert bridge.confirm_flip(True)["ok"] is True
    assert bridge._flip_result is True
    assert bridge.confirm_phase(False)["ok"] is True
    assert bridge._phase_confirm_result is False


def test_pause_toggle_requires_running_job(bridge):
    bridge._printing = False
    result = bridge.pause_toggle()
    assert result["ok"] is False
    assert "没有进行中" in result["message"]


def test_worker_entry_reports_done_and_final_summary(bridge, monkeypatch):
    events = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: events.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._last_prepared = []

    def run(snapshot):
        # _worker_entry 会重置列表，真实 worker 在运行时原地追加已排版项。
        bridge._last_prepared = [{"sides": 2, "job": {"copies": 1}}]

    bridge._worker_main = run
    bridge._worker_entry(snapshot_minimal())

    assert [kind for kind, _payload in events] == ["print-done", "summary"]
    assert events[0][1] == {"completed": True, "cancelled": False, "error": ""}
    assert events[1][1]["title"] == "打印完成"
    assert events[1][1]["sheets"] == 2
    assert bridge._printing is False


def test_worker_entry_reports_cancelled_and_failed(bridge, monkeypatch):
    events = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: events.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._last_prepared = []

    def cancelled(snapshot):
        raise CancelledByUser("用户取消")

    bridge._worker_main = cancelled
    bridge._worker_entry(snapshot_minimal())
    assert events[0][1]["cancelled"] is True
    assert events[1][1]["title"] == "已取消"
    assert events[1][1]["cancelled"] is True

    def boom(snapshot):
        raise RuntimeError("boom")

    bridge._worker_main = boom
    bridge._worker_entry(snapshot_minimal())
    assert events[2][1]["error"] == "boom"
    assert events[3][1]["title"] == "打印失败"
    assert "boom" in events[3][1]["message"]
    assert bridge._printing is False


def test_print_phase_queue_delay_progress_and_timeout(bridge, monkeypatch):
    events = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: events.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    progress = []

    class Printer:
        last_job_id = 5

        def print_pages(self, path, order, name, stop_event,
                        on_page=None, pause_event=None, on_job_start=None):
            if on_job_start:
                on_job_start(5)
            if on_page:
                on_page(1, len(order), order[0])
            return 5

    monkeypatch.setattr("webapp.bridge.GdiPrinter", lambda *_args: Printer())
    monkeypatch.setattr(bridge, "_sleep_interruptible", lambda seconds: progress.append(seconds))
    monkeypatch.setattr("webapp.bridge.wait_for_print_job", lambda *args, **kwargs: True)

    bridge._print_phase("a.pdf", [1, 2], "job", snapshot_minimal(), "阶段一")
    assert bridge.current_job_id == 5
    assert ("status", "阶段一") in events
    assert any(kind == "progress" for kind, _payload in events)
    assert progress == [2.0]  # POST_JOB_SETTLE_SECONDS

    # 空页序直接跳过，不产生状态事件。
    events.clear()
    bridge._print_phase("a.pdf", [], "job", snapshot_minimal(), "空阶段")
    assert events == []

    # 固定延时模式：先等待 delay，再等结算时长。
    progress.clear()
    bridge._print_phase("a.pdf", [1], "job",
                        snapshot_minimal(wait_mode="delay", delay=3), "延时阶段")
    assert progress == [3, 2.0]

    # 队列超时进入“不确定阶段确认”。
    def timeout(*args, **kwargs):
        raise PrintTimeoutError("queue timeout")

    monkeypatch.setattr("webapp.bridge.wait_for_print_job", timeout)
    confirmed = []
    monkeypatch.setattr(bridge, "_confirm_uncertain_phase",
                        lambda text, reason: confirmed.append((text, reason)))
    bridge._print_phase("a.pdf", [1], "job", snapshot_minimal(), "超时阶段")
    assert confirmed and confirmed[0][0] == "超时阶段"

    # 任务号无效视为提交失败。
    class ZeroPrinter(Printer):
        def print_pages(self, *args, **kwargs):
            return 0

    monkeypatch.setattr("webapp.bridge.GdiPrinter", lambda *_args: ZeroPrinter())
    with pytest.raises(PrintError, match="提交失败"):
        bridge._print_phase("a.pdf", [1], "job", snapshot_minimal(), "失败阶段")


def test_flip_and_done_modals_use_expected_titles_and_labels(bridge, monkeypatch):
    posts = []
    answer = {"proceed": True}
    monkeypatch.setattr(bridge, "_post", answering_post(bridge, posts, answer))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._modal_timeout = 1.0

    assert bridge._ask_flip_batch([2], [1], [2], snapshot_minimal(copies=2), 1) is True
    assert posts[-1][0] == "flip-modal"
    assert posts[-1][1]["title"] == "手动双面 · 请翻面"
    assert posts[-1][1]["proceed_label"] == "已翻面，继续打印反面"

    assert bridge._ask_flip_file({"name": "a.pdf"}, 2, [2], snapshot_minimal()) is True
    assert posts[-1][1]["title"] == "逐文件双面 · 请翻面"

    bridge._file_done_modal({"name": "a.pdf"}, 3, snapshot_minimal(),
                            next_name="b.pdf", done=1, total=2)
    assert posts[-1][1]["title"] == "请取出《a.pdf》"
    assert posts[-1][1]["proceed_label"] == "已取出，继续打印"

    bridge._file_done_modal({"name": "last.pdf"}, 2, snapshot_minimal())
    assert posts[-1][1]["title"] == "全部完成，请取出最后一叠纸"
    assert posts[-1][1]["proceed_label"] == "已取出，全部完成"

    bridge._batch_done_modal(4, 2, snapshot_minimal())
    assert posts[-1][1]["title"] == "全部完成，请取出全部纸张"
    assert posts[-1][1]["proceed_label"] == "已取出，全部完成"

    # 用户在任一取纸/翻面提示中取消，都必须终止任务。
    answer["proceed"] = False
    # 翻面提示只返回用户选择，由调用方决定终止；取纸/完成提示自身直接抛取消。
    assert bridge._ask_flip_batch([2], [1], [2], snapshot_minimal(), 1) is False
    assert bridge._ask_flip_file({"name": "a.pdf"}, 2, [2], snapshot_minimal()) is False
    with pytest.raises(CancelledByUser, match="取纸"):
        bridge._file_done_modal({"name": "a.pdf"}, 2, snapshot_minimal())
    with pytest.raises(CancelledByUser, match="完成提示"):
        bridge._batch_done_modal(2, 1, snapshot_minimal())


def test_batch_progress_modals_show_copy_index_and_final_state(bridge, monkeypatch):
    """多份逐份执行：翻面弹窗带“第 X/Y 份”，取纸弹窗区分非最后/最后一份。"""
    posts = []
    answer = {"proceed": True}
    monkeypatch.setattr(bridge, "_post", answering_post(bridge, posts, answer))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._modal_timeout = 1.0

    assert bridge._ask_flip_batch([2], [1], [2], snapshot_minimal(copies=3), 1,
                                  copy_label="第 2/3 份 · ") is True
    payload = posts[-1][1]
    assert payload["title"] == "手动双面 · 请翻面（第 2/3 份）"
    assert payload["detail"][0].startswith("第 2/3 份 · 本批")

    bridge._batch_done_modal(4, 2, snapshot_minimal(), done=1, total=3, final=False)
    payload = posts[-1][1]
    assert payload["title"] == "第 1/3 份已打完，请取出"
    assert payload["proceed_label"] == "已取出，继续打印下一份"
    assert payload["sample"] == "剩余 2 份待打印。"

    bridge._batch_done_modal(4, 2, snapshot_minimal(), done=3, total=3, final=True)
    payload = posts[-1][1]
    assert payload["title"] == "全部完成，请取出全部纸张"
    assert payload["proceed_label"] == "已取出，全部完成"


def test_confirm_uncertain_phase_and_stop_helpers(bridge, monkeypatch):
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    posts = []
    answer = {"proceed": True}
    monkeypatch.setattr(bridge, "_post", answering_post(bridge, posts, answer))
    bridge._modal_timeout = 1.0

    bridge._confirm_uncertain_phase("阶段", "队列查询失败")
    assert posts[-1] == ("phase-confirm", {"text": "阶段", "reason": "队列查询失败"})

    answer["proceed"] = False
    with pytest.raises(CancelledByUser, match="用户取消"):
        bridge._confirm_uncertain_phase("阶段", "状态未知")

    # stop/pause 双状态都必须能打断等待。
    bridge.stop_event.clear()
    bridge.pause_event.clear()
    bridge._raise_if_stopped()
    bridge.stop_event.set()
    with pytest.raises(CancelledByUser):
        bridge._raise_if_stopped()

    sleeps = []
    bridge.stop_event.clear()
    monkeypatch.setattr("time.sleep", lambda seconds: sleeps.append(seconds))
    bridge._sleep_interruptible(0)
    assert sleeps == []


# ============================================================ 文件与设置 API
def test_file_option_validation_and_bulk_range_helpers(bridge, tmp_path):
    rec = add_pdf(bridge, tmp_path / "opts.pdf")
    iid = rec["iid"]

    assert bridge.set_files_options([], "1", 1)["ok"] is False
    assert bridge.set_files_options([iid], "abc", 1)["ok"] is False
    assert bridge.set_files_options([iid], "1", 0)["ok"] is False
    assert bridge.set_files_options([iid], "1", 100)["ok"] is False
    # 目标记录已不存在时按“跳过”处理，而不是把整批设为失败。
    assert bridge.set_files_options(["missing"], "1", 1)["ok"] is True

    result = bridge.set_files_options([iid], "1-2", 2)
    assert result["ok"] is True
    assert result["files"][0]["range_text"] == "1-2"
    assert result["files"][0]["copies"] == 2
    assert bridge.vars["per_file"] is True
    assert bridge.vars["recent_ranges"][0] == "1-2"

    # set_file_options 必须等价于单文件版 set_files_options。
    assert bridge.set_file_options(iid, "2", 3)["ok"] is True
    assert bridge.files[iid]["range_text"] == "2"
    assert bridge.files[iid]["copies"] == 3

    bridge.apply_range_to_all("3")
    assert bridge.files[iid]["range_text"] == "3"
    bridge.clear_all_ranges()
    assert bridge.files[iid]["range_text"] == ""


def test_add_recent_directory_and_empty_path(bridge, tmp_path):
    folder = tmp_path / "batch"
    folder.mkdir()
    make_pdf(folder / "b.pdf")
    make_pdf(folder / "a.PDF")
    (folder / "skip.exe").write_bytes(b"not-supported")

    result = bridge.add_recent(str(folder))
    assert result["ok"] is True
    assert [os.path.basename(item["path"]) for item in result["added"]] == ["a.PDF", "b.pdf"]
    assert bridge.add_recent("")["ok"] is False


def test_show_in_explorer_and_refresh_files(bridge, tmp_path, monkeypatch):
    rec = add_pdf(bridge, tmp_path / "shown.pdf")
    iid = rec["iid"]
    opened = []
    monkeypatch.setattr("webapp.bridge.subprocess.Popen", lambda args: opened.append(args))
    assert bridge.show_in_explorer(iid)["ok"] is True
    assert opened == [["explorer", "/select,", os.path.normpath(rec["path"])]]
    assert bridge.show_in_explorer("missing")["ok"] is False

    # 正常刷新：重新读取源 PDF 页数并清除历史错误。
    bridge.files[iid]["pages"] = 99
    bridge.files[iid]["error"] = "旧错误"
    assert bridge.refresh_files()["ok"] is True
    assert bridge.files[iid]["pages"] == 3
    assert bridge.files[iid]["status"] == "就绪"
    assert "error" not in bridge.files[iid]

    # 源文件丢失：标记失败而不是继续保留可打印状态。
    os.unlink(rec["path"])
    bridge.refresh_files()
    assert bridge.files[iid]["status"] == "失败"
    assert bridge.files[iid]["error"] == "文件不存在"

    # 损坏 PDF：读取页数异常也要标记失败。
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"broken")
    bridge.files["broken"] = {"iid": "broken", "kind": "pdf", "path": str(broken),
                              "source": str(broken), "pages": 1, "status": "就绪"}
    bridge.refresh_files()
    assert bridge.files["broken"]["status"] == "失败"
    assert bridge.files["broken"]["error"]

    # 非 PDF（图片/Office）不参与页数刷新。
    bridge.files["img"] = {"iid": "img", "kind": "image", "path": "missing.png",
                           "source": "missing.png", "pages": 0, "status": "失败"}
    bridge.refresh_files()
    assert bridge.files["img"]["status"] == "失败"


def test_output_directory_recent_and_printer_api(bridge, tmp_path, monkeypatch):
    bridge._window = SimpleNamespace(
        create_file_dialog=lambda *args, **kwargs: (str(tmp_path),))
    assert bridge.choose_output_directory()["path"] == str(tmp_path)
    bridge._window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: None)
    assert bridge.choose_output_directory()["path"] == ""

    def fail(*_args, **_kwargs):
        raise RuntimeError("dialog failure")

    bridge._window = SimpleNamespace(create_file_dialog=fail)
    assert bridge.choose_output_directory()["ok"] is False
    bridge._window = None
    assert bridge.choose_output_directory()["ok"] is False

    posts = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: posts.append((kind, payload)))
    bridge.vars["recent_files"] = ["a.pdf"]
    assert bridge.clear_recent()["ok"] is True
    assert bridge.vars["recent_files"] == []
    assert bridge.get_recent() == {"ok": True, "files": [], "ranges": []}
    assert posts[-1] == ("recent-updated", [])

    monkeypatch.setattr("webapp.bridge.list_printers", lambda: ["P1"])
    monkeypatch.setattr("webapp.bridge.get_default_printer", lambda: "P1")
    assert bridge.list_printers() == {"ok": True, "printers": ["P1"], "default": "P1"}

    monkeypatch.setattr("webapp.bridge.list_printers",
                        lambda: (_ for _ in ()).throw(RuntimeError("spooler down")))
    failed = bridge.list_printers()
    assert failed["ok"] is False and "spooler down" in failed["message"]

    monkeypatch.setattr("webapp.bridge.get_printer_status",
                        lambda name: {"name": name, "state": "就绪"})
    monkeypatch.setattr("webapp.bridge.check_printer_ready", lambda name: None)
    assert bridge.printer_status("P1")["ready"] is True

    monkeypatch.setattr("webapp.bridge.check_printer_ready",
                        lambda name: (_ for _ in ()).throw(RuntimeError("打印机离线")))
    offline = bridge.printer_status("P1")
    assert offline["ok"] is True
    assert offline["ready"] is False
    assert "打印机离线" in offline["detail"]

    monkeypatch.setattr("webapp.bridge.get_printer_sources", lambda name: [("自动选择", 7)])
    assert bridge.paper_sources("P1")["sources"] == [("自动选择", 7)]
    assert bridge.vars["_paper_sources"] == [("自动选择", 7)]

    monkeypatch.setattr("webapp.bridge.get_printer_sources",
                        lambda name: (_ for _ in ()).throw(RuntimeError("no source")))
    empty = bridge.paper_sources("P1")
    assert empty["ok"] is True and empty["sources"] == [] and empty["error"] == "no source"


def test_maintenance_helpers(bridge, tmp_path, monkeypatch):
    logs = []
    monkeypatch.setattr(bridge, "_log",
                        lambda text, error=False: logs.append((text, error)))
    monkeypatch.setattr("webapp.bridge.cleanup_stale_temp_dirs", lambda: ["a", "b"])
    assert bridge.cleanup_stale()["removed"] == 2
    assert ("已清理 2 个上次残留的临时目录。", False) in logs

    monkeypatch.setattr("webapp.bridge.available_converters", lambda: {"word": True})
    assert bridge.converters_available()["converters"] == {"word": True}

    opened = []
    monkeypatch.setattr("webapp.bridge.open_directory", lambda path: opened.append(path))
    assert bridge.open_log_dir()["ok"] is True
    assert len(opened) == 1

    monkeypatch.setattr("webapp.bridge.open_directory",
                        lambda path: (_ for _ in ()).throw(RuntimeError("no shell")))
    assert bridge.open_log_dir()["ok"] is False


def test_export_log_writes_utf8_and_handles_cancel(bridge, tmp_path, monkeypatch):
    monkeypatch.setitem(sys.modules, "webview",
                        SimpleNamespace(FileDialog=SimpleNamespace(SAVE=1)))
    target = tmp_path / "日志.txt"
    bridge._window = SimpleNamespace(
        create_file_dialog=lambda *args, **kwargs: (str(target),))
    result = bridge.export_log("第一行\n第二行")
    assert result == {"ok": True, "path": str(target)}
    assert target.read_text(encoding="utf-8-sig") == "第一行\n第二行"

    bridge._window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: ())
    assert bridge.export_log("x") == {"ok": True, "cancelled": True}

    def fail(*_args, **_kwargs):
        raise RuntimeError("save dialog broken")

    bridge._window = SimpleNamespace(create_file_dialog=fail)
    assert bridge.export_log("x")["ok"] is False

    bridge._window = None
    assert bridge.export_log("x")["ok"] is False


# ------------------------------------------------------------- 单面计划（不得误报偶数面）
def test_compute_plan_single_sided_accepts_odd_page_files(bridge, tmp_path):
    """单面打印 2 个 3 页文件：原本会误报“每个文件的面数都必须先补为偶数”。"""
    add_pdf(bridge, tmp_path / "odd-a.pdf", pages=3)
    add_pdf(bridge, tmp_path / "odd-b.pdf", pages=3)
    result = bridge.compute_plan(opts(manual_duplex=False, nup=1))
    assert result["ok"] is True, result.get("message")
    summary = result["summary"]
    assert result["total_pages"] == 6
    assert summary["sheets"] == 6              # 单面：每个输出面一张纸
    assert summary["flip_count"] == 0
    assert "单面" in summary["note"]
    # 单面绝不能再给出双面页序。
    assert summary["fronts"] == [] and summary["backs"] == []


def test_compute_plan_single_sided_nup_copies_and_per_file(bridge, tmp_path):
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=3)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=3)
    bridge.set_files_options([first["iid"]], "1-3", 2)
    # 2 个 3 页文件按 2 合 1 → 每个文件 2 个输出面。
    single = bridge.compute_plan(opts(manual_duplex=False, nup=2, copies=3))
    assert single["ok"] is True, single.get("message")
    assert single["summary"]["sheets"] == (2 * 3) * 2      # (2 面 × 3 份) × 2 文件
    assert single["summary"]["flip_count"] == 0
    # 逐文件份数：第一份 2 份、第二份 1 份（仅逐文件流程生效）。
    per_file = bridge.compute_plan(opts(manual_duplex=False, nup=2,
                                        per_file=True, copies=1, print_flow="per_file"))
    assert per_file["ok"] is True, per_file.get("message")
    assert per_file["summary"]["sheets"] == 2 * 2 + 2 * 1
    assert second["effective_copies"] == 1


def test_compute_plan_duplex_still_pads_and_pairs(bridge, tmp_path):
    add_pdf(bridge, tmp_path / "odd-a.pdf", pages=3)
    add_pdf(bridge, tmp_path / "odd-b.pdf", pages=3)
    result = bridge.compute_plan(opts(manual_duplex=True, nup=1))
    assert result["ok"] is True, result.get("message")
    assert result["summary"]["sheets"] == 4     # 两个 3 页文件各补到 4 面 → 8 面 = 4 张纸
    assert result["summary"]["flip_count"] == 1


# ------------------------------------------------------------- 横版预览与实际合成
def test_preview_landscape_paper_and_scale_match_composition(bridge, tmp_path):
    rec = add_pdf(bridge, tmp_path / "land.pdf", pages=3)
    options = opts(orientation="landscape", nup=2, manual_duplex=True)
    result = bridge.preview_page(rec["iid"], 1, options)
    assert result["ok"] is True, result.get("message")
    assert result["paper"] == "297×210 mm"
    assert result["landscape"] is True
    assert result["paper_pt"][0] > result["paper_pt"][1]
    # 真实合成页面尺寸必须与预览文案完全一致。
    resolved = bridge._record_pages(bridge.files[rec["iid"]], options)
    with __import__("tempfile").TemporaryDirectory() as temp_dir:
        out = os.path.join(temp_dir, "compose.pdf")
        composed = __import__("pdf_layout").compose_layout_pdf(
            bridge.files[rec["iid"]]["path"], resolved["valid"], out, resolved["layout"])
        doc = fitz.open(out)
        try:
            page = doc[0]
            assert result["paper_pt"] == pytest.approx(
                [page.rect.width, page.rect.height], abs=0.05)
            assert page.rect.width > page.rect.height
            # 预览缩放比例 = 排版时实测比例（不是按纸张尺寸估算的值）。
            assert result["scale_percent"] == round(composed.side_scales[0][0])
            assert result["nup"] == 2
        finally:
            doc.close()


@pytest.mark.parametrize("nup", (2, 4))
@pytest.mark.parametrize("orientation", ("landscape", "portrait"))
def test_preview_paper_orientation_matrix(bridge, tmp_path, nup, orientation):
    rec = add_pdf(bridge, tmp_path / f"p-{nup}-{orientation}.pdf", pages=4)
    result = bridge.preview_page(rec["iid"], 1, opts(
        orientation=orientation, nup=nup, manual_duplex=True))
    assert result["ok"] is True, result.get("message")
    assert result["landscape"] == (orientation == "landscape")
    assert result["paper"] == ("297×210 mm" if orientation == "landscape" else "210×297 mm")


def test_preview_single_page_landscape_on_portrait_paper_keeps_direction(bridge, tmp_path):
    """单页横版内容 + 竖版纸：预览/合成都不旋转，只缩放居中留白。"""
    path = tmp_path / "wide-preview.pdf"
    doc = fitz.open()
    page = doc.new_page(width=842, height=595)
    page.insert_text((40, 90), "WIDE-PREVIEW-MARK", fontsize=24)
    doc.save(str(path))
    doc.close()
    rec = bridge.add_files([str(path)])["added"][0]
    options = opts(nup=1, orientation="portrait", auto_rotate=True,
                   manual_duplex=False)
    result = bridge.preview_page(rec["iid"], 1, options)
    assert result["ok"] is True, result.get("message")
    assert result["paper"] == "210×297 mm" and result["landscape"] is False
    assert result["nup"] == 1
    # 预览与真实合成同源：合成页文字必须保持水平阅读方向（不能旋转 90°）。
    resolved = bridge._record_pages(bridge.files[rec["iid"]], options)
    with __import__("tempfile").TemporaryDirectory() as temp_dir:
        out = os.path.join(temp_dir, "compose.pdf")
        composed = __import__("pdf_layout").compose_layout_pdf(
            bridge.files[rec["iid"]]["path"], resolved["valid"], out, resolved["layout"])
        composed_doc = fitz.open(out)
        try:
            composed_page = composed_doc[0]
            dirs = [tuple(round(v, 2) for v in line["dir"])
                    for block in composed_page.get_text("dict")["blocks"]
                    for line in block.get("lines", [])]
            assert dirs and all(abs(d[1]) < 0.01 and d[0] > 0.99 for d in dirs)
            assert result["scale_percent"] == round(composed.side_scales[0][0])
        finally:
            composed_doc.close()


# ------------------------------------------------------------- 任务 19：生效范围与脱机状态
def test_compute_plan_writes_effective_range_for_each_file(bridge, tmp_path):
    """文件行显示的“当前实际生效范围”必须来自 compute_plan，而不是一直“全部”。"""
    rec = add_pdf(bridge, tmp_path / "range.pdf", pages=4)
    assert bridge.compute_plan(opts())["ok"] is True
    assert rec["effective_range"] == "全部" and rec["selected_pages"] == 4

    assert bridge.compute_plan(opts(range_mode="odd"))["ok"] is True
    assert rec["effective_range"] == "奇数页" and rec["selected_pages"] == 2

    assert bridge.compute_plan(opts(range_mode="even"))["ok"] is True
    assert rec["effective_range"] == "偶数页" and rec["selected_pages"] == 2

    assert bridge.compute_plan(opts(range_mode="custom", range_text="1-2"))["ok"] is True
    assert rec["effective_range"] == "1-2" and rec["selected_pages"] == 2

    assert bridge.compute_plan(opts(range_mode="current", current_page=3))["ok"] is True
    assert rec["effective_range"].startswith("当前页") and rec["selected_pages"] == 1


def test_compute_plan_per_file_range_wins_over_global(bridge, tmp_path):
    rec = add_pdf(bridge, tmp_path / "perfile.pdf", pages=4)
    bridge.set_files_options([rec["iid"]], "2", 1)
    assert bridge.compute_plan(opts(range_mode="all", per_file=True))["ok"] is True
    current = bridge.files[rec["iid"]]
    assert current["effective_range"] == "2" and current["selected_pages"] == 1


def test_printer_status_failure_reports_offline_with_actionable_detail(bridge, monkeypatch):
    import webapp.bridge as bridge_module

    def boom(_name):
        raise RuntimeError("DriverUnavailableError: 打印机驱动无响应，请重启打印机服务。")

    monkeypatch.setattr(bridge_module, "get_printer_status", boom)
    result = bridge.printer_status("OFFLINE-TEST")
    assert result["ok"] is True and result["ready"] is False
    assert result["status"]["state"] == "脱机"
    # 去掉 Python 异常类名前缀，保留可操作的中文说明。
    assert result["detail"].startswith("打印机驱动无响应")
    assert "Error" not in result["detail"]

    monkeypatch.setattr(bridge_module, "get_printer_status",
                        lambda _name: (_ for _ in ()).throw(ValueError("")))
    empty = bridge.printer_status("OFFLINE-TEST")
    assert empty["status"]["state"] == "脱机"
    assert empty["detail"] == "请检查打印机电源、USB/网络连接与纸张后重试。"


def test_compute_plan_returns_per_file_rows_for_frontend(bridge, tmp_path):
    """前端 state.files 是副本，compute_plan 必须回传 rows，范围文案才能即时刷新。"""
    rec = add_pdf(bridge, tmp_path / "rows.pdf", pages=4)
    result = bridge.compute_plan(opts(range_mode="odd"))
    rows = {row["iid"]: row for row in result["rows"]}
    row = rows[rec["iid"]]
    assert row["effective_range"] == "奇数页"
    assert row["selected_pages"] == 2
    assert row["effective_copies"] == 1
    assert row["plan_text"].endswith("面")
# -*- coding: utf-8 -*-


def make_landscape_pdf(path, pages=2, prefix="LAND"):
    doc = fitz.open()
    for index in range(pages):
        page = doc.new_page(width=842, height=595)
        page.insert_text((40, 70), f"{prefix}_{index + 1}", fontsize=20)
    doc.save(path)
    doc.close()


def add_landscape(bridge, path, pages=2, prefix="LAND"):
    make_landscape_pdf(path, pages, prefix)
    return bridge.add_files([str(path)])["added"][0]


def test_auto_landscape_first_page_plan_preview_and_devmode_match(bridge, tmp_path):
    """auto + ?????plan/metrics/preview/DEVMODE ???????"""
    rec = add_landscape(bridge, tmp_path / "land-first.pdf", pages=2)
    options = opts(orientation="auto", manual_duplex=False)

    resolved = bridge._record_pages(bridge.files[rec["iid"]], options)
    plan = resolved["plan"]
    assert plan.paper_pt[0] > plan.paper_pt[1]          # ?? 841.89?595.28
    assert plan.paper_pt[0] == pytest.approx(841.89, abs=0.05)

    described = bridge.plan_description(rec["iid"], options)
    assert described["metrics"]["paper"][0] > described["metrics"]["paper"][1]

    preview = bridge.preview_page(rec["iid"], 1, options)
    assert preview["ok"] is True
    assert preview["landscape"] is True
    assert preview["paper_pt"][0] == pytest.approx(plan.paper_pt[0], abs=0.05)
    assert preview["paper_pt"][1] == pytest.approx(plan.paper_pt[1], abs=0.05)

    snapshot = bridge._build_snapshot(options)
    printer_options = snapshot["printer_options"]
    assert snapshot["jobs"][0]["paper"] == pytest.approx(plan.paper_pt, abs=0.05)
    assert printer_options.landscape is True
    assert devmode_field_updates(printer_options)["dmOrientation"] == 2


def test_booklet_auto_plan_preview_and_devmode_are_landscape(bridge, tmp_path):
    """??? auto ???????? DEVMODE ????????dmOrientation=2??"""
    rec = add_pdf(bridge, tmp_path / "booklet.pdf", pages=8)
    options = opts(booklet=True, orientation="auto", manual_duplex=False)
    resolved = bridge._record_pages(bridge.files[rec["iid"]], options)
    assert resolved["plan"].paper_pt[0] > resolved["plan"].paper_pt[1]

    preview = bridge.preview_page(rec["iid"], 1, options)
    assert preview["ok"] is True and preview["landscape"] is True

    snapshot = bridge._build_snapshot(options)
    assert snapshot["printer_options"].landscape is True
    assert devmode_field_updates(snapshot["printer_options"])["dmOrientation"] == 2


def test_per_file_printer_options_follow_each_file_orientation(bridge):
    """?????????????????? PrinterOptions?DEVMODE ????"""
    base = PrinterOptions(landscape=False)
    landscape = bridge._printer_options_for_paper(base, (841.89, 595.28))
    portrait = bridge._printer_options_for_paper(base, (595.28, 841.89))
    assert landscape.landscape is True
    assert devmode_field_updates(landscape)["dmOrientation"] == 2
    assert portrait.landscape is False
    assert devmode_field_updates(portrait)["dmOrientation"] == 1
    # ????????????????
    assert base.landscape is False


def test_by_file_flow_passes_each_item_printer_options(bridge, tmp_path, monkeypatch):
    """?????????????? _print_phase?????????"""
    plan = build_layout_plan(2, LayoutOptions(manual_duplex=True))
    items = []
    for name, paper, landscape in (
            ("land", (841.89, 595.28), True),
            ("port", (595.28, 841.89), False)):
        items.append({
            "job": {"iid": name, "name": f"{name}.pdf", "copies": 1},
            "path": str(tmp_path / f"{name}.pdf"), "sides": 2, "plan": plan,
            "paper": paper,
            "printer_options": PrinterOptions(landscape=landscape),
        })
    captured = []
    monkeypatch.setattr(
        bridge, "_print_phase",
        lambda path, order, name, snap, text, **kwargs: captured.append(
            kwargs.get("printer_options")))
    monkeypatch.setattr(bridge, "_flip_modal", lambda *args, **kwargs: True)
    monkeypatch.setattr(bridge, "_file_done_modal", lambda *args, **kwargs: None)
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._run_by_file_duplex(
        {"copies": 1, "collate": True, "reverse": False, "manual_duplex": True}, items)

# ???????????????
    assert [options.landscape for options in captured] == [True, True, False, False]


# ------------------------------------------------------------- 任务 29：更多菜单 / 范围标签 / 作用对象
def test_range_label_for_single_page_file_never_shows_overflow_pages(bridge):
    """1 页文件请求 1,3,5,…,31 时只能显示有效结果，绝不回显越界页。"""
    text = "1,3,5,7,9,11,13,15,17,19,21,23,25,27,29,31"
    label = bridge._range_label("custom", text, valid=[1], total_pages=1)
    assert label == "全部"
    for page in ("3", "5", "31"):
        assert page not in label


def test_range_label_uses_effective_pages_for_odd_even_and_custom(bridge):
    assert bridge._range_label("odd", "", valid=[1, 3], total_pages=4) == "奇数页"
    assert bridge._range_label("even", "", valid=[2, 4], total_pages=4) == "偶数页"
    assert bridge._range_label("custom", "1-2", valid=[1, 2], total_pages=4) == "1-2"
    # 越界页已被 resolve_range_pages 丢弃：标签只列有效页。
    assert bridge._range_label("custom", "1,9,11", valid=[1], total_pages=8) == "1"
    assert bridge._range_label("custom", "1,3,5", valid=[1, 3, 5], total_pages=8) == "1,3,5"
    # 全部页码 / 全部奇数页依旧给出可读文案。
    assert bridge._range_label("custom", "1-4", valid=[1, 2, 3, 4], total_pages=4) == "全部"
    assert bridge._range_label(
        "current", "", 3, valid=[3], total_pages=4) == "当前页（第 3 页）"


def test_range_label_compacts_and_truncates_long_page_lists(bridge):
    assert bridge._format_range_pages([1, 2, 3, 5, 8]) == "1-3,5,8"
    long_label = bridge._format_range_pages([1, 3, 5, 7, 9, 11, 13, 15, 17, 19])
    assert long_label.startswith("1,3,5,7,9,11,13,15")
    assert long_label.endswith("…")


def test_compute_plan_single_page_custom_range_hides_overflow(bridge, tmp_path):
    """文件行 effective_range 不得出现越界页（1 页文件 + 1,3,…,31）。"""
    rec = add_pdf(bridge, tmp_path / "one.pdf", pages=1)
    result = bridge.compute_plan(opts(
        range_mode="custom",
        range_text="1,3,5,7,9,11,13,15,17,19,21,23,25,27,29,31"))
    assert result["ok"] is True
    row = {item["iid"]: item for item in result["rows"]}[rec["iid"]]
    assert row["effective_range"] == "全部"
    assert "31" not in row["effective_range"]
    assert row["selected_pages"] == 1


def test_move_files_moves_only_current_single_file(bridge, tmp_path):
    """默认全选时，上移/下移也只移动前端传入的那一个文件。"""
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=1)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=1)
    third = add_pdf(bridge, tmp_path / "c.pdf", pages=1)
    moved = bridge.move_files([second["iid"]], -1)["files"]
    assert [rec["iid"] for rec in moved] == [
        second["iid"], first["iid"], third["iid"]]
    # 已在最前时上移无效，其它文件顺序保持不变。
    at_top = bridge.move_files([second["iid"]], -1)["files"]
    assert [rec["iid"] for rec in at_top] == [
        second["iid"], first["iid"], third["iid"]]
    down = bridge.move_files([second["iid"]], 1)["files"]
    assert [rec["iid"] for rec in down] == [
        first["iid"], second["iid"], third["iid"]]


def test_dedupe_without_duplicates_reports_empty_removed(bridge, tmp_path):
    """无重复时必须返回空列表（前端据此提示“未发现重复文件”）。"""
    add_pdf(bridge, tmp_path / "a.pdf", pages=1)
    add_pdf(bridge, tmp_path / "b.pdf", pages=1)
    before = list(bridge.files)
    result = bridge.dedupe_files()
    assert result["ok"] is True
    assert result["removed"] == []
    assert list(bridge.files) == before


def test_set_files_options_default_scope_only_changes_target_file(bridge, tmp_path):
    """单独设置（scope=current）不得连带修改另一个文件。"""
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=4)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=4)
    result = bridge.set_files_options([first["iid"]], "1-2", 3)
    assert result["ok"] is True
    assert bridge.files[first["iid"]]["range_text"] == "1-2"
    assert bridge.files[first["iid"]]["copies"] == 3
    assert bridge.files[second["iid"]]["range_text"] == ""
    assert bridge.files[second["iid"]]["copies"] == 1


def test_set_files_options_selected_and_all_scopes(bridge, tmp_path):
    """scope=selected 命中已选集合；scope=all 覆盖全部文件；恢复全局同样按作用域。"""
    first = add_pdf(bridge, tmp_path / "a.pdf", pages=4)
    second = add_pdf(bridge, tmp_path / "b.pdf", pages=4)
    third = add_pdf(bridge, tmp_path / "c.pdf", pages=4)

    bridge.set_files_options([first["iid"], second["iid"]], "1", 2)
    assert bridge.files[first["iid"]]["range_text"] == "1"
    assert bridge.files[second["iid"]]["range_text"] == "1"
    assert bridge.files[third["iid"]]["range_text"] == ""

    bridge.set_files_options([first["iid"]], "2", 5, apply_all=True)
    assert all(rec["range_text"] == "2" and rec["copies"] == 5
               for rec in bridge.files.values())

    bridge.set_files_options([first["iid"]], "", 1, inherit=True, apply_all=True)
    assert all(not rec["range_text"] and rec["copies"] is None
               for rec in bridge.files.values())

# ------------------------------------------------- 任务 31：单文件范围无效隔离（O1）
def test_set_files_options_rejects_invalid_range_without_partial_write(bridge, tmp_path):
    """1 页文件设范围 2 必须被整体拒绝，且同批其它文件的 range_text 不变。"""
    one = add_pdf(bridge, tmp_path / "one.pdf", pages=1)
    many = add_pdf(bridge, tmp_path / "many.pdf", pages=16)
    assert bridge.set_files_options([one["iid"], many["iid"]], "1", 1)["ok"] is True

    # scope=current：只针对 1 页文件设范围 2（超出页数）→ 明确拒绝。
    rejected = bridge.set_files_options([one["iid"]], "2", 1)
    assert rejected["ok"] is False
    assert "《one.pdf》不包含第 2 页（共 1 页）" in rejected["message"]
    assert bridge.files[one["iid"]]["range_text"] == "1"
    assert bridge.files[many["iid"]]["range_text"] == "1"

    # scope=selected：其中一个文件范围无效时，整批不写入（合法文件也不变）。
    rejected = bridge.set_files_options([one["iid"], many["iid"]], "2-4", 1)
    assert rejected["ok"] is False
    assert "《one.pdf》不包含第 2-4 页（共 1 页）" in rejected["message"]
    assert bridge.files[one["iid"]]["range_text"] == "1"
    assert bridge.files[many["iid"]]["range_text"] == "1"

    # apply_range_to_all 同样先全量校验：任一无效则全部不改。
    assert bridge.apply_range_to_all("9-10")["ok"] is False
    assert bridge.files[one["iid"]]["range_text"] == "1"
    assert bridge.files[many["iid"]]["range_text"] == "1"

    # 语法错误也要给出可读的“范围格式无效”，而不是裸异常名。
    bad = bridge.set_files_options([many["iid"]], "abc", 1)
    assert bad["ok"] is False and bad["message"].startswith("范围格式无效：")
    assert bridge.files[many["iid"]]["range_text"] == "1"

    # 有效范围仍可正常写入，且不误伤同批文件。
    assert bridge.set_files_options([many["iid"]], "2-4", 1)["ok"] is True
    assert bridge.files[many["iid"]]["range_text"] == "2-4"
    assert bridge.files[one["iid"]]["range_text"] == "1"

    # 转换中/失败文件本就不可打印，页数不可信：不能因此拒绝整批设置。
    bridge.files[many["iid"]]["status"] = "失败"
    assert bridge.set_files_options([one["iid"], many["iid"]], "1", 1)["ok"] is True
    bridge.files[many["iid"]]["status"] = "就绪"


def test_compute_plan_isolates_one_invalid_range_and_keeps_other_rows(bridge, tmp_path):
    """全局 custom 2-4：1 页文件标“范围无效”，16 页文件照常出计划。"""
    one = add_pdf(bridge, tmp_path / "one.pdf", pages=1)
    many = add_pdf(bridge, tmp_path / "many.pdf", pages=16)

    result = bridge.compute_plan(opts(range_mode="custom", range_text="2-4"))
    assert result["ok"] is True
    # 部分成功仍要回传错误明细，供前端写一条可读日志。
    assert len(result["errors"]) == 1
    assert result["errors"][0]["name"] == "one.pdf"
    assert "范围没有有效页面" in result["errors"][0]["message"]

    rows = {row["iid"]: row for row in result["rows"]}
    assert len(rows) == 2
    bad, good = rows[one["iid"]], rows[many["iid"]]
    assert bad["effective_range"] == "范围无效"
    assert bad["selected_pages"] == 0
    assert bad["plan_text"] == "—"
    assert "范围没有有效页面" in bad["range_error"]
    # 合法文件的行必须照常刷新，不能因为同批有无效范围而停留在旧值。
    assert good["effective_range"] == "2-4"
    assert good["selected_pages"] == 3
    assert good["plan_text"].endswith("面")
    assert good["range_error"] is None
    assert result["total_pages"] == 3 and result["files"] == 1
    assert bridge.files[one["iid"]]["range_error"]
    assert not bridge.files[many["iid"]].get("range_error")

    # 修正范围后必须清除 range_error，行恢复为有效状态。
    assert bridge.set_files_options([one["iid"]], "", 1, inherit=True)["ok"] is True
    fixed = bridge.compute_plan(opts(range_mode="custom", range_text="1-2"))
    assert fixed["ok"] is True and fixed["errors"] == []
    assert "range_error" not in bridge.files[one["iid"]]
    assert bridge.files[one["iid"]]["effective_range"] == "全部"
    assert bridge.files[one["iid"]]["selected_pages"] == 1


def test_compute_plan_all_ranges_invalid_returns_rows_for_frontend(bridge, tmp_path):
    """全部文件范围无效：ok=false 仍要回传 rows/errors，前端才能把行标成无效。"""
    one = add_pdf(bridge, tmp_path / "one.pdf", pages=1)

    result = bridge.compute_plan(opts(range_mode="custom", range_text="2"))
    assert result["ok"] is False
    assert "所有文件的范围都无效" in result["message"]
    assert len(result["rows"]) == 1 and result["rows"][0]["iid"] == one["iid"]
    assert result["rows"][0]["range_error"]
    assert len(result["errors"]) == 1 and result["errors"][0]["name"] == "one.pdf"
