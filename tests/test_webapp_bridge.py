# -*- coding: utf-8 -*-
"""pywebview 文件选择、导入与转换生命周期回归。"""
import inspect
import sys
import threading
from pathlib import Path
from types import SimpleNamespace

import fitz
import pytest
from PIL import Image

import settings_store
import window_geometry
from core_utils import CancelledByUser
from pdf_layout import LayoutOptions, build_layout_plan
from printer_core import PrinterOptions
from webapp.bridge import AppBridge, DaemonThreadPool


class ImmediateExecutor:
    """测试用同步执行器，让转换异常在断言前可见。"""

    def submit(self, fn, *args, **kwargs):
        fn(*args, **kwargs)


@pytest.fixture
def bridge(monkeypatch):
    monkeypatch.setattr(settings_store, "load_settings", lambda: {})
    monkeypatch.setattr(settings_store, "save_settings", lambda *args, **kwargs: None)
    instance = AppBridge()
    instance._executor.shutdown(wait=True)
    instance._executor = ImmediateExecutor()
    instance._log = lambda *_args, **_kwargs: None
    return instance


def make_pdf(path, pages=2):
    doc = fitz.open()
    for index in range(pages):
        page = doc.new_page(width=300, height=400)
        page.insert_text((30, 40), f"PAGE_{index + 1}")
    doc.save(path)
    doc.close()


def fake_webview(monkeypatch):
    module = SimpleNamespace(FileDialog=SimpleNamespace(OPEN=10, FOLDER=20))
    monkeypatch.setitem(sys.modules, "webview", module)
    return module


def test_choose_files_uses_native_dialog_and_returns_real_paths(bridge, monkeypatch):
    fake_webview(monkeypatch)
    calls = []

    class Window:
        def create_file_dialog(self, *args, **kwargs):
            calls.append((args, kwargs))
            return (r"D:\docs\a.pdf", r"D:\docs\b.pdf")

    bridge._window = Window()
    result = bridge.choose_files()
    assert result["ok"] is True
    assert result["paths"] == [r"D:\docs\a.pdf", r"D:\docs\b.pdf"]
    args, kwargs = calls[0]
    assert args == (10,)
    assert kwargs["allow_multiple"] is True
    assert "*.pdf" in kwargs["file_types"][0]


def test_choose_files_cancel_is_not_an_error(bridge, monkeypatch):
    fake_webview(monkeypatch)
    bridge._window = SimpleNamespace(create_file_dialog=lambda *_a, **_kw: None)
    assert bridge.choose_files() == {"ok": True, "paths": []}


def test_choose_folder_filters_supported_files(bridge, monkeypatch, tmp_path):
    fake_webview(monkeypatch)
    make_pdf(tmp_path / "b.pdf")
    make_pdf(tmp_path / "a.PDF")
    (tmp_path / "note.exe").write_bytes(b"not-an-import-format")
    (tmp_path / "nested").mkdir()
    bridge._window = SimpleNamespace(
        create_file_dialog=lambda *_a, **_kw: (str(tmp_path),))
    result = bridge.choose_folder()
    assert result["ok"] is True
    assert result["paths"] == [str(tmp_path / "a.PDF"), str(tmp_path / "b.pdf")]


def test_choose_folder_cancel_and_dialog_failure(bridge, monkeypatch):
    fake_webview(monkeypatch)
    bridge._window = SimpleNamespace(create_file_dialog=lambda *_a, **_kw: ())
    assert bridge.choose_folder() == {"ok": True, "paths": []}

    def fail(*_args, **_kwargs):
        raise RuntimeError("simulated dialog failure")

    bridge._window = SimpleNamespace(create_file_dialog=fail)
    result = bridge.choose_folder()
    assert result["ok"] is False
    assert "simulated dialog failure" in result["message"]


def test_add_files_reads_pdf_and_rejects_duplicate_invalid(bridge, tmp_path):
    path = tmp_path / "sample.pdf"
    make_pdf(path, pages=3)
    result = bridge.add_files([str(path), str(path), str(tmp_path / "bad.exe")])
    assert result["ok"] is True
    assert result["total"] == 1
    assert result["duplicates"] == 1
    assert result["ignored"] == 1
    assert result["added"][0]["pages"] == 3
    assert result["added"][0]["path"] == str(path)


def test_corrupt_pdf_does_not_create_unprintable_row(bridge, tmp_path):
    path = tmp_path / "broken.pdf"
    path.write_bytes(b"broken")
    result = bridge.add_files([str(path)])
    assert result["total"] == 0
    assert result["added"] == []
    assert result["ignored"] == 1


def test_image_conversion_output_survives_until_remove(bridge, tmp_path):
    source = tmp_path / "source.png"
    Image.new("RGB", (120, 80), "white").save(source)
    result = bridge.add_files([str(source)])
    assert result["total"] == 1
    rec = result["added"][0]
    assert rec["status"] == "就绪"
    assert rec["pages"] == 1
    converted = rec["path"]
    temp_dir = rec["temp_dir"]
    assert converted != str(source)
    assert fitz.open(converted).page_count == 1

    bridge.remove_files([rec["iid"]])
    assert not bridge.files
    from pathlib import Path
    assert not Path(temp_dir).exists()


def test_office_converter_receives_output_file_not_temp_directory(
        bridge, tmp_path, monkeypatch):
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    calls = []

    def convert(src, dst):
        calls.append((src, dst))
        make_pdf(dst, pages=4)
        return {"pages": 4}

    monkeypatch.setattr("webapp.bridge.convert_file", convert)
    result = bridge.add_files([str(source)])
    rec = result["added"][0]
    assert calls == [(str(source), rec["path"])]
    assert rec["status"] == "就绪"
    assert rec["pages"] == 4
    assert rec["path"].lower().endswith(".pdf")
    bridge.clear_files()
    assert not bridge.files


def test_recent_files_are_saved_and_broadcast(bridge, tmp_path, monkeypatch):
    path = tmp_path / "recent.pdf"
    make_pdf(path)
    saved = []
    monkeypatch.setattr(settings_store, "save_settings",
                        lambda settings, *args, **kwargs: saved.append(dict(settings)))
    events = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: events.append((kind, payload)))
    result = bridge.add_files([str(path)])
    assert result["total"] == 1
    assert saved[-1]["recent_files"] == [str(path)]
    assert ("recent-updated", [str(path)]) in events


def test_frontend_has_no_empty_file_path_or_directory_upload_fallback():
    from pathlib import Path
    root = Path(__file__).resolve().parents[1]
    script = (root / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
    html = (root / "webapp" / "static" / "index.html").read_text(encoding="utf-8")
    assert "choose_files" in script
    assert "choose_folder" in script
    assert 'postMessageWithAdditionalObjects("FilesDropped", files)' in script
    assert "webkitdirectory" not in script
    assert ".map((f) => f.path)" not in script
    assert '<img class="brand-icon" src="/app-icon.ico"' in html
    assert 'id="recent-select"' in html


def test_attach_does_not_wait_for_dom_and_close_stops_workers(bridge):
    class ClosedEvent:
        def __init__(self):
            self.handler = None

        def __iadd__(self, handler):
            self.handler = handler
            return self

    class Executor:
        def __init__(self):
            self.calls = []
            self.shutdown_calls = []

        def submit(self, fn, *args):
            self.calls.append((fn, args))

        def shutdown(self, **kwargs):
            self.shutdown_calls.append(kwargs)

    executor = Executor()
    bridge._executor = executor
    # 故意不提供 window.dom：启动前访问它会直接失败，防止恢复 20 秒等待路径。
    window = SimpleNamespace(events=SimpleNamespace(closed=ClosedEvent()))
    bridge.attach(window)
    assert [call[0].__name__ for call in executor.calls] == [
        "_drop_pump", "_apply_window_icon", "_restore_ime"]

    window.events.closed.handler()
    assert bridge._shutdown_event.is_set()
    assert bridge.stop_event.is_set()
    assert executor.shutdown_calls == [{"wait": False, "cancel_futures": True}]


def _shown_window():
    return SimpleNamespace(
        events=SimpleNamespace(shown=SimpleNamespace(wait=lambda _timeout: True)),
        title="PDF 批量双面打印工具",
    )


def test_restore_ime_respects_disable_and_selftest(bridge, monkeypatch):
    sent = []
    monkeypatch.setattr("webapp.bridge._send_single_shift", lambda: sent.append(1))
    bridge._window = _shown_window()

    monkeypatch.setenv("PDF_DUPLEX_RESTORE_IME", "0")
    bridge._restore_ime()
    assert sent == []

    monkeypatch.delenv("PDF_DUPLEX_RESTORE_IME")
    monkeypatch.setenv("WV_SELFTEST", "1")
    bridge._restore_ime()
    assert sent == []


def test_restore_ime_sends_once_only_for_own_foreground_window(bridge, monkeypatch):
    monkeypatch.delenv("PDF_DUPLEX_RESTORE_IME", raising=False)
    monkeypatch.delenv("WV_SELFTEST", raising=False)
    monkeypatch.setattr("webapp.bridge.time.sleep", lambda _seconds: None)
    sent = []
    monkeypatch.setattr("webapp.bridge._send_single_shift", lambda: sent.append(1))
    monkeypatch.setattr("webapp.bridge._is_own_window_visible", lambda _w: False)
    bridge._window = _shown_window()

    bridge._restore_ime()
    assert sent == []  # 窗口不可见时绝不往别的程序发按键

    monkeypatch.setattr("webapp.bridge._is_own_window_visible", lambda _w: True)
    bridge._restore_ime()
    assert sent == [1]  # 本次启动只发一次，不回读状态再重试


def test_drop_pump_returns_when_window_is_closed(bridge, monkeypatch):
    dom = SimpleNamespace(_dnd_state={"num_listeners": 1, "paths": []})
    monkeypatch.setitem(sys.modules, "webview", SimpleNamespace(dom=dom))
    monkeypatch.setitem(sys.modules, "webview.dom", dom)
    bridge._shutdown_event.set()
    bridge._drop_pump()  # 若关闭事件未生效，这里会永久循环。


# ============================================================ 前端契约回归
def frontend_sources():
    root = Path(__file__).resolve().parents[1]
    script = (root / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
    html = (root / "webapp" / "static" / "index.html").read_text(encoding="utf-8")
    return script, html


def test_acrobat_defaults_button_restores_full_backend_settings():
    """“恢复 Acrobat 默认”必须走后端 acrobat 分支，而不是只重置高级草稿。

    只调 resetAdvanced 会让 manual_duplex/range/copies/flow 停留在旧值，
    用户看到的“Acrobat 默认”与真实打印参数不一致；这里锁定完整链路。
    """
    script, html = frontend_sources()
    assert 'onclick="restoreAcrobatDefaults()"' in html
    # 原有“恢复默认”按钮必须保留。
    assert 'onclick="resetAdvanced(false)"' in html

    assert "window.restoreAcrobatDefaults = async function" in script
    start = script.index("window.restoreAcrobatDefaults")
    body = script[start:script.index("\n};", start)]
    assert "api().restore_defaults(true)" in body
    assert "applySettingsToUI(res.settings)" in body
    assert 'document.getElementById("advanced-summary").textContent = "高级设置：Acrobat 默认"' in body \
        or '$("advanced-summary").textContent = "高级设置：Acrobat 默认"' in body
    assert 'value="simplex"' in body
    # 失败分支必须先 return，保证弹窗仍打开；成功后才 hideModal。
    assert body.index("return;") < body.index('hideModal("modal-advanced")')
    assert "无法恢复默认设置" in body


def test_empty_state_round_plus_is_real_button_bound_to_add_files():
    """空状态圆形加号必须是 BUTTON，且初始化时绑定 addFiles（点击走原生对话框）。"""
    script, html = frontend_sources()
    empty_state = html[html.index('id="empty-state"'):html.index('id="selection-bar"')]
    assert 'class="empty-add"' in empty_state
    assert 'id="btn-empty-add"' in empty_state
    # 只有 BUTTON 才能被键盘/鼠标原生激活；回归为 div/span 会让加号失去可点击性。
    assert '<button class="empty-add" id="btn-empty-add" type="button"' in empty_state
    assert 'aria-label="添加文件"' in empty_state

    assert '$("btn-empty-add").onclick = addFiles;' in script
    assert "async function addFiles()" in script
    assert "await api().choose_files()" in script
    # 防止重新引入浏览器目录上传/空路径回退。
    assert "webkitdirectory" not in script


def test_summary_event_renders_into_summary_modal_and_can_be_closed():
    script, html = frontend_sources()
    summary_case = script[script.index('case "summary":'):script.index('};', script.index('case "summary":'))]
    assert '$("summary-title").textContent = payload.title' in summary_case
    assert 'fillTextLines($("summary-body")' in summary_case
    assert 'showModal("modal-summary")' in summary_case

    assert 'id="modal-summary"' in html
    assert 'id="summary-title"' in html
    assert 'id="summary-body"' in html
    # 汇总弹窗由用户手动关闭：关闭按钮必须绑定 hideModal。
    assert "hideModal('modal-summary')" in html


def test_flip_modal_event_reuses_shared_modal_and_updates_proceed_label():
    script, html = frontend_sources()
    flip_case = script[script.index('case "flip-modal":'):script.index('case "summary":')]
    assert '$("flip-title").textContent = payload.title;' in flip_case
    assert '$("btn-flip-confirm").textContent = payload.proceed_label' in flip_case
    assert "payload.detail" in flip_case
    assert "payload.sample" in flip_case
    assert "showModal(\"modal-flip\")" in flip_case

    assert 'id="modal-flip"' in html
    assert 'id="flip-title"' in html
    assert 'id="btn-flip-confirm"' in html


def test_persist_settings_sends_layout_preview_range_and_copy_exec():
    """设置变更必须把 nup/缩放/方向/预览/范围/逐份执行/延时统一回写后端。"""
    script, _html = frontend_sources()
    body = script[script.index("function persistSettings()"):script.index("function onSettingChange()")]
    assert "const changes = { ...state.advanced };" in body
    for key in ("printer", "copies", "manual_duplex", "print_flow", "range_mode",
                "range_text", "copy_exec_on", "preview_view"):
        assert f'"{key}"' in body, f"persistSettings 未回写 {key}"
    assert "api().save_settings(changes)" in body

    on_change = script[script.index("function onSettingChange()"):][:600]
    for assignment in ("state.advanced.scale_mode", "state.advanced.scale_percent",
                       "state.advanced.nup", "state.advanced.orientation"):
        assert assignment in on_change, f"onSettingChange 未同步 {assignment}"
    assert 'settings.delay_seconds' in script


def test_save_settings_persists_all_documented_fields(bridge, monkeypatch):
    saved = []
    monkeypatch.setattr(settings_store, "save_settings",
                        lambda settings: saved.append(dict(settings)))
    result = bridge.save_settings({
        "nup": 4, "scale_mode": "percent", "scale_percent": 135,
        "orientation": "landscape", "preview_view": "spread",
        "range_text": "1-3,5", "copy_exec_on": True, "delay": 9,
    })
    assert result["ok"] is True
    persisted = saved[-1]
    # 前端字段 delay 必须落盘为 delay_seconds，且不能残留旧键。
    assert persisted["delay_seconds"] == 9
    assert "delay" not in persisted
    assert persisted["nup"] == 4
    assert persisted["scale_mode"] == "percent"
    assert persisted["scale_percent"] == 135
    assert persisted["orientation"] == "landscape"
    assert persisted["preview_view"] == "spread"
    assert persisted["range_text"] == "1-3,5"
    assert persisted["copy_exec_on"] is True

    current = bridge.get_settings()["settings"]
    for key, expected in (("nup", 4), ("scale_mode", "percent"), ("scale_percent", 135),
                          ("orientation", "landscape"), ("preview_view", "spread"),
                          ("range_text", "1-3,5"), ("copy_exec_on", True),
                          ("delay_seconds", 9)):
        assert current[key] == expected, f"内存设置未更新 {key}"


# ============================================================ 打印结果汇总
def prepared_item(sides=2, copies=1, name="a.pdf"):
    return {"sides": sides,
            "job": {"iid": "1", "name": name, "copies": copies}}


def test_summary_payload_reports_completed_cancelled_and_failed(bridge):
    bridge._last_prepared = [prepared_item(sides=2, copies=2), prepared_item(sides=2)]

    completed = bridge._summary_payload(True, False, "",
                                        {"copies": 1, "manual_duplex": True})
    assert completed["title"] == "打印完成"
    assert completed["completed"] is True
    assert completed["files"] == 2
    assert completed["total_sides"] == 4
    # 手动双面按每 2 面 1 张，再乘逐文件份数：2 面×2 份 + 2 面×1 份 = 3 张。
    assert completed["sheets"] == 3
    assert "共 2 个文件" in completed["message"]
    assert "约 3 张纸" in completed["message"]

    cancelled = bridge._summary_payload(False, True, "用户取消。",
                                       {"copies": 1, "manual_duplex": True})
    assert cancelled["title"] == "已取消"
    assert "已取消" in cancelled["message"]

    failed = bridge._summary_payload(False, False, "printer offline",
                                     {"copies": 1, "manual_duplex": True})
    assert failed["title"] == "打印失败"
    assert "printer offline" in failed["message"]


# ============================================================ 纸张不一致确认
def worker_snapshot(tmp_path, count=2):
    jobs = []
    for index in range(1, count + 1):
        path = tmp_path / f"job-{index}.pdf"
        make_pdf(path, pages=2)
        jobs.append({"iid": str(index), "path": str(path), "pages": 2,
                     "abs_pages": [1, 2], "copies": 1, "name": f"job-{index}.pdf"})
    return {"printer": "SIMULATED", "jobs": jobs, "copies": 1,
            "per_file_settings": False, "collate": True,
            "manual_duplex": False, "print_flow": "batch_all",
            "copy_exec_on": False, "reverse": False, "wait_mode": "queue",
            "delay": 0, "keep_output": False, "auto_open": False,
            "output_dir": "", "layout": LayoutOptions(),
            "printer_options": SimpleNamespace()}


def patch_worker_layout(monkeypatch, papers):
    """用不同纸张尺寸模拟批量文件中排版纸张不一致。"""
    queue = list(papers)

    def fake_compose(src, pages, out, layout):
        Path(out).write_bytes(b"%PDF-1.4")
        plan = SimpleNamespace(output_count=2, paper_pt=queue.pop(0),
                               sides=[[0, 1], [2, None]])
        return SimpleNamespace(plan=plan)

    monkeypatch.setattr("webapp.bridge.compose_layout_pdf", fake_compose)


def test_paper_mismatch_prompts_with_flip_modal_and_continue_label(bridge, tmp_path, monkeypatch):
    snapshot = worker_snapshot(tmp_path)
    patch_worker_layout(monkeypatch, [(595.0, 842.0), (842.0, 595.0)])
    prompts = []
    monkeypatch.setattr(
        bridge, "_flip_modal",
        lambda title, detail, sample, proceed_label="": prompts.append(
            (title, detail, sample, proceed_label)) or True)
    ran = []
    monkeypatch.setattr(bridge, "_run_single_sided", lambda *args: ran.append(args))
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._worker_main(snapshot)

    assert len(prompts) == 1
    title, detail, sample, proceed_label = prompts[0]
    assert title == "纸张尺寸/方向不一致"
    assert detail[0] == "检测到各文件排版后的纸张尺寸或方向不一致："
    assert "《job-1.pdf》210×297mm" in detail[1]
    assert "《job-2.pdf》297×210mm" in detail[1]
    assert sample == "仍要继续整批打印吗？"
    assert proceed_label == "仍要继续整批打印"
    assert len(ran) == 1  # 用户确认后继续整批流程


def test_paper_mismatch_cancel_aborts_batch(bridge, tmp_path, monkeypatch):
    snapshot = worker_snapshot(tmp_path)
    patch_worker_layout(monkeypatch, [(595.0, 842.0), (842.0, 595.0)])
    monkeypatch.setattr(bridge, "_flip_modal", lambda *args, **kwargs: False)
    ran = []
    monkeypatch.setattr(bridge, "_run_single_sided", lambda *args: ran.append(args))
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    with pytest.raises(CancelledByUser, match="纸张不一致"):
        bridge._worker_main(snapshot)
    assert ran == []


def test_uniform_paper_does_not_prompt(bridge, tmp_path, monkeypatch):
    snapshot = worker_snapshot(tmp_path)
    patch_worker_layout(monkeypatch, [(595.0, 842.0), (595.0, 842.0)])
    prompts = []
    monkeypatch.setattr(bridge, "_flip_modal",
                        lambda *args, **kwargs: prompts.append(args) or True)
    monkeypatch.setattr(bridge, "_run_single_sided", lambda *args: None)
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._worker_main(snapshot)
    assert prompts == []


# ============================================================ 逐文件双面流程
def test_worker_main_runs_per_file_duplex_flow(bridge, tmp_path, monkeypatch):
    """逐文件手动双面必须真正执行，且正反面阶段数正确（防止调用签名回归）。"""
    snapshot = worker_snapshot(tmp_path)
    snapshot["manual_duplex"] = True
    snapshot["print_flow"] = "per_file"
    patch_worker_layout(monkeypatch, [(595.0, 842.0), (595.0, 842.0)])
    phases = []
    monkeypatch.setattr(
        bridge, "_print_phase",
        lambda path, order, name, snap, text, **kwargs: phases.append(
            (name, list(order), text)))
    monkeypatch.setattr(bridge, "_flip_modal", lambda *args, **kwargs: True)
    monkeypatch.setattr(bridge, "_file_done_modal", lambda *args, **kwargs: None)
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._worker_main(snapshot)

    assert [order for _name, order, _text in phases] == [[1], [2], [1], [2]]
    assert [text for _name, _order, text in phases] == [
        "逐文件 1/2 · 《job-1.pdf》正面", "逐文件 1/2 · 《job-1.pdf》反面",
        "逐文件 2/2 · 《job-2.pdf》正面", "逐文件 2/2 · 《job-2.pdf》反面"]


def test_worker_main_per_file_uses_each_paper_direction(bridge, tmp_path, monkeypatch):
    """逐文件流程：横版文件与竖版文件各自用实际纸向，不能整批静默竖版。"""
    snapshot = worker_snapshot(tmp_path)
    snapshot["manual_duplex"] = True
    snapshot["print_flow"] = "per_file"
    snapshot["printer_options"] = PrinterOptions(landscape=False)
    patch_worker_layout(monkeypatch, [(842.0, 595.0), (595.0, 842.0)])
    directions = []
    monkeypatch.setattr(
        bridge, "_print_phase",
        lambda path, order, name, snap, text, **kwargs: directions.append(
            kwargs["printer_options"].landscape))
    monkeypatch.setattr(bridge, "_flip_modal", lambda *args, **kwargs: True)
    monkeypatch.setattr(bridge, "_file_done_modal", lambda *args, **kwargs: None)
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._worker_main(snapshot)

    # 横版文件正/反两次为 True，竖版文件正/反两次为 False。
    assert directions == [True, True, False, False]


# ============================================================ 自检脚本回归
def test_selftest_script_verifies_frontend_contract_without_gui(bridge, monkeypatch, capsys):
    """用假窗口跑通 _selftest，锁定 SELFTEST_API 中关键前端契约。

    真实 WebView 自检仍由 WV_SELFTEST 运行；这里防止自检脚本本身回归成
    “永远为 true”或漏掉新增检查项。
    """
    import json as json_module

    monkeypatch.setattr("time.sleep", lambda *_args, **_kwargs: None)

    class FakeWindow:
        def __init__(self):
            self.resized = []
            self.preview_srcs = iter(["data:image/png;base64,AAA",
                                      "data:image/png;base64,BBB"])
            self.native = SimpleNamespace(Handle=SimpleNamespace(ToInt64=lambda: 12345))

        def resize(self, width, height):
            self.resized.append((width, height))

        def evaluate_js(self, code):
            if "JSON.stringify({metrics" in code:
                return json_module.dumps({
                    "metrics": 4, "brandIcon": True, "topbar": 65, "workspace": 294,
                    "printbar": 125, "logpanel": 160, "vh": 643, "vw": 1010,
                    "pageOverflow": False, "buttonsOkay": True, "fieldsOkay": True,
                    "fileButtonsSameRow": True, "selectionBarHidden": True, "emptyBoxFits": True,
                    "jsRan": True, "printers": 5,
                    "advBtn": True, "pauseBtn": True, "cancelBtn": True,
                    "planBtn": True, "progressBar": True})
            if ".src || ''" in code:
                return next(self.preview_srcs, "data:image/png;base64,ZZZ")
            if "cs.backgroundColor" in code:
                # 任务 33：自检读取 disabled “开始打印”的计算样式，必须是禁用灰。
                return json_module.dumps({
                    "disabled": True, "bg": "rgb(238, 240, 243)",
                    "color": "rgb(98, 109, 123)", "opacity": "1",
                    "border": "rgb(217, 222, 231)"})
            if "preview-view][value=spread]" in code:
                bridge.save_settings({"preview_view": "spread"})
            elif "preview-nup').value = '2'" in code:
                bridge.save_settings({"nup": 2})
            return True

    window = FakeWindow()
    bridge._window = window

    # 几何记忆：假窗口句柄无法真的查询 Win32，按 150% DPI 实测值桩接。
    monkeypatch.setattr("webapp.bridge.window_geometry.get_window_geometry",
                        lambda _hwnd: (1024, 680, 100, 80))
    monkeypatch.setattr("webapp.bridge.window_geometry.window_dpi_scale",
                        lambda _hwnd: 1.5)
    monkeypatch.setattr("webapp.bridge.window_geometry.list_work_areas",
                        lambda: [(0, 0, 2560, 1528)])

    from webview.platforms import edgechromium

    def fake_notify(_self, _ignored, event):
        bridge.add_files([item.Path for item in event.get_AdditionalObjects()])

    monkeypatch.setattr(edgechromium.EdgeChrome, "on_script_notify",
                        staticmethod(fake_notify))

    bridge._selftest()

    output = capsys.readouterr().out
    assert "SELFTEST_RESULT:" in output
    payload = json_module.loads(output.split("SELFTEST_API: ", 1)[1].splitlines()[0])
    for key in ("empty_add_button", "empty_selection_bar_hidden", "empty_state_fits",
                "selection_buttons_single_row", "summary_modal", "paper_confirm_modal",
                "settings_persist", "drop_path_correct", "advanced_apply",
                "acrobat_defaults",
                "geometry_save", "geometry_parse", "geometry_logical_size",
                "print_disabled_on_range_error", "print_disabled_grey_settled",
                "print_disabled_not_accent_blue",
                "print_disabled_uses_grey_palette",
                "print_enabled_after_range_fixed"):
        assert payload[key] is True, f"自检项 {key} 未通过"
    assert bridge.vars["window_geometry"] == "1024x680+100+80"
    assert window.resized == [(1024, 680)]
    # 自检结束必须清理文件表，避免残留临时文件。
    assert bridge.files == {}


# ============================================================ closing 语义
def test_on_window_closing_matches_pywebview_621_should_cancel_semantics(bridge, monkeypatch):
    """pywebview 6.2.1：closing.set() 同步执行处理器，返回 True 表示取消关闭。

    Event.set() 收集返回值，任一处理器返回 False 才视为“取消关闭”。
    因此处理器必须返回 False 阻止关闭、返回 True 放行关闭；且必须 0 参数同步返回。
    """
    from webview.event import Event

    class Window:
        def __init__(self, answer):
            self.answer = answer
            self.dialog_calls = 0

        def create_confirmation_dialog(self, *_args, **_kwargs):
            self.dialog_calls += 1
            if isinstance(self.answer, Exception):
                raise self.answer
            return self.answer

    def should_cancel(window):
        bridge._window = window
        event = Event(window, True)  # 与 pywebview Window.events.closing 一致
        event += bridge._on_window_closing
        return event.set()

    # pywebview 以 0 参数同步调用处理器；带参数会被传 window 而偏离真实语义。
    assert len(inspect.signature(bridge._on_window_closing).parameters) == 0

    # 未打印：放行关闭。
    assert bridge._printing is False
    assert should_cancel(Window(False)) is False

    # 打印中用户选择“不退出”：阻止关闭，且不得取消打印。
    bridge._printing = True
    window = Window(False)
    assert should_cancel(window) is True
    assert window.dialog_calls == 1
    assert bridge.stop_event.is_set() is False

    # 打印中用户确认退出：先屏蔽 JS，再取消任务，然后放行关闭。
    cancelled = []
    monkeypatch.setattr(bridge, "cancel_print", lambda: cancelled.append(True))
    bridge._suppress_js = False
    assert should_cancel(Window(True)) is False
    assert cancelled == [True]
    assert bridge._suppress_js is True

    # 对话框抛异常：为避免窗口无法关闭，允许退出。
    bridge._suppress_js = False
    monkeypatch.setattr("webapp.bridge.get_logger",
                        lambda: SimpleNamespace(exception=lambda *a, **k: None,
                                                warning=lambda *a, **k: None))
    assert should_cancel(Window(RuntimeError("dialog unavailable"))) is False

    # 对话框返回 None（无法显示）：保护打印任务，阻止关闭。
    bridge._suppress_js = False
    assert should_cancel(Window(None)) is True
    assert bridge.stop_event.is_set() is False


# ============================================================ 并发 / 超时 / 关闭回归
def default_print_opts(**changes):
    result = settings_store.default_settings()
    result.update(printer="SIMULATED", current_page=1)
    result.update(changes)
    return result


def test_start_print_reserves_slot_atomically_across_threads(bridge, tmp_path, monkeypatch):
    """两个并发 start_print 只能有一个成功；第二个必须立刻看到“已有任务”。"""
    path = tmp_path / "concurrent.pdf"
    make_pdf(path)
    assert bridge.add_files([str(path)])["total"] == 1
    monkeypatch.setattr("webapp.bridge.check_printer_ready", lambda name: None)

    entered = threading.Event()
    release = threading.Event()
    real_build = bridge._build_snapshot

    def slow_build(opts):
        entered.set()
        assert release.wait(3), "并发测试等待超时"
        return real_build(opts)

    monkeypatch.setattr(bridge, "_build_snapshot", slow_build)
    submitted = []
    bridge._executor = SimpleNamespace(
        submit=lambda fn, *args: submitted.append(fn))

    results = []
    worker = threading.Thread(target=lambda: results.append(
        bridge.start_print(default_print_opts())))
    worker.start()
    assert entered.wait(3), "第一个 start_print 未进入快照流程"

    # 第一个调用仍在校验阶段（尚未提交任务），占位已经生效。
    second = bridge.start_print(default_print_opts())
    release.set()
    worker.join(3)

    assert results and results[0]["ok"] is True
    assert second["ok"] is False and "已有打印任务" in second["message"]
    assert bridge._printing is True
    assert [fn.__name__ for fn in submitted] == ["_worker_entry"]


def test_start_print_releases_slot_after_snapshot_failure(bridge, tmp_path, monkeypatch):
    path = tmp_path / "release.pdf"
    make_pdf(path)
    bridge.add_files([str(path)])
    monkeypatch.setattr("webapp.bridge.check_printer_ready", lambda name: None)

    bad = bridge.start_print(default_print_opts(copies=0))
    assert bad["ok"] is False and "份数" in bad["message"]
    # 失败路径必须回滚占位，否则界面会永久认为“已有任务”。
    assert bridge._printing is False

    bridge._worker_entry = lambda snapshot: None
    assert bridge.start_print(default_print_opts())["ok"] is True
    assert bridge._printing is True


def test_flip_modal_timeout_resets_result_and_returns_false(bridge, monkeypatch):
    posts = []
    logs = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: posts.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log",
                        lambda text, error=False: logs.append((text, error)))
    bridge._modal_timeout = 0.05
    bridge._flip_result = True   # 上一轮“继续”的残留
    bridge._flip_event.set()     # 上一轮未清理的唤醒

    assert bridge._flip_modal("请翻面", ["详情"], "样张") is False
    assert bridge._flip_result is False
    assert bridge._flip_event.is_set() is False
    assert posts[-1][0] == "flip-modal"
    assert logs and logs[-1][1] is True


def test_flip_modal_returns_true_only_after_fresh_confirm(bridge, monkeypatch):
    posts = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: posts.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._modal_timeout = 1.0

    def confirm(kind, payload):
        posts.append((kind, payload))
        if kind == "flip-modal":
            bridge.confirm_flip(True)

    monkeypatch.setattr(bridge, "_post", confirm)
    bridge._flip_result = True  # 残留值不能直接当成本次确认
    assert bridge._flip_modal("请翻面", ["详情"], "样张") is True


def test_confirm_uncertain_phase_timeout_is_treated_as_cancel(bridge, monkeypatch):
    posts = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: posts.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    bridge._modal_timeout = 0.05
    bridge._phase_confirm_result = True  # 残留结果不能让超时被误判为确认
    bridge._phase_confirm_event.set()

    with pytest.raises(CancelledByUser, match="超时"):
        bridge._confirm_uncertain_phase("阶段", "队列状态未知")
    assert bridge._phase_confirm_result is False
    assert posts[-1][0] == "phase-confirm"


def test_worker_entry_logs_exception_and_still_emits_summary(bridge, monkeypatch):
    """打包后无 stderr 时，异常堆栈必须落到日志文件，且 finally 汇总不能丢。"""
    events = []
    monkeypatch.setattr(bridge, "_post", lambda kind, payload: events.append((kind, payload)))
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)
    recorded = []

    class Recorder:
        def exception(self, message, *args, **kwargs):
            recorded.append((message, sys.exc_info()[0]))

        def info(self, *args, **kwargs):
            pass

        def warning(self, *args, **kwargs):
            pass

    monkeypatch.setattr("webapp.bridge.get_logger", lambda: Recorder())

    def boom(snapshot):
        raise RuntimeError("boom")

    bridge._worker_main = boom
    bridge._printing = True
    bridge._worker_entry({"copies": 1, "manual_duplex": True})

    assert recorded and recorded[0][1] is RuntimeError
    assert [kind for kind, _payload in events] == ["print-done", "summary"]
    assert events[-1][1]["title"] == "打印失败"
    assert bridge._printing is False


def test_worker_main_checks_stop_before_compose(bridge, tmp_path, monkeypatch):
    """排版循环必须在每份 compose 前响应取消，避免关闭窗口后继续排版。"""
    snapshot = worker_snapshot(tmp_path)
    calls = []

    def fake_compose(src, pages, out, layout):
        calls.append(src)
        Path(out).write_bytes(b"%PDF-1.4")
        plan = SimpleNamespace(output_count=2, paper_pt=(595.0, 842.0),
                               sides=[[0, 1], [2, None]])
        return SimpleNamespace(plan=plan)

    def post(kind, payload):
        # 第二份“正在排版”后立刻取消：顶层检查已通过，只剩 compose 前的检查能拦住。
        if kind == "rows" and payload and payload[0][0] == "2":
            bridge.stop_event.set()

    monkeypatch.setattr("webapp.bridge.compose_layout_pdf", fake_compose)
    monkeypatch.setattr(bridge, "_post", post)
    monkeypatch.setattr(bridge, "_log", lambda *_args, **_kwargs: None)

    with pytest.raises(CancelledByUser):
        bridge._worker_main(snapshot)
    assert len(calls) == 1


def test_executor_workers_are_daemon_and_pool_rejects_after_shutdown(monkeypatch):
    monkeypatch.setattr(settings_store, "load_settings", lambda: {})
    instance = AppBridge()
    pool = instance._executor
    assert isinstance(pool, DaemonThreadPool)
    flags = []
    done = threading.Event()

    def task():
        flags.append(threading.current_thread().daemon)
        done.set()

    pool.submit(task)
    assert done.wait(3), "工作线程未执行"
    assert flags == [True]
    pool.shutdown(wait=True, cancel_futures=True)
    with pytest.raises(RuntimeError, match="已关闭"):
        pool.submit(task)


# ============================================================ 前端契约：最近/缩放/等待型弹窗
def test_render_recent_toggles_clear_button_state():
    script, html = frontend_sources()
    body = script[script.index("function renderRecent(paths)"):
                  script.index("async function refreshRecent()")]
    assert '$("btn-clear-recent").disabled = files.length === 0;' in body
    assert 'id="btn-clear-recent"' in html


def test_get_settings_scale_modes_use_custom_not_percent(bridge):
    schemes = bridge.get_settings()["schemes"]
    labels = {item["value"]: item["label"] for item in schemes["scale_modes"]}
    assert "custom" in labels
    assert "percent" not in labels
    assert labels["custom"] == "自定义缩放"


def test_frontend_wait_modals_always_report_back_on_close_or_cancel():
    script, _html = frontend_sources()
    assert ('window.flipAnswer = (proceed) => { hideModal("modal-flip"); '
            'api().confirm_flip(proceed); };') in script
    assert ('window.phaseAnswer = (proceed) => { hideModal("modal-phase"); '
            'api().confirm_phase(proceed); };') in script
    # Esc 关闭等待型弹窗必须回传取消，否则后端只能等超时。
    escape_block = script[script.index('e.key === "Escape"'):][:400]
    assert "flipAnswer(false)" in escape_block
    assert "phaseAnswer(false)" in escape_block


# ============================================================ 窗口几何记忆
# 实测环境：2560x1600 @150% DPI，任务栏 72px → 物理工作区 2560x1528，
# 逻辑工作区 1707x1019（2560/1.5、1528/1.5 四舍五入）。
WORK_AREAS_150 = [(0, 0, 2560, 1528)]


def test_window_geometry_parse_and_format_negative_offsets():
    assert window_geometry.parse_geometry("1024x680+120+80") == (1024, 680, 120, 80)
    assert window_geometry.parse_geometry("1024x680-20-40") == (1024, 680, -20, -40)
    assert window_geometry.parse_geometry("not-a-geometry") is None
    assert window_geometry.parse_geometry("") is None
    assert window_geometry.format_geometry(1024, 680, -20, 40) == "1024x680-20+40"


def test_geometry_dpi_conversion_matches_pywebview_logical_units():
    physical = (1800, 1200, 120, 60)
    logical = window_geometry.physical_geometry_to_logical(physical, 1.5)
    assert logical == (1200, 800, 80, 40)
    assert window_geometry.logical_geometry_to_physical(logical, 1.5) == physical
    # 物理矩形 → 逻辑几何；非法缩放（0）按 1.0 处理，不产生除零。
    assert window_geometry.rect_to_geometry((120, 60, 1920, 1260), 1.5) == logical
    assert window_geometry.physical_geometry_to_logical(physical, 0) == physical
    # 工作区同样按 DPI 换算（2560x1528 物理 → 1707x1019 逻辑）。
    assert window_geometry.logical_work_areas(WORK_AREAS_150, 1.5) == [
        (0, 0, 1707, 1019)]


def test_choose_logical_geometry_restores_visible_window():
    saved = window_geometry.format_geometry(1200, 800, 80, 67)
    chosen = window_geometry.choose_logical_geometry(
        saved, WORK_AREAS_150, 1.5, (1280, 860), (1024, 680))
    assert chosen == (1200, 800, 80, 67)
    areas = window_geometry.logical_work_areas(WORK_AREAS_150, 1.5)
    assert window_geometry.geometry_visible(chosen, areas)
    assert not window_geometry.geometry_visible((1200, 800, 6000, 4000), areas)


def test_choose_logical_geometry_falls_back_when_monitor_disconnected():
    # 记录来自已拔掉的外接屏（x=6000），必须回退主屏居中的默认尺寸。
    saved = window_geometry.format_geometry(1200, 800, 6000, 4000)
    chosen = window_geometry.choose_logical_geometry(
        saved, WORK_AREAS_150, 1.5, (1280, 860), (1024, 680))
    assert chosen == (1280, 860, 213, 79)
    # 无记录同样回退主屏居中。
    assert window_geometry.choose_logical_geometry(
        "", WORK_AREAS_150, 1.5, (1280, 860), (1024, 680)) == chosen


def test_choose_logical_geometry_clamps_partially_offscreen_window():
    # 标题栏仍可见但右下越界：夹回工作区，尺寸保持不变。
    saved = window_geometry.format_geometry(1200, 800, 600, 300)
    chosen = window_geometry.choose_logical_geometry(
        saved, WORK_AREAS_150, 1.5, (1280, 860), (1024, 680))
    assert chosen == (1200, 800, 507, 219)


def test_compose_window_geometry_formats_logical_geometry_for_bridge():
    assert AppBridge.compose_window_geometry(
        (1200, 800, 600, 300), WORK_AREAS_150, 1.5) == "1200x800+507+219"
    # 完全离屏 → 主屏居中回退。
    assert AppBridge.compose_window_geometry(
        (1200, 800, 6000, 5000), WORK_AREAS_150, 1.5) == "1280x860+213+79"
    # 小于最小尺寸的记录视为不可用，同样回退默认尺寸。
    assert AppBridge.compose_window_geometry(
        (800, 600, 40, 30), WORK_AREAS_150, 1.5) == "1280x860+213+79"


def test_save_window_geometry_persists_and_skips_unchanged(bridge, monkeypatch):
    writes = []
    monkeypatch.setattr(bridge, "_capture_window_geometry",
                        lambda: ((1200, 800, 80, 67), 1.5))
    monkeypatch.setattr(window_geometry, "list_work_areas",
                        lambda: list(WORK_AREAS_150))
    monkeypatch.setattr(settings_store, "save_settings",
                        lambda data: writes.append(dict(data)))

    assert bridge.save_window_geometry() is True
    assert bridge.vars["window_geometry"] == "1200x800+80+67"
    assert [item["window_geometry"] for item in writes] == ["1200x800+80+67"]
    # 几何未变化时不重复写盘。
    assert bridge.save_window_geometry() is True
    assert len(writes) == 1


def test_save_window_geometry_failure_logs_and_reverts(bridge, monkeypatch):
    errors = []
    monkeypatch.setattr(bridge, "_capture_window_geometry", lambda: None)
    assert bridge.save_window_geometry() is False

    monkeypatch.setattr(bridge, "_capture_window_geometry",
                        lambda: ((1200, 800, 80, 67), 1.5))
    monkeypatch.setattr(window_geometry, "list_work_areas",
                        lambda: list(WORK_AREAS_150))

    def failing_save(_data):
        raise settings_store.SettingsError("磁盘不可写")

    monkeypatch.setattr(settings_store, "save_settings", failing_save)
    monkeypatch.setattr("webapp.bridge.get_logger", lambda: SimpleNamespace(
        exception=lambda *args, **_kwargs: errors.append(args[0] if args else "")))

    assert bridge.save_window_geometry() is False
    assert errors and "保存窗口几何失败" in errors[0]
    # 失败后回滚内存值，关闭时还能再尝试一次。
    assert bridge.vars.get("window_geometry", "") == ""


def test_on_window_closing_saves_geometry_without_touching_js(bridge, monkeypatch):
    calls = []

    class Window:
        def evaluate_js(self, code):
            calls.append(("js", code))
            raise AssertionError("closing 回调禁止调用 evaluate_js")

    bridge._window = Window()
    monkeypatch.setattr(bridge, "save_window_geometry",
                        lambda: calls.append(("save", None)) or True)

    assert bridge._printing is False
    assert bridge._on_window_closing() is True
    assert calls == [("save", None)]


def test_window_geometry_events_debounce_and_stop_after_close(bridge, monkeypatch):
    tasks = []
    saves = []

    class Executor:
        def submit(self, fn, *args):
            tasks.append(fn)

    bridge._executor = Executor()
    bridge._geometry_save_delay = 0
    monkeypatch.setattr(bridge, "save_window_geometry",
                        lambda: saves.append(True) or True)

    bridge._on_window_geometry_event(10, 20)
    bridge._on_window_geometry_event(30, 40)  # 已有待处理任务，不重复排队
    assert [fn.__name__ for fn in tasks] == ["_flush_geometry_save"]
    tasks[0]()  # 模拟 daemon 线程执行去抖落盘
    assert saves == [True]

    # 落盘结束后可再次排队；关闭后不再排队（最终几何由 closing 保存）。
    bridge._on_window_geometry_event(50, 60)
    assert len(tasks) == 2
    bridge._geometry_closed = True
    tasks[1]()
    bridge._on_window_geometry_event(70, 80)
    assert len(tasks) == 2


def test_attach_subscribes_moved_and_resized_geometry_events(bridge):
    class Event:
        def __init__(self):
            self.handlers = []

        def __iadd__(self, handler):
            self.handlers.append(handler)
            return self

    class Executor:
        def submit(self, *_args, **_kwargs):
            return None

        def shutdown(self, **_kwargs):
            return None

    events = SimpleNamespace(closed=Event(), moved=Event(), resized=Event())
    bridge._executor = Executor()
    bridge.attach(SimpleNamespace(events=events))

    assert bridge._on_window_geometry_event in events.moved.handlers
    assert bridge._on_window_geometry_event in events.resized.handlers


# ============================================================ 第三轮 UI 反馈（静态契约）
def frontend_css():
    root = Path(__file__).resolve().parents[1]
    return (root / "webapp" / "static" / "style.css").read_text(encoding="utf-8")


def _css_rule(css, selector):
    """取出某个选择器（含其后紧跟的声明块）的文本，便于精确断言，避免全文误匹配。"""
    start = css.find(selector)
    assert start >= 0, selector
    end = css.find("}", start)
    assert end > start, selector
    return css[start:end + 1]


def test_selection_buttons_are_short_single_row():
    script, html = frontend_sources()
    rule = _css_rule(frontend_css(), ".selection-buttons {")
    assert "display: flex" in rule
    assert "flex-wrap: nowrap" in rule
    assert "grid-template-columns" not in rule
# 文案缩短，且确实是 4 个块级按钮同一行（只在选中操作条内断言）。
    bar = html[html.find('id="selection-bar"'):]
    bar = bar[:bar.find("</section>")]
    for label in ("范围…", "份数…", "恢复全局", "应用到全部"):
        assert f">{label}</button>" in bar
    assert "设置范围…" not in bar
    assert "设置份数…" not in bar
    assert "应用到全部文件" not in bar
    assert bar.count('id="btn-file-') == 4
    size_rule = _css_rule(frontend_css(), ".selection-buttons .btn.sm {")
    assert "height: 28px" in size_rule
    assert "padding: 3px 6px" in size_rule
    # 选中只由左侧复选框切换；行单击只预览（不改变选中）。
    assert "function toggleSelect(iid)" in script
    assert "check.onclick" in script and "toggleSelect(f.iid)" in script
    assert "row.onclick = () => { selectPreviewFile(f); };" in script
    assert "row.onclick = () => { toggleSelect(f.iid);" not in script


def test_file_row_click_previews_but_selection_only_follows_checkbox():
    """行单击不得改变选中：否则“已选中行再点取消”会回来。"""
    script, _html = frontend_sources()
    render = script[script.index("function renderFiles"):]
    render = render[:render.find("function toggleSelect")]
    assert ("check.onclick = (e) => { e.stopPropagation(); toggleSelect(f.iid); };"
            in render)
    assert "row.onclick = () => { selectPreviewFile(f); };" in render
    row_click = [line for line in render.splitlines() if "row.onclick" in line][0]
    assert "toggleSelect" not in row_click
    # 双击打开单文件设置、右键菜单行为必须保留；两者都固定作用于“当前文件”。
    assert ("row.ondblclick = () => { selectPreviewFile(f); "
            'openFileOpt(f.iid, "range", "current"); };' in render)
    assert 'openFileOpt(f.iid, "range", "current");' in render
    assert "row.oncontextmenu = (e) => {" in render


def test_range_text_shows_mode_hint_and_preserves_custom_text():
    """#range-text 非自定义模式显示说明，custom 原文切换模式后不能丢。"""
    script, html = frontend_sources()
    assert ('title="全局范围仅对未单独设置的文件生效；'
            '各文件可独立设置范围与份数"') in html
    assert "function updateRangeTextDisplay()" in script
    display = script[script.index("function updateRangeTextDisplay()"):]
    display = display[:display.find("function renderFiles")]
    assert '"各文件独立"' in display
    assert "逐文件（见左侧）" not in display
    assert "input.value = state.rangeText;" in display
    mode_fn = script[script.index("function rangeModeText(mode)"):
                      script.index("function updateRangeTextDisplay()")]
    for text in ("全部", "仅奇数页", "仅偶数页", "当前页"):
        assert text in mode_fn
    # 三个刷新入口都要统一更新 #range-text。
    sync = script[script.index("function syncInputStates()"):script.index("function persistSettings")]
    assert "updateRangeTextDisplay();" in sync
    metrics = script[script.index("async function refreshMetrics()"):script.index("function collectOptions")]
    assert "updateRangeTextDisplay();" in metrics
    render = script[script.index("function renderFiles"):script.index("function toggleSelect")]
    assert "updateRangeTextDisplay();" in render
    # 非 custom/recent 模式不能把提示文案当成范围文本持久化。
    collect = script[script.index("function collectOptions()"):script.index("function applySettingsToUI")]
    assert 'range_text: customMode ? $("range-text").value.trim() : state.rangeText.trim()' in collect


def test_drag_hint_uses_exactly_one_visual_layer():
    """拖入只允许一层高亮：空状态实线蓝框，有文件时列表内侧描边，禁止第二圈虚线。"""
    script, _html = frontend_sources()
    css = frontend_css()
    empty_rule = _css_rule(css, ".empty.drag-over {")
    assert "border: 2px solid var(--accent)" in empty_rule
    assert "background: #eff6ff" in empty_rule
    list_rule = _css_rule(css, ".file-list.drag-over {")
    assert "box-shadow: inset 0 0 0 2px var(--accent)" in list_rule
    assert "background: #eff6ff" in list_rule
    assert "dashed" not in list_rule
    drag = script[script.index("const showDropHint"):script.index('document.addEventListener("drop"')]
    assert 'dropEmpty.classList.toggle("drag-over", on && empty)' in drag
    assert 'dropZone.classList.toggle("drag-over", on && !empty)' in drag
    # 不允许再给 file-list 设 outline，否则空状态会同时出现第二圈虚线。
    assert "style.outline" not in script
    assert "relatedTarget === null" in script


def test_orientation_labels_are_short_and_backend_scheme_matches(bridge):
    """纸张方向只保留“自动/竖版/横版”，前后端标签一致。"""
    _script, html = frontend_sources()
    for label in (">自动</option>", ">竖版</option>", ">横版</option>"):
        assert label in html
    assert "自动（跟随首页）" not in html
    assert "竖版纸（高&gt;宽）" not in html
    labels = {item["value"]: item["label"]
              for item in bridge.get_settings()["schemes"]["orientations"]}
    assert labels == {"auto": "自动", "portrait": "竖版", "landscape": "横版"}


def test_advanced_auto_rotate_text_and_default_preserve_content_direction():
    """高级选项默认关闭并保持内容方向，用户仍可显式开启旋转。"""
    script, _html = frontend_sources()
    assert 'chk("auto_rotate", "方向不符时旋转 90° 适配槽位（默认关闭）", a.auto_rotate)' in script
    initial = script[script.index("state.advanced = {"):script.index("let advancedDraft")]
    assert "auto_rotate: false" in initial
    reset = script[script.index("window.resetAdvanced = function"):
                   script.index("window.restoreAcrobatDefaults")]
    assert "auto_rotate: !!_acrobat" in reset


def test_layout_options_auto_rotate_defaults_off_and_allows_opt_in(bridge):
    assert bridge._layout_options({}).auto_rotate is False
    assert bridge._layout_options({"auto_rotate": True}).auto_rotate is True


def test_round_plus_icon_is_thick_rounded_and_clickable():
    _script, html = frontend_sources()
    css = frontend_css()
    add_rule = _css_rule(css, ".empty-add {")
    assert "width: 30px" in add_rule and "height: 30px" in add_rule
    assert "border-radius: 50%" in add_rule
    assert "background: var(--accent)" in add_rule
    # 白加号：viewBox 24 内 5→19（臂长 14），stroke 3.4，圆头端点。
    assert 'd="M12 5v14M5 12h14"' in html
    assert 'stroke-width="3.4"' in html
    assert 'stroke-linecap="round"' in html
    assert 'stroke="#fff"' in html
    assert 'id="btn-empty-add"' in html and 'class="empty-add"' in html


def test_printer_select_has_bounded_width_and_full_name_titles():
    script, html = frontend_sources()
    rule = _css_rule(frontend_css(), "#printer-select {")
    assert "width: 220px" in rule
    assert "max-width: 240px" in rule
    assert "text-overflow: ellipsis" in rule
    assert 'id="printer-select"' in html
    # option 保留完整名称，title 展示全名（省略只发生在视觉层）。
    assert "opt.textContent = name" in script
    assert "sel.title = sel.value" in script
    assert 'label.title = label.textContent' in script
    assert '$( "printer-select" )' not in script


def test_do_add_and_drop_default_select_added_files():
    script, _html = frontend_sources()
    do_add = script[script.find("async function doAdd"):]
    do_add = do_add[:do_add.find("async function addFiles")]
    assert "state.selected.add(rec.iid)" in do_add
    assert "renderFiles();" in do_add and "refreshMetrics();" in do_add
    files_added = script[script.find('case "files-added"'):]
    files_added = files_added[:files_added.find('case "status"')]
    assert "state.selected.add(rec.iid)" in files_added
    # 转换完成只更新对象，不清空选中；预览同一 iid 时按当前页码重绘。
    file_update = script[script.find('case "file-update"'):]
    file_update = file_update[:file_update.find('case "files-added"')]
    assert "state.selected" not in file_update
    assert "state.previewFile.iid === payload.iid" in file_update
    assert "state.previewFile = payload" in file_update
    assert "renderPreview(Math.min(wanted, payload.pages))" in file_update


def test_empty_state_hides_selection_bar_and_fits_low_height():
    """1024×680 空状态：选中操作条必须隐藏，否则会覆盖圆形加号；低高度收紧空状态。"""
    script, _html = frontend_sources()
    css = frontend_css()
    render = script[script.find("function renderFiles"):]
    render = render[:render.find("function toggleSelect")]
    assert 'const selectionBar = $("selection-bar");' in render
    assert "if (!state.files.length)" in render
    assert "selectionBar.hidden = true;" in render
    assert "selectionBar.hidden = false;" in render
    assert ".selection-bar[hidden] { display: none; }" in css
    media = css[css.find("@media (max-height: 760px)"):]
    assert ".empty { padding: 8px; gap: 2px; }" in media
    assert ".empty-note { display: none; }" in media


# ============================================================ 任务 19：交互与文案契约
def test_no_native_browser_dialogs_in_frontend():
    """全部确认都必须走应用内模态框，前端不得残留系统原生提示。"""
    script, html = frontend_sources()
    for source in (script, html):
        assert "window.confirm" not in source
        assert "alert(" not in source
        assert "prompt(" not in source
    assert 'id="modal-confirm"' in html
    assert 'id="btn-confirm-ok"' in html and 'id="btn-confirm-cancel"' in html
    # 取消打印 / 批量移除 / 清空列表 / 恢复默认 共 4 处确认。
    assert script.count("await confirmAction({") == 4
    assert "window.answerConfirm" in script
    assert "answerConfirm(false)" in script  # Esc 取消


def test_selection_buttons_stretch_equal_width_without_scale_jitter():
    css = frontend_css()
    bar_rule = _css_rule(css, ".selection-buttons {")
    assert "width: 100%" in bar_rule
    assert "flex-wrap: nowrap" in bar_rule
    btn_rule = _css_rule(css, ".selection-buttons .btn.sm {")
    assert "flex: 1 1 0" in btn_rule
    assert "min-width: 0" in btn_rule
    # 圆形加号不得再用 transform 缩放：鼠标边缘进出会因命中区变化反复抖动。
    assert "transform: scale" not in css
    add_rule = _css_rule(css, ".empty-add {")
    assert "transform:" not in add_rule
    assert "transition: background .15s, box-shadow .15s" in add_rule


def test_output_radios_put_manual_duplex_first_and_default():
    _script, html = frontend_sources()
    manual = html.find('name="duplex" value="manual"')
    simplex = html.find('name="duplex" value="simplex"')
    assert 0 < manual < simplex, "输出选项顺序必须是“手动双面”在前、“单面”在后"
    assert 'name="duplex" value="manual" checked' in html


# ============================================================ 任务 29：更多菜单 / 范围标签 / 应用对象
def test_more_menu_operates_on_current_preview_file():
    """“更多”菜单只作用于当前预览文件（回退已选首个），不再吃整堆已选。"""
    script, _html = frontend_sources()
    helper = script[script.index("function currentFileIid()"):
                    script.index("function updateControlStates()")]
    assert "state.previewFile" in helper
    assert "state.files.some((f) => f.iid === state.previewFile.iid)" in helper
    assert "[...state.selected][0]" in helper

    move = script[script.index("const closeMore = () =>"):script.index('$("btn-dedupe").onclick')]
    assert "const iid = currentFileIid();" in move
    assert "api().move_files([iid], direction)" in move
    assert "api().move_files([...state.selected], direction)" not in move
    assert "已在最前" in move and "已在最后" in move
    assert "无法移动" in move

    explorer = script[script.index('$("btn-show-explorer").onclick'):
                      script.index('$("btn-file-range").onclick')]
    assert "const iid = currentFileIid();" in explorer
    assert "api().show_in_explorer(iid)" in explorer
    assert "api().show_in_explorer([...state.selected][0])" not in explorer
    assert "已在资源管理器中定位" in explorer
    assert "无法在资源管理器中定位" in explorer


def test_dedupe_logs_removed_count_or_no_duplicate_feedback():
    script, _html = frontend_sources()
    block = script[script.index('$("btn-dedupe").onclick'):
                   script.index('$("btn-show-explorer").onclick')]
    assert "res.removed" in block
    assert '"未发现重复文件"' in block
    assert "已移除 ${removed} 个重复文件" in block
    assert "去重失败" in block
    # 有重复时先刷新列表再提示数量。
    assert block.index("replaceFiles(res.files)") < block.index("已移除 ${removed}")


def test_move_and_explorer_enabled_with_preview_without_selection():
    """上移/下移/定位只要求存在当前文件，不要求多选。"""
    script, _html = frontend_sources()
    fn = script[script.index("function updateControlStates()"):
                script.index("function replaceFiles")]
    assert "const hasCurrent = !!currentFileIid();" in fn
    loop = fn[fn.index('["btn-move-up", "btn-move-down", "btn-show-explorer"]'):]
    loop = loop[:loop.index("}")]
    assert "$(id).disabled = !hasCurrent || state.printing;" in loop


def test_fileopt_modal_has_scope_radios_default_current():
    _script, html = frontend_sources()
    modal = html[html.index('id="modal-fileopt"'):html.index('id="modal-flip"')]
    assert 'name="fileopt-scope"' in modal
    for scope in ("current", "selected", "all"):
        assert f'id="fileopt-scope-{scope}"' in modal
        assert f'value="{scope}"' in modal
    assert 'id="fileopt-scope-count"' in modal
    current_tag = modal[modal.index('id="fileopt-scope-current"'):]
    current_tag = current_tag[:current_tag.index(">")]
    assert "checked" in current_tag
    # 默认作用对象必须是“当前文件”，否则列表默认全选会误改其它文件。
    assert not modal[modal.index('id="fileopt-scope-selected"'):].split(">")[0].count("checked")


def test_fileopt_save_resolves_scope_on_the_spot():
    script, _html = frontend_sources()
    assert 'function openFileOpt(iid, focus = "range", defaultScope = "current")' in script
    open_fn = script[script.index("function openFileOpt("):script.index("window.pickFileRange")]
    assert "fileoptTargetIid = iid;" in open_fn
    assert 'fileoptScope = FILEOPT_SCOPE_IDS[defaultScope] ? defaultScope : "current";' in open_fn

    resolver = script[script.index("function fileoptScopeIids()"):
                      script.index("function fileoptScopeDescription")]
    assert "return [...state.selected];" in resolver
    assert "return state.files.map((f) => f.iid);" in resolver
    assert "return fileoptTargetIid ? [fileoptTargetIid] : [];" in resolver

    save = script[script.index("async function saveFileOptions"):
                  script.index("async function applyFileOpt")]
    assert 'const iids = applyAll ? state.files.map((f) => f.iid) : fileoptScopeIids();' in save
    assert "await api().set_files_options(iids, range, copies, inherit, false)" in save

    # 行双击/右键、选中条范围/份数固定当前文件；“应用到全部”预选全部。
    assert 'openFileOpt(f.iid, "range", "current");' in script
    assert 'openFileOpt(currentFileIid(), "range", "current")' in script
    assert 'openFileOpt(currentFileIid(), "copies", "current")' in script
    assert 'openFileOpt(currentFileIid(), "range", "all")' in script
    assert "btn-fileopt-all" in script and "saveFileOptions({ applyAll: true })" in script

    ui = script[script.index("function updateFileOptScopeUI()"):
                script.index("window.setFileOptScope")]
    assert '`将应用到${desc}`' in ui
    assert "将应用到${desc}；留空使用全部页面" in ui


def test_global_range_hint_uses_independent_files_wording():
    """全局范围输入框的逐文件提示改为“各文件独立”，title 说明仅作用于未单独设置文件。"""
    script, html = frontend_sources()
    assert 'RANGE_DISPLAY_TEXTS = new Set(["全部", "各文件独立"' in script
    # 文件行/输入框的只读说明文案由前端常量提供（HTML 里只有 title 文案）。
    assert "各文件独立" in script
    assert "逐文件（见左侧）" not in script
    assert "逐文件（见左侧）" not in html
    assert "全局范围仅对未单独设置的文件生效；各文件可独立设置范围与份数" in html
    assert "全局范围仅对未单独设置的文件生效；各文件可独立设置范围与份数" in script


# ---------------------------------------- 任务 31：单文件范围无效的行内呈现（O1）
def test_range_error_row_is_red_and_disables_print():
    """f.range_error 必须在行内可见（红色“范围无效”）并禁用“开始打印”。"""
    script, _html = frontend_sources()

    render = script[script.index("function renderFiles"):
                    script.index("function toggleSelect")]
    assert 'const rangeError = f.range_error || "";' in render
    # 状态文本固定为“范围无效”，并带 err 类（红色）与错误原文 title。
    assert 'st.textContent = rangeError ? "范围无效" : f.status;' in render
    assert '(rangeError || f.status === "失败" ? " err" : "")' in render
    assert "if (rangeError) st.title = rangeError;" in render
    # 行内页数/范围也要跟着变成“范围无效”，不能沿用旧的有效范围。
    assert "pages.title = rangeError ||" in render

    controls = script[script.index("function updateControlStates()"):
                      script.index("function replaceFiles")]
    assert "const hasRangeError = state.files.some((f) => f.range_error);" in controls
    assert '$("btn-print").disabled = state.printing || !ready || hasRangeError;' in controls

    metrics = script[script.index("async function refreshMetrics()"):
                     script.index("function collectOptions")]
    # compute_plan 回传的 range_error 要合并进本地行（有效文件回传 null 即清除）。
    assert 'f.range_error = row.range_error || "";' in metrics
    # 存在无效范围时写一条“部分文件范围无效：…”日志，但其它行仍要重绘。
    assert "部分文件范围无效：" in metrics
    assert metrics.index("if (rows.length) renderFiles();") < metrics.index("部分文件范围无效：")

# ---------------------------------------- 任务 33：变体按钮禁用态必须真的变灰（O1 复现）
def test_disabled_variant_buttons_override_accent_blue():
    """disabled 的 accent/soft/danger 按钮必须走统一的灰色禁用态。

    根因：`.btn:disabled` 与 `.btn.accent/.btn.soft/.btn.danger` 同为 (0,2,0)
    特异性，且原规则写在变体之前，导致 disabled 的“开始打印”仍是 accent 蓝 + 白字。
    修复：在变体规则之后补一条同特异性组合选择器，靠层叠顺序覆盖。
    """
    css = frontend_css()
    rule = _css_rule(css, ".btn.accent:disabled, .btn.soft:disabled, .btn.danger:disabled {")
    assert "background: var(--disabled-bg)" in rule
    assert "color: var(--disabled-text)" in rule
    assert "border-color: var(--border)" in rule
    assert "cursor: default" in rule
    # 不能用 opacity 半透明冒充禁用，避免叠加出“仍可点击”的观感。
    assert "opacity" not in rule

    # 顺序断言：组合禁用规则必须晚于基础 .btn:disabled 与三个变体规则。
    unified = css.index(".btn.accent:disabled, .btn.soft:disabled, .btn.danger:disabled {")
    for earlier in (".btn:disabled {", ".btn.soft {", ".btn.accent {", ".btn.danger {"):
        assert css.index(earlier) < unified, earlier
    # 变体之后的 .btn.sm 等后续规则不得再改写 background/color。
    tail = css[unified + len(rule):]
    assert ".btn.accent {" not in tail
    assert ".btn.soft {" not in tail

    # 悬停/按下必须排除 disabled，否则鼠标移到禁用按钮上会把灰色改回蓝色。
    for selector in (".btn:hover:not(:disabled)", ".btn:active:not(:disabled)",
                     ".btn.soft:hover:not(:disabled)", ".btn.accent:hover:not(:disabled)",
                     ".btn.accent:active:not(:disabled)",
                     ".btn.danger:hover:not(:disabled)"):
        assert selector in css, selector
    for line in css.splitlines():
        if ".btn" in line and (":hover" in line or ":active" in line):
            assert ":not(:disabled)" in line, line

    # 前端确实在存在 range_error 时禁用开始打印（与视觉禁用态配套）。
    script, _html = frontend_sources()
    assert '$("btn-print").disabled = state.printing || !ready || hasRangeError;' in script


# ---------------------------------------- 任务 44：执行流程严格跟随界面选择
def print_opts(**changes):
    """构造打印参数：与前端 collectOptions 一致的关键字段。"""
    result = settings_store.default_settings()
    result.update(printer="SIMULATED", current_page=1, preview_view="single")
    result.update(changes)
    return result


def add_print_pdf(bridge, path, pages=3):
    """写入一个真实 PDF 并加入桥文件列表，返回新增记录。"""
    make_pdf(path, pages=pages)
    return bridge.add_files([str(path)])["added"][0]


def test_batch_flow_never_rewrites_ui_flow_and_uses_global_copies(bridge, tmp_path):
    """界面选“整批”时：流程不得被静默改成逐文件，所有文件用全局份数。

    复现问题：逐文件开关打开且某文件残留旧份数 1（全局 2）时，旧实现把
    snapshot.print_flow 偷偷改成 per_file，日志出现“逐文件 1/2”。
    """
    first = add_print_pdf(bridge, tmp_path / "a.pdf", pages=4)
    second = add_print_pdf(bridge, tmp_path / "b.pdf", pages=3)
    bridge.set_files_options([first["iid"]], "", 1)   # 文件级旧份数 = 1

    options = print_opts(copies=2, per_file=True, print_flow="batch_all")
    snapshot = bridge._build_snapshot(options)
    assert snapshot["print_flow"] == "batch_all"
    assert snapshot["copies"] == 2
    assert [job["copies"] for job in snapshot["jobs"]] == [2, 2]
    # 整批流程不启用逐文件设置（空字典为假值）。
    assert not snapshot["per_file_settings"]

    plan = bridge.compute_plan(options)
    assert plan["ok"] is True, plan.get("message")
    assert [(row["iid"], row["effective_copies"]) for row in plan["rows"]] == [
        (first["iid"], 2), (second["iid"], 2)]


def test_per_file_flow_uses_per_file_copies(bridge, tmp_path):
    """只有“逐文件执行”流程才采用文件级份数，并启用逐文件设置。"""
    first = add_print_pdf(bridge, tmp_path / "a.pdf", pages=4)
    add_print_pdf(bridge, tmp_path / "b.pdf", pages=3)
    bridge.set_files_options([first["iid"]], "", 1)

    options = print_opts(copies=2, per_file=True, print_flow="per_file")
    snapshot = bridge._build_snapshot(options)
    assert snapshot["print_flow"] == "per_file"
    copies_by_iid = {job["iid"]: job["copies"] for job in snapshot["jobs"]}
    assert copies_by_iid[first["iid"]] == 1
    assert snapshot["per_file_settings"][first["iid"]]["copies"] == 1

    plan = bridge.compute_plan(options)
    rows = {row["iid"]: row for row in plan["rows"]}
    assert rows[first["iid"]]["effective_copies"] == 1


@pytest.mark.parametrize("flow,label", [
    ("batch_all", "整批-最后取纸"),
    ("batch_take", "整批-逐文件取纸"),
    ("per_file", "逐文件执行"),
])
def test_worker_main_logs_execution_flow_matching_ui(bridge, tmp_path, monkeypatch,
                                                     flow, label):
    """启动日志必须明确本次执行流程，且文案与界面选择一致。"""
    snapshot = worker_snapshot(tmp_path, count=1)
    snapshot["print_flow"] = flow
    snapshot["manual_duplex"] = False
    logs = []
    monkeypatch.setattr(bridge, "_log", lambda text, error=False: logs.append(text))
    monkeypatch.setattr(bridge, "_run_single_sided", lambda *args: None)
    monkeypatch.setattr(bridge, "_post", lambda *args, **kwargs: None)

    bridge._worker_main(snapshot)

    assert [line for line in logs if line.startswith("执行流程：")] == [
        f"执行流程：{label}（与界面选择一致）"]


def test_confirm_uncertain_phase_logs_reason_before_modal(bridge, monkeypatch):
    """弹“打印状态待确认”前必须先写日志，否则日志里看不出弹窗原因。"""
    events = []

    def post(kind, payload):
        events.append(("post", kind, payload))
        bridge.confirm_phase(True)

    monkeypatch.setattr(bridge, "_post", post)
    monkeypatch.setattr(bridge, "_log",
                        lambda text, error=False: events.append(("log", text, error)))

    bridge._confirm_uncertain_phase("正面", "连续读取队列失败")

    assert events[0][0] == "log"
    assert "打印状态待确认" in events[0][1]
    assert "连续读取队列失败" in events[0][1]
    assert events[0][2] is True
    assert events[1][0] == "post" and events[1][1] == "phase-confirm"
