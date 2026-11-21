# -*- coding: utf-8 -*-
"""文件转换队列回归。"""

import os
import queue
import tempfile
import threading
import time
from types import SimpleNamespace

import pytest

import ui_files


class ConversionHarness(ui_files.FilesPanelMixin):
    def __init__(self, files):
        self.files = files
        self.ui_queue = queue.Queue()
        self.word_thread = None
        self.word_temp_dirs = {}
        self.shutting_down = False
        self.stop_event = threading.Event()


def test_file_added_during_conversion_is_processed(tmp_path, monkeypatch):
    first = tmp_path / "first.docx"
    second = tmp_path / "second.docx"
    first.write_bytes(b"first")
    second.write_bytes(b"second")
    first_record = {"iid": "first", "path": str(first), "kind": "word",
                    "pdf_path": "", "pages": None, "status": "待转换"}
    second_record = {"iid": "second", "path": str(second), "kind": "word",
                     "pdf_path": "", "pages": None, "status": "待转换"}
    harness = ConversionHarness([first_record])
    first_started = threading.Event()
    release_first = threading.Event()
    converted = []

    monkeypatch.setattr(ui_files, "available_converters", lambda: ["Word"])
    monkeypatch.setattr(ui_files, "create_temp_dir",
                        lambda: tempfile.mkdtemp(dir=tmp_path))

    def fake_convert(src, dst, timeout=180):
        converted.append(src)
        if src == str(first):
            first_started.set()
            if not release_first.wait(2):
                raise RuntimeError("timed out waiting to finish the first conversion")
        return {"converter": "Word", "pages": 1}

    monkeypatch.setattr(ui_files, "convert_file", fake_convert)
    harness._start_word_conversions()
    assert first_started.wait(2)

    harness.files.append(second_record)
    harness._start_word_conversions()
    release_first.set()

    completed = set()
    pumps = 0
    deadline = time.monotonic() + 5
    while pumps < 2 and time.monotonic() < deadline:
        try:
            item = harness.ui_queue.get(timeout=0.1)
        except queue.Empty:
            continue
        if item[0] == "word_done":
            record = first_record if item[1] == "first" else second_record
            record["pdf_path"] = item[2]
            record["pages"] = item[3]
            record["status"] = "Word（已转换为 PDF）"
            completed.add(item[1])
        elif item[0] == "pump_word":
            pumps += 1
            harness._start_word_conversions()

    assert converted == [str(first), str(second)]
    assert completed == {"first", "second"}
    assert pumps == 2


@pytest.mark.skipif(os.name != "nt", reason="透明 Tk 窗口回归仅在 Windows 桌面运行")
@pytest.mark.parametrize("scheme", ("stacked", "split"))
def test_resize_preserves_selection_text_and_button_spacing(tmp_path, monkeypatch, scheme):
    import tkinter.font as tkfont

    from ui_main import DuplexPrintApp, _create_root

    monkeypatch.setenv("APPDATA", str(tmp_path))
    root = _create_root()
    root.withdraw()
    root.tk.call("tk", "scaling", 1.3333)
    app = DuplexPrintApp(root, layout_scheme=scheme, ui_scale_override=1.0)
    app.preview_enabled = False
    for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
        poll_id = getattr(app, attr, None)
        if poll_id is not None:
            root.after_cancel(poll_id)
            setattr(app, attr, None)
    root.attributes("-alpha", 0.0)
    root.deiconify()
    try:
        for width, height in ((1024, 720), (1440, 720), (1024, 1100), (1440, 1100)):
            root.geometry(f"{width}x{height}+40+40")
            root.update()
            app._sync_file_panel_density()
            app._on_file_panel_configure()
            app._on_sel_bar_configure()
            app._on_preview_frame_configure()
            root.update()

            assert int(app.scope_hint_label.cget("wraplength")) == app.tab_files.winfo_width()
            assert int(app.preview_info_label.cget("wraplength")) == \
                app.preview_frame.winfo_width()
            assert int(app.sel_bar.grid_columnconfigure(0, "weight")) == 0
            assert int(app.sel_bar.grid_columnconfigure(len(app.sel_buttons), "weight")) == 1

            positions = [(button.winfo_x(), button.winfo_y(), button.winfo_width())
                         for button in app.sel_buttons]
            for left, right in zip(positions, positions[1:]):
                if left[1] == right[1]:
                    assert right[0] - left[0] - left[2] == app._px(ui_files.GAP_S)
            for button in app.sel_buttons:
                assert button.winfo_x() >= 0
                assert button.winfo_x() + button.winfo_width() <= app.sel_bar.winfo_width()

            if height == 1100:
                text = "选择文件后可单独设置范围与份数"
                font = tkfont.Font(root=root, font=app.sel_info_label.cget("font"))
                assert app.sel_info_label.winfo_ismapped()
                assert app.sel_info_label.grid_info()["columnspan"] == len(app.sel_buttons) + 1
                assert app.sel_info_label.winfo_width() >= font.measure(text)
    finally:
        app.shutting_down = True
        root.destroy()


@pytest.mark.skipif(os.name != "nt", reason="透明 Tk 窗口回归仅在 Windows 桌面运行")
@pytest.mark.parametrize("scheme", ("stacked", "split"))
@pytest.mark.parametrize("ui_scale", (1.0, 1.5))
def test_log_sash_resize_keeps_workspace_usable(tmp_path, monkeypatch, scheme, ui_scale):
    from ui_main import DuplexPrintApp, _create_root

    monkeypatch.setenv("APPDATA", str(tmp_path))
    root = _create_root()
    root.withdraw()
    root.tk.call("tk", "scaling", 1.3333 * ui_scale)
    app = DuplexPrintApp(root, layout_scheme=scheme, ui_scale_override=ui_scale)
    app.preview_enabled = False
    for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
        poll_id = getattr(app, attr, None)
        if poll_id is not None:
            root.after_cancel(poll_id)
            setattr(app, attr, None)
    root.attributes("-alpha", 0.0)
    root.deiconify()
    try:
        root.geometry(f"{round(1024 * ui_scale)}x{round(720 * ui_scale)}+40+40")
        root.update()
        assert app.duplex_mode_var.get() == "duplex"
        assert app.vars["manual_duplex"].get() is True

        app._on_log_sash_press(SimpleNamespace(y_root=600))
        app._on_log_sash_drag(SimpleNamespace(y_root=-1000))
        root.update()

        assert app._log_height > app._log_min_height
        assert app._log_height <= root.winfo_height() // 3
        tree_area = app.tree.master
        assert tree_area.winfo_height() >= app._px(54)
        assert app.preview_canvas.winfo_height() >= app._px(48)
        for child in (tree_area, app.preview_canvas):
            parent_bottom = child.master.winfo_rooty() + child.master.winfo_height()
            assert child.winfo_rooty() + child.winfo_height() <= parent_bottom
        assert app.preview_frame.winfo_rooty() + app.preview_frame.winfo_height() <= \
            app.workspace.winfo_rooty() + app.workspace.winfo_height()
        if scheme == "stacked":
            assert not app.recent_frame.winfo_ismapped()

        app._on_log_sash_press(SimpleNamespace(y_root=100))
        app._on_log_sash_drag(SimpleNamespace(y_root=1200))
        root.update()
        assert app._log_height == app._log_min_height
        if scheme == "stacked":
            assert app.recent_frame.winfo_ismapped()
        assert app.preview_frame.winfo_rooty() + app.preview_frame.winfo_height() <= \
            app.workspace.winfo_rooty() + app.workspace.winfo_height()
    finally:
        app.shutting_down = True
        root.destroy()
