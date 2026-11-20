# -*- coding: utf-8 -*-
"""文件转换队列回归。"""

import queue
import tempfile
import threading
import time

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
