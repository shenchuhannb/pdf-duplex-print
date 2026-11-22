# -*- coding: utf-8 -*-
"""审计发现项的修复回归，不访问真实打印机、Office 或可见窗口。"""
import json
import os
from types import SimpleNamespace

import fitz
import pytest
from PIL import Image

import image_convert
import printer_core
import settings_store
import ui_worker
import word_convert
from core_utils import PdfError, PrintError
from pdf_layout import (
    LayoutOptions, build_layout_plan, compose_layout_pdf, compute_duplex_orders,
    crop_rect_for_page, preview_metrics, side_scale_percents,
)


def test_crop_rect_uses_current_cropbox_origin():
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    doc.xref_set_key(page.xref, "MediaBox", "[100 100 695 942]")
    doc.xref_set_key(page.xref, "CropBox", "[100 100 695 942]")
    options = LayoutOptions(crop_left_mm=10, crop_top_mm=10,
                            crop_right_mm=10, crop_bottom_mm=10)
    result = crop_rect_for_page(page, options)
    box = page.cropbox
    margin = 10 * 72 / 25.4
    assert result == fitz.Rect(box.x0 + margin, box.y0 + margin,
                               box.x1 - margin, box.y1 - margin)
    doc.close()


@pytest.mark.parametrize("rotation", (0, 90, 180, 270))
def test_crop_rotation_preserves_inner_rect_size(rotation):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.set_rotation(rotation)
    options = LayoutOptions(crop_left_mm=10, crop_top_mm=20,
                            crop_right_mm=30, crop_bottom_mm=40)
    result = crop_rect_for_page(page, options)
    expected_w = page.rect.width - 40 * 72 / 25.4
    expected_h = page.rect.height - 60 * 72 / 25.4
    if rotation in (0, 180):
        assert result.width == pytest.approx(expected_w)
        assert result.height == pytest.approx(expected_h)
    else:
        assert result.width == pytest.approx(expected_h)
        assert result.height == pytest.approx(expected_w)
    doc.close()


def test_annotations_are_flattened_into_visible_content(tmp_path):
    src = tmp_path / "annot.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.add_text_annot((100, 100), "VISIBLE_NOTE")
    doc.save(src)
    doc.close()
    hashes = {}
    for enabled in (False, True):
        out = tmp_path / f"out-{enabled}.pdf"
        compose_layout_pdf(str(src), [1], str(out),
                           LayoutOptions(manual_duplex=False, annots=enabled))
        doc = fitz.open(out)
        hashes[enabled] = doc[0].get_pixmap().samples
        doc.close()
    assert hashes[False] != hashes[True]


def test_preview_and_compose_use_same_nup_gap():
    options = LayoutOptions(manual_duplex=False, nup=2)
    plan = build_layout_plan(2, options)
    preview = preview_metrics(plan, options, 595, 842)
    actual = side_scale_percents(plan, options, lambda _rel: (595, 842))
    assert preview["scale_percent"] == pytest.approx(actual[0][0])


def test_invalid_nup_is_rejected_instead_of_dropping_pages():
    with pytest.raises(PdfError):
        build_layout_plan(3, LayoutOptions(manual_duplex=False, nup=3))


def test_single_side_copies_are_kept_and_odd_duplex_rejected():
    assert compute_duplex_orders(1, 3) == ([1, 1, 1], [])
    with pytest.raises(PdfError):
        compute_duplex_orders(5, 1)


def test_non_finite_bool_and_enum_settings_are_sanitized(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text(json.dumps({
        "dpi": float("nan"), "manual_duplex": "false",
        "orientation": "diagonal", "nup": 3,
    }), encoding="utf-8")
    result = settings_store.load_settings(str(path))
    assert result["dpi"] == 300
    assert result["manual_duplex"] is False
    assert result["orientation"] == "auto"
    assert result["nup"] == 1


def test_settings_merge_preserves_other_instance_changes(tmp_path):
    path = str(tmp_path / "settings.json")
    base = settings_store.default_settings()
    first = dict(base, dpi=600)
    settings_store.save_settings(first, path, base=base)
    second = dict(base, copies=3)
    settings_store.save_settings(second, path, base=base)
    merged = settings_store.load_settings(path)
    assert merged["dpi"] == 600
    assert merged["copies"] == 3


def test_corrupt_json_is_quarantined(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text("{broken", encoding="utf-8")
    result = settings_store.load_settings(str(path))
    assert result == settings_store.default_settings()
    assert not path.exists()
    assert len(list(tmp_path.glob("settings.json.corrupt-*"))) == 1


def test_owned_office_cleanup_ignores_unregistered_pids(tmp_path, monkeypatch):
    pid_file = tmp_path / "owners.txt"
    pid_file.write_text("101\n", encoding="ascii")
    monkeypatch.setattr(word_convert, "conversion_processes", lambda: {100, 101, 102})
    killed = []
    monkeypatch.setattr(word_convert, "_kill_pids", killed.extend)
    word_convert._terminate_owned_office(str(pid_file))
    assert killed == [101]
    assert not pid_file.exists()


def test_batch_rejects_duplicate_iid(tmp_path):
    items = [
        {"iid": "same", "src": str(tmp_path / "a.docx"), "dst": str(tmp_path / "a.pdf")},
        {"iid": "same", "src": str(tmp_path / "b.docx"), "dst": str(tmp_path / "b.pdf")},
    ]
    with pytest.raises(word_convert.WordConvertError):
        word_convert.convert_batch(items)


def test_word_macro_disable_can_be_required():
    class App:
        @property
        def AutomationSecurity(self):
            return 0

        @AutomationSecurity.setter
        def AutomationSecurity(self, _value):
            raise RuntimeError("unsupported")

    with pytest.raises(word_convert.WordConvertError):
        word_convert._disable_macros(App(), required=True)


def test_devmode_failure_never_falls_back_to_default(monkeypatch):
    class FailedSpool:
        def OpenPrinterW(self, *_args):
            return 0

        def ClosePrinter(self, _handle):
            pass

    create_dc_called = []
    monkeypatch.setattr(printer_core, "_spool", FailedSpool())
    monkeypatch.setattr(printer_core, "_gdi32", SimpleNamespace(
        CreateDCW=lambda *args: create_dc_called.append(args)))
    with pytest.raises(PrintError):
        printer_core._create_printer_dc("SIMULATED", printer_core.PrinterOptions())
    assert not create_dc_called


def test_job_status_combines_numeric_and_text_error():
    flags = printer_core._job_status_flags({"Status": 0, "pStatus": "Error"})
    assert flags & printer_core.win32print.JOB_STATUS_ERROR


def test_summary_counts_each_files_actual_copies(monkeypatch):
    harness = ui_worker.WorkerMixin()
    harness.root = None
    harness._ui_call = lambda fn: fn()
    messages = []
    monkeypatch.setattr(ui_worker.messagebox, "showinfo",
                        lambda title, message, **_kwargs: messages.append((title, message)))
    prepared = [
        {"sides": 2, "job": {"copies": 2}},
        {"sides": 4, "job": {"copies": 3}},
    ]
    harness._show_summary(True, False, "", prepared)
    assert "约 8 张纸" in messages[0][1]


def test_uncertain_phase_stops_without_confirmation():
    harness = ui_worker.WorkerMixin()
    harness._post = lambda _item: None
    harness._ui_call = lambda _fn: None
    harness.current_job_id = 0
    with pytest.raises(PrintError):
        harness._confirm_uncertain_phase("阶段一", "队列未知")


def test_word_failed_conversion_restores_previous_pdf(tmp_path, monkeypatch):
    src = tmp_path / "source.docx"
    src.write_bytes(b"docx")
    dst = tmp_path / "output.pdf"
    dst.write_bytes(b"previous-pdf")

    class FailedProcess:
        returncode = 2

        def communicate(self, timeout=None):
            return b"", b"simulated failure"

        def poll(self):
            return self.returncode

        def kill(self):
            pass

    monkeypatch.setattr(word_convert.subprocess, "Popen", lambda *args, **kwargs: FailedProcess())
    monkeypatch.setattr(word_convert, "_terminate_owned_office", lambda _path: None)
    with pytest.raises(word_convert.WordConvertError):
        word_convert.convert_file(str(src), str(dst))
    assert dst.read_bytes() == b"previous-pdf"
    assert not list(tmp_path.glob(".office_backup_*"))


def test_image_rejects_same_path_and_preserves_existing_output(tmp_path):
    src = tmp_path / "source.png"
    src.write_bytes(b"broken")
    with pytest.raises(image_convert.ImageConvertError):
        image_convert.image_to_pdf(str(src), str(src))
    dst = tmp_path / "output.pdf"
    dst.write_bytes(b"previous-pdf")
    with pytest.raises(image_convert.ImageConvertError):
        image_convert.image_to_pdf(str(src), str(dst))
    assert dst.read_bytes() == b"previous-pdf"


def test_image_frame_limit_is_enforced(tmp_path, monkeypatch):
    src = tmp_path / "many.tiff"
    src.write_bytes(b"x")
    fake = SimpleNamespace(n_frames=201, close=lambda: None)
    monkeypatch.setattr(image_convert.Image, "open", lambda _path: fake)
    with pytest.raises(image_convert.ImageConvertError):
        image_convert.image_to_pdf(str(src), str(tmp_path / "out.pdf"))


def test_batch_backs_uses_snapshot_copies_for_completion_modal():
    harness = ui_worker.WorkerMixin()
    harness._post = lambda _item: None
    harness._print_phase = lambda *args, **kwargs: None
    harness._ui_call = lambda _fn: "continue"
    harness.stop_event = SimpleNamespace(is_set=lambda: False)
    harness.pause_event = SimpleNamespace(is_set=lambda: False)
    prepared = [
        {"sides": 2, "path": "a.pdf", "job": {"iid": "a", "name": "a.pdf"}},
        {"sides": 2, "path": "b.pdf", "job": {"iid": "b", "name": "b.pdf"}},
    ]
    snapshot = {"copies": 2, "collate": True, "reverse": False}
    harness._print_batch_backs(snapshot, prepared, [(1, 2), (3, 4)], [4, 3, 2, 1],
                               "merged.pdf", pause_between_files=True)
    assert all(item["completed"] for item in prepared)


def test_frozen_worker_command_is_selected_by_target_executable():
    command = word_convert._worker_command(
        r"D:\App\PDFDuplexPrint.exe", "--src", "a.docx")
    assert command[1] == "--word-convert-worker"


def test_font_warning_reports_actual_export_mode():
    text = word_convert.font_warning_text("a.docx", ["MissingFont"],
                                          bitmap_missing_fonts=False)
    assert "普通 PDF 导出" in text
    assert "缺字体位图化”导出" not in text
