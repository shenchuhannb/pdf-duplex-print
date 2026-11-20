# -*- coding: utf-8 -*-
"""printer_core 纯函数回归：DEVMODE 字段、页序校验、任务状态解析、缩放几何。"""

import threading
import types

import pytest

win32print = pytest.importorskip("win32print", reason="需要 pywin32")

import printer_core
from core_utils import PrintError
from printer_core import (
    PrinterOptions, _fit_rect, _job_status_flags, devmode_field_updates,
    validate_side_order,
)

DMDUP_SIMPLEX = 1
DMNUP_ONEUP = 2


class TestDevmodeUpdates:
    """★ 安全线：任何选项组合都必须强制单面、单份、一页一版。"""

    def test_duplex_simplex_copies_one_nup_one(self):
        for quality in ("draft", "normal", "high"):
            updates = devmode_field_updates(PrinterOptions(print_quality=quality))
            assert updates["dmDuplex"] == DMDUP_SIMPLEX
            assert updates["dmCopies"] == 1
            assert updates["dmDisplayFlags"] == DMNUP_ONEUP

    def test_paper_code(self):
        assert devmode_field_updates(PrinterOptions(paper_name="A4"))["dmPaperSize"] == 9
        assert devmode_field_updates(PrinterOptions(paper_name="Letter"))["dmPaperSize"] == 1

    def test_custom_paper_dimensions(self):
        updates = devmode_field_updates(
            PrinterOptions(paper_name="自定义", paper_width_mm=100.0, paper_height_mm=200.0))
        # 0.1 毫米单位
        assert updates["dmPaperWidth"] == 1000
        assert updates["dmPaperLength"] == 2000

    def test_quality_mapping(self):
        assert devmode_field_updates(PrinterOptions(print_quality="draft"))["dmPrintQuality"] == -1
        assert devmode_field_updates(PrinterOptions(print_quality="high"))["dmPrintQuality"] == -4
        # normal 不覆盖驱动默认
        assert "dmPrintQuality" not in devmode_field_updates(PrinterOptions(print_quality="normal"))

    def test_color_mapping(self):
        assert devmode_field_updates(PrinterOptions(color_mode="mono"))["dmColor"] == 2
        assert devmode_field_updates(PrinterOptions(color_mode="gray"))["dmColor"] == 1


class TestValidateSideOrder:
    def test_valid(self):
        validate_side_order([1, 2, 3], 3)

    def test_out_of_range_rejected_before_submit(self):
        with pytest.raises(PrintError):
            validate_side_order([1, 5], 3)
        with pytest.raises(PrintError):
            validate_side_order([0, 1], 3)

    def test_empty_rejected(self):
        with pytest.raises(PrintError):
            validate_side_order([], 3)
        with pytest.raises(PrintError):
            validate_side_order([1], 0)


class TestJobStatusFlags:
    def test_int_status_passthrough(self):
        status = win32print.JOB_STATUS_PRINTING | win32print.JOB_STATUS_OFFLINE
        assert _job_status_flags({"Status": status}) == status

    def test_text_word_match(self):
        flags = _job_status_flags({"Status": None, "pStatus": "Error, Printing"})
        assert flags & win32print.JOB_STATUS_ERROR
        assert flags & win32print.JOB_STATUS_PRINTING

    def test_substring_no_false_positive(self):
        """厂商自定义文本包含 error 子串但不是一个词时，不能误判为任务错误。"""
        flags = _job_status_flags({"Status": None, "pStatus": "terrorlevel warning"})
        assert not (flags & win32print.JOB_STATUS_ERROR)


class TestFitRect:
    def test_contain_and_center(self):
        x, y, w, h = _fit_rect(100, 200, 1000, 1000)
        assert (w, h) == (500, 1000) and x == 250 and y == 0

    def test_no_upscale_distortion(self):
        x, y, w, h = _fit_rect(100, 100, 400, 200)
        assert abs(w / h - 1.0) < 0.01


class TestOfflineDetectionUnion:
    """★ 安全线：脱机判定必须取属性位与状态位的并集。

    HP P1007 等 host-based USB 驱动在设备关机/断连时往往只置
    PRINTER_ATTRIBUTE_WORK_OFFLINE 属性位，Status 位不更新；
    只信状态位会把脱机设备放行，任务堆积后楔死驱动会话。
    """

    def _fake_printer(self, monkeypatch, status=0, attributes=0):
        fake = types.SimpleNamespace(**{
            name: getattr(win32print, name)
            for name in dir(win32print) if name.isupper()
        })
        fake.OpenPrinter = lambda name: object()
        fake.ClosePrinter = lambda handle: None
        fake.GetPrinter = lambda handle, level: {
            "Status": status, "Attributes": attributes, "cJobs": 0,
        }
        monkeypatch.setattr(printer_core, "win32print", fake)
        monkeypatch.setattr(printer_core, "get_default_printer", lambda: "")

    def test_work_offline_attribute_marks_offline(self, monkeypatch):
        self._fake_printer(
            monkeypatch, status=0,
            attributes=win32print.PRINTER_ATTRIBUTE_WORK_OFFLINE)
        assert printer_core.get_printer_status("P1007")["offline"] is True

    def test_status_bit_marks_offline(self, monkeypatch):
        self._fake_printer(
            monkeypatch, status=win32print.PRINTER_STATUS_OFFLINE, attributes=0)
        assert printer_core.get_printer_status("P1007")["offline"] is True

    def test_ready_printer_not_offline(self, monkeypatch):
        self._fake_printer(monkeypatch, status=0, attributes=0)
        assert printer_core.get_printer_status("P1007")["offline"] is False


class TestValidatePrinterConnection:
    """驱动探测必须有超时兜底：host-based 驱动的阻塞调用不可中断。"""

    def _stub_ready(self, monkeypatch):
        monkeypatch.setattr(printer_core, "check_printer_ready", lambda name: None)

    def test_timeout_raises_readable_error(self, monkeypatch):
        self._stub_ready(monkeypatch)
        hang = threading.Event()
        monkeypatch.setattr(
            printer_core, "_create_printer_dc",
            lambda name, options: hang.wait(30))
        try:
            with pytest.raises(PrintError, match="超时"):
                printer_core.validate_printer_connection(
                    "P1007", PrinterOptions(), timeout=0.3)
        finally:
            hang.set()  # 释放探测线程，避免悬挂到测试进程退出

    def test_driver_error_wrapped(self, monkeypatch):
        self._stub_ready(monkeypatch)

        def boom(name, options):
            raise RuntimeError("driver broken")

        monkeypatch.setattr(printer_core, "_create_printer_dc", boom)
        with pytest.raises(PrintError, match="无法初始化打印机驱动"):
            printer_core.validate_printer_connection(
                "P1007", PrinterOptions(), timeout=5.0)

    def test_success_path(self, monkeypatch):
        self._stub_ready(monkeypatch)
        monkeypatch.setattr(
            printer_core, "_create_printer_dc",
            lambda name, options: (12345, None))
        fake_gdi = types.SimpleNamespace(
            GetDeviceCaps=lambda hdc, index: 100,
            DeleteDC=lambda hdc: None,
        )
        monkeypatch.setattr(printer_core, "_gdi32", fake_gdi)
        printer_core.validate_printer_connection(
            "P1007", PrinterOptions(), timeout=5.0)


class TestValidatePrecheckTimeout:
    """队列预检（OpenPrinter/GetPrinter 同为 spooler RPC）也必须在超时保护内。"""

    def test_hanging_precheck_times_out(self, monkeypatch):
        hang = threading.Event()
        monkeypatch.setattr(
            printer_core, "check_printer_ready", lambda name: hang.wait(30))
        try:
            with pytest.raises(PrintError, match="超时"):
                printer_core.validate_printer_connection(
                    "P1007", PrinterOptions(), timeout=0.3)
        finally:
            hang.set()

    def test_offline_precheck_keeps_exception_type(self, monkeypatch):
        from core_utils import PrinterOfflineError

        def offline(name):
            raise PrinterOfflineError("离线")

        monkeypatch.setattr(printer_core, "check_printer_ready", offline)
        with pytest.raises(PrinterOfflineError):
            printer_core.validate_printer_connection(
                "P1007", PrinterOptions(), timeout=5.0)


class TestWaitNeverObserved:
    """任务从未在队列中观察到：首次观察及时→报错；被 OpenPrinter 重试推迟→按完成兜底。"""

    def _fake_queue(self, monkeypatch, open_fail_times=0):
        state = {"opens": 0}

        def open_printer(name):
            state["opens"] += 1
            if state["opens"] <= open_fail_times:
                raise OSError("spooler busy")
            return object()

        def get_job(handle, job_id, level):
            raise OSError("job not found")

        fake = types.SimpleNamespace(**{
            name: getattr(win32print, name)
            for name in dir(win32print) if name.isupper()
        })
        fake.OpenPrinter = open_printer
        fake.ClosePrinter = lambda handle: None
        fake.GetPrinter = lambda handle, level: {
            "Status": 0, "Attributes": 0, "cJobs": 0,
        }
        fake.GetJob = get_job
        fake.SetJob = lambda *a: None
        monkeypatch.setattr(printer_core, "win32print", fake)
        monkeypatch.setattr(printer_core, "JOB_POLL_INTERVAL", 0.01)
        monkeypatch.setattr(printer_core, "RETRY_BACKOFF_SECONDS", 0.01)
        monkeypatch.setattr(printer_core, "PRINTER_STATUS_POLL_INTERVAL", 0.01)

    def test_prompt_observation_missing_raises(self, monkeypatch):
        self._fake_queue(monkeypatch)
        with pytest.raises(PrintError, match="未能在队列中观察到"):
            printer_core.wait_for_print_job(
                "P1007", 42, threading.Event())

    def test_delayed_observation_falls_back_to_complete(self, monkeypatch):
        self._fake_queue(monkeypatch, open_fail_times=1)
        assert printer_core.wait_for_print_job(
            "P1007", 42, threading.Event()) is True
