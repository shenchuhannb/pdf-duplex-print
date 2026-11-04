# -*- coding: utf-8 -*-
"""printer_core 纯函数回归：DEVMODE 字段、页序校验、任务状态解析、缩放几何。"""

import pytest

win32print = pytest.importorskip("win32print", reason="需要 pywin32")

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
