# -*- coding: utf-8 -*-
"""core_utils 纯函数回归：页码解析、拖拽数据解析、单位换算、文件类型判断。"""

import pytest

from core_utils import (
    PdfError, format_page_list, is_pdf_path, is_word_path, is_wps_path,
    mm_to_pt, parse_page_range, pt_to_mm, split_drop_data,
)


class TestParsePageRange:
    def test_basic(self):
        assert parse_page_range("1-5,8,10-12") == [1, 2, 3, 4, 5, 8, 10, 11, 12]

    def test_chinese_punctuation(self):
        # 中文逗号/顿号/分号/全角连字符都要兼容
        assert parse_page_range("1，3、5；8－10") == [1, 3, 5, 8, 9, 10]
        assert parse_page_range("2—4") == [2, 3, 4]
        assert parse_page_range("2～3") == [2, 3]

    def test_reversed_range_normalized(self):
        assert parse_page_range("5-3") == [3, 4, 5]

    def test_dedup_and_sorted(self):
        assert parse_page_range("3,1,3,2,1-2") == [1, 2, 3]

    def test_spaces_tolerated(self):
        assert parse_page_range(" 1 - 3 , 7 ") == [1, 2, 3, 7]

    def test_empty_rejected(self):
        with pytest.raises(PdfError):
            parse_page_range("")
        with pytest.raises(PdfError):
            parse_page_range("   ")
        with pytest.raises(PdfError):
            parse_page_range(",")

    def test_zero_and_negative_rejected(self):
        with pytest.raises(PdfError):
            parse_page_range("0")
        with pytest.raises(PdfError):
            parse_page_range("0-3")

    def test_bad_format_rejected(self):
        for bad in ("abc", "1-2-3", "1.5", "第3页", "-3"):
            with pytest.raises(PdfError):
                parse_page_range(bad)

    def test_huge_range_rejected(self):
        with pytest.raises(PdfError):
            parse_page_range("1-999999")


class TestSplitDropData:
    """tkinterdnd2 数据串解析：★ 反斜杠绝不能被当转义符吃掉（历史 bug）。"""

    def test_braces_with_spaces(self):
        assert split_drop_data("{D:/a b/1.pdf} D:/c.pdf") == ["D:/a b/1.pdf", "D:/c.pdf"]

    def test_backslash_not_escaped(self):
        # \t \n 在 Tcl splitlist 下会变成制表符/换行，这里必须原样保留
        assert split_drop_data(r"C:\temp\note.pdf") == [r"C:\temp\note.pdf"]

    def test_mixed_multiple_files(self):
        data = r"{C:\我的 文档\a.pdf} C:\plain\b.pdf {D:\x y.pdf}"
        assert split_drop_data(data) == [r"C:\我的 文档\a.pdf", r"C:\plain\b.pdf", r"D:\x y.pdf"]

    def test_quoted_path(self):
        assert split_drop_data('"D:/a b/1.pdf"') == ["D:/a b/1.pdf"]

    def test_empty(self):
        assert split_drop_data("") == []
        assert split_drop_data(None) == []
        assert split_drop_data("   ") == []


class TestFormatPageList:
    def test_short(self):
        assert format_page_list([1, 2, 3]) == "1、2、3"

    def test_empty(self):
        assert format_page_list([]) == "（无）"

    def test_truncated(self):
        text = format_page_list(list(range(1, 40)), limit=5)
        assert "共 39 页" in text and text.startswith("1、2、3、4、5")


class TestFileType:
    @pytest.mark.parametrize("path,expect", [
        ("a.pdf", True), ("A.PDF", True), (" a.Pdf ", True),
        ("a.pdfx", False), ("a.doc", False),
    ])
    def test_is_pdf(self, path, expect):
        assert is_pdf_path(path) is expect

    def test_word_and_wps(self):
        assert is_word_path("x.docx") and is_word_path("x.DOC") and is_word_path("x.txt")
        assert not is_word_path("x.pdf")
        assert is_wps_path("x.WPS") and not is_wps_path("x.docx")


class TestUnitConvert:
    def test_mm_pt_roundtrip(self):
        assert abs(mm_to_pt(25.4) - 72.0) < 1e-9
        assert abs(pt_to_mm(72.0) - 25.4) < 1e-9
        assert abs(pt_to_mm(mm_to_pt(210.0)) - 210.0) < 1e-9
