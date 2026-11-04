# -*- coding: utf-8 -*-
"""pdf_layout 核心业务规则回归：双面页序、补页、N-up/小册子、范围解析、排版 PDF。"""

import pytest

fitz = pytest.importorskip("fitz", reason="需要 PyMuPDF")

from core_utils import PdfError
from pdf_layout import (
    LayoutOptions, append_blank_page_like, build_layout_plan, build_range_pdf,
    compute_duplex_orders, compute_merged_duplex_orders, compute_merged_per_copy_orders,
    compute_print_orders, compute_single_sided_orders, merged_side_ranges,
    merged_sheet_count, nup_grid, paper_size_pt, plan_batch_summary,
    plan_booklet_sides, plan_nup_sides, resolve_range_pages,
)


# ------------------------------------------------------------------ 双面页序
class TestDuplexOrders:
    """核心规则：正面奇数正序、反面偶数逆序，且正反在同一张纸上（p 与 p+1）。"""

    def test_single_copy(self):
        fronts, backs = compute_print_orders(6, 1)
        assert fronts == [1, 3, 5]
        assert backs == [6, 4, 2]

    def test_collated_two_copies(self):
        fronts, backs = compute_duplex_orders(6, 2, collate=True)
        assert fronts == [1, 3, 5, 1, 3, 5]
        assert backs == [6, 4, 2, 6, 4, 2]

    def test_uncollated_two_copies(self):
        fronts, backs = compute_duplex_orders(6, 2, collate=False)
        assert fronts == [1, 1, 3, 3, 5, 5]
        assert backs == [6, 6, 4, 4, 2, 2]

    def test_pairing_invariant(self):
        """任意份数/分页方式下：反面第 i 张必须等于正面第 i 张 +1（翻面后整体倒序）。"""
        for copies in (1, 2, 3):
            for collate in (True, False):
                fronts, backs = compute_duplex_orders(10, copies, collate=collate)
                assert len(fronts) == len(backs)
                for f, b in zip(fronts, reversed(backs)):
                    assert b == f + 1

    def test_reverse(self):
        fronts, backs = compute_duplex_orders(6, 1, reverse=True)
        assert fronts == [5, 3, 1]
        for f, b in zip(fronts, reversed(backs)):
            assert b == f + 1


# ------------------------------------------------------------------ 整批一次翻面
class TestMergedOrders:
    def test_global_reversal_not_per_file(self):
        """两文件各 4 面、2 面：反面必须按全局逆序，不能每个文件各自逆序。"""
        fronts, backs = compute_merged_duplex_orders([4, 2])
        assert fronts == [1, 3, 5]
        assert backs == [6, 4, 2]

    def test_odd_total_rejected(self):
        with pytest.raises(PdfError):
            compute_merged_duplex_orders([4, 3])

    def test_side_ranges_start_odd(self):
        """整批翻面的前提：每个文件的全局起始页必须是奇数（各自已补成偶数）。"""
        ranges = merged_side_ranges([4, 2, 6])
        assert ranges == [(1, 4), (5, 6), (7, 12)]
        assert all(start % 2 == 1 for start, _ in ranges)

    def test_sheet_count(self):
        assert merged_sheet_count([4, 2]) == 3

    def test_per_copy_orders(self):
        items = compute_merged_per_copy_orders([4, 2], 2)
        assert len(items) == 2
        assert all(f == [1, 3, 5] and b == [6, 4, 2] for _, f, b in items)

    def test_batch_summary_flip_counts(self):
        s1 = plan_batch_summary([4, 2], copies=3, exec_mode="batch", copy_exec="all")
        assert s1["flip_count"] == 1
        s2 = plan_batch_summary([4, 2], copies=3, exec_mode="batch", copy_exec="per_copy")
        assert s2["flip_count"] == 3
        s3 = plan_batch_summary([4, 2], exec_mode="by_file")
        assert s3["flip_count"] == 2


class TestSingleSided:
    def test_orders(self):
        assert compute_single_sided_orders(3, 2, collate=True) == [1, 2, 3, 1, 2, 3]
        assert compute_single_sided_orders(3, 2, collate=False) == [1, 1, 2, 2, 3, 3]
        assert compute_single_sided_orders(3, 1, reverse=True) == [3, 2, 1]


# ------------------------------------------------------------------ 补页规则
class TestLayoutPlanPadding:
    def test_duplex_odd_gets_blank(self):
        """核心规则：选定 5 页 → 补 1 张 → 6 个输出面。"""
        plan = build_layout_plan(5, LayoutOptions(manual_duplex=True))
        assert plan.padded_count == 6 and plan.duplex_blank_added
        assert plan.output_count == 6

    def test_even_no_blank(self):
        plan = build_layout_plan(6, LayoutOptions(manual_duplex=True))
        assert plan.padded_count == 6 and not plan.duplex_blank_added

    def test_single_sided_no_duplex_blank(self):
        plan = build_layout_plan(5, LayoutOptions(manual_duplex=False))
        assert plan.padded_count == 5

    def test_nup_output_side_padded(self):
        """4 页按 4-up 只有 1 个输出面 → 双面必须再补 1 个空白输出面。"""
        plan = build_layout_plan(4, LayoutOptions(manual_duplex=True, nup=4))
        assert plan.output_count == 2 and plan.output_blank_added == 1
        assert plan.sides[1] == [None, None, None, None]

    def test_booklet_pads_to_multiple_of_4(self):
        """小册子：6 页 → 先双面补到 6 → 再补到 8（4 的倍数）。"""
        plan = build_layout_plan(6, LayoutOptions(manual_duplex=True, booklet=True))
        assert plan.padded_count == 8
        assert plan.output_count == 4

    def test_booklet_page_positions(self):
        """骑马钉 8 页：正面 [8,1] 背面 [2,7] / 正面 [6,3] 背面 [4,5]。"""
        assert plan_booklet_sides(8) == [[8, 1], [2, 7], [6, 3], [4, 5]]

    def test_booklet_non_multiple_rejected(self):
        with pytest.raises(PdfError):
            plan_booklet_sides(6)

    def test_plan_requires_page(self):
        with pytest.raises(PdfError):
            build_layout_plan(0, LayoutOptions())


# ------------------------------------------------------------------ 范围解析
class TestResolveRange:
    def test_all(self):
        valid, dropped = resolve_range_pages("all", "", 5)
        assert valid == [1, 2, 3, 4, 5] and dropped == []

    def test_custom_clips_beyond_page_count(self):
        valid, dropped = resolve_range_pages("custom", "2-9", 4)
        assert valid == [2, 3, 4] and dropped == [5, 6, 7, 8, 9]

    def test_odd_even(self):
        assert resolve_range_pages("odd", "", 5)[0] == [1, 3, 5]
        assert resolve_range_pages("even", "", 5)[0] == [2, 4]

    def test_current_clamped(self):
        assert resolve_range_pages("current", "", 3, current_page=99)[0] == [3]
        assert resolve_range_pages("current", "", 3, current_page=-2)[0] == [1]

    def test_empty_pdf_rejected(self):
        with pytest.raises(PdfError):
            resolve_range_pages("all", "", 0)

    def test_unknown_mode_rejected(self):
        with pytest.raises(PdfError):
            resolve_range_pages("bogus", "", 5)


# ------------------------------------------------------------------ 版式几何
class TestGeometry:
    def test_nup_grid(self):
        assert nup_grid(1) == (1, 1)
        assert nup_grid(6) == (3, 2)
        assert nup_grid(99) == (1, 1)

    def test_paper_size_orientation(self):
        opts = LayoutOptions(paper_name="A4", orientation="landscape")
        w, h = paper_size_pt(opts)
        assert w > h
        opts2 = LayoutOptions(paper_name="A4", orientation="auto")
        w2, h2 = paper_size_pt(opts2, first_page_landscape=True)
        assert w2 > h2

    def test_nup_sides_chunks(self):
        assert plan_nup_sides(5, 2) == [[1, 2], [3, 4], [5, None]]


# ------------------------------------------------------------------ 真实 PDF（补页与 build_range_pdf）
def _make_pdf(path, count, size=(595, 842)):
    doc = fitz.open()
    for _ in range(count):
        doc.new_page(width=size[0], height=size[1])
    doc.save(str(path))
    doc.close()


class TestRealPdf:
    def test_build_range_pdf_adds_blank(self, tmp_path):
        src = tmp_path / "src.pdf"
        out = tmp_path / "out.pdf"
        _make_pdf(src, 3)
        total, added = build_range_pdf(str(src), [1, 2, 3], str(out))
        assert (total, added) == (4, True)
        doc = fitz.open(str(out))
        try:
            assert doc.page_count == 4
            # 补页尺寸必须与模板页一致
            assert abs(doc[3].rect.width - doc[2].rect.width) < 0.01
            assert abs(doc[3].rect.height - doc[2].rect.height) < 0.01
        finally:
            doc.close()

    def test_build_range_pdf_even_no_blank(self, tmp_path):
        src = tmp_path / "src.pdf"
        out = tmp_path / "out.pdf"
        _make_pdf(src, 4)
        total, added = build_range_pdf(str(src), [1, 2, 3, 4], str(out))
        assert (total, added) == (4, False)

    def test_blank_page_keeps_rotation(self, tmp_path):
        src = tmp_path / "rot.pdf"
        _make_pdf(src, 1)
        doc = fitz.open(str(src))
        try:
            doc[0].set_rotation(90)
            doc.saveIncr()
        finally:
            doc.close()
        doc = fitz.open(str(src))
        try:
            append_blank_page_like(doc, doc[0])
            blank = doc[1]
            assert blank.rotation == 90
            # 显示尺寸（计入 /Rotate）必须与模板页一致
            assert abs(blank.rect.width - doc[0].rect.width) < 0.01
            assert abs(blank.rect.height - doc[0].rect.height) < 0.01
        finally:
            doc.close()

    def test_encrypted_pdf_rejected(self, tmp_path):
        src = tmp_path / "enc.pdf"
        _make_pdf(src, 1)
        doc = fitz.open(str(src))
        try:
            doc.save(str(tmp_path / "enc2.pdf"), encryption=fitz.PDF_ENCRYPT_AES_256,
                     user_pw="pw")
        finally:
            doc.close()
        with pytest.raises(PdfError):
            build_range_pdf(str(tmp_path / "enc2.pdf"), [1], str(tmp_path / "o.pdf"))
