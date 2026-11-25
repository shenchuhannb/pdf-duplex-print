# -*- coding: utf-8 -*-
"""pdf_layout 核心业务规则回归：双面页序、补页、N-up/小册子、范围解析、排版 PDF。"""

import pytest

fitz = pytest.importorskip("fitz", reason="需要 PyMuPDF")

from core_utils import PdfError
from pdf_layout import (
    NUP_CHOICES, LayoutOptions, append_blank_page_like, auto_rotate_allowed,
    build_layout_plan, build_range_pdf, compose_layout_pdf, compute_duplex_orders,
    compute_merged_duplex_orders, compute_merged_per_copy_orders, compute_print_orders,
    compute_single_sided_orders, content_side_numbers, merged_side_ranges,
    merged_sheet_count, nup_grid, paper_size_pt, plan_batch_summary, plan_booklet_sides,
    plan_nup_sides, preview_metrics, resolve_range_pages, side_scale_percents, slot_rects,
)


class TestSideScalePercents:
    """缩放必须逐页按源页尺寸计算（对齐 Acrobat“适合”），不是整批固定值。"""

    OPTIONS = LayoutOptions(printable_insets_pt=(11.28, 11.28, 15.72, 11.76))

    def test_percent_follows_each_source_page(self):
        plan = build_layout_plan(2, self.OPTIONS)
        sizes = {1: (595.28, 841.89), 2: (300.0, 400.0)}   # A4 + 小页面
        scales = side_scale_percents(plan, self.OPTIONS, lambda rel: sizes[rel])
        assert len(scales) == 2
        big, small = scales[0][0], scales[1][0]
        assert abs(big - 95.47) < 0.2        # A4 -> P1007 可打印区约 95.5%
        assert small > big + 30              # 小页面放大到槽位，比例明显不同

    def test_blank_side_has_no_percent(self):
        plan = build_layout_plan(1, self.OPTIONS)          # 双面补 1 空白面
        scales = side_scale_percents(
            plan, self.OPTIONS, lambda rel: (595.28, 841.89))
        assert scales[0][0] > 0
        assert scales[1] == (0.0, False)

    def test_actual_size_is_hundred_percent(self):
        options = LayoutOptions(scale_mode="actual",
                                printable_insets_pt=(11.28, 11.28, 15.72, 11.76))
        plan = build_layout_plan(1, options)
        scales = side_scale_percents(plan, options, lambda rel: (595.28, 841.89))
        assert abs(scales[0][0] - 100.0) < 0.01
        assert scales[0][1] is True            # 超出可打印区，标记溢出裁切

    def test_eight_up_percent_uses_auto_rotated_source(self):
        options = LayoutOptions(nup=8, orientation="portrait", manual_duplex=False,
                                scale_mode="fit", auto_rotate=True)
        plan = build_layout_plan(8, options)
        scales = side_scale_percents(plan, options, lambda rel: (595.28, 841.89))
        slot = slot_rects(plan.paper_pt[0], plan.paper_pt[1], 8, options, gap_mm=2.0)[0]
        expected = 100.0 * min(slot.width / 841.89, slot.height / 595.28)
        assert abs(scales[0][0] - expected) < 0.01
        assert 30 < scales[0][0] < 40


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
        assert nup_grid(8, landscape=False) == (2, 4)
        assert nup_grid(8, landscape=True) == (4, 2)
        assert nup_grid(99) == (1, 1)

    def test_eight_up_slots_follow_paper_orientation(self):
        portrait = slot_rects(595, 842, 8, LayoutOptions())
        landscape = slot_rects(842, 595, 8, LayoutOptions())
        assert len(portrait) == len(landscape) == 8
        assert len({round(slot.x0, 2) for slot in portrait}) == 2
        assert len({round(slot.y0, 2) for slot in portrait}) == 4
        assert len({round(slot.x0, 2) for slot in landscape}) == 4
        assert len({round(slot.y0, 2) for slot in landscape}) == 2

    def test_paper_size_orientation(self):
        opts = LayoutOptions(paper_name="A4", orientation="landscape")
        w, h = paper_size_pt(opts)
        assert w > h
        opts2 = LayoutOptions(paper_name="A4", orientation="auto")
        w2, h2 = paper_size_pt(opts2, first_page_landscape=True)
        assert w2 > h2

    def test_nup_sides_chunks(self):
        assert plan_nup_sides(5, 2) == [[1, 2], [3, 4], [5, None]]
        assert 8 in NUP_CHOICES
        assert plan_nup_sides(9, 8) == [
            [1, 2, 3, 4, 5, 6, 7, 8],
            [9, None, None, None, None, None, None, None],
        ]


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

    @pytest.mark.parametrize("orientation", ["portrait", "landscape"])
    def test_eight_up_compose_orientation(self, tmp_path, orientation):
        src = tmp_path / "eight.pdf"
        out = tmp_path / f"eight-{orientation}.pdf"
        _make_pdf(src, 8)
        result = compose_layout_pdf(
            str(src), list(range(1, 9)), str(out),
            LayoutOptions(nup=8, orientation=orientation, manual_duplex=False))
        assert result.plan.sides == [[1, 2, 3, 4, 5, 6, 7, 8]]
        doc = fitz.open(str(out))
        try:
            assert doc.page_count == 1
            page = doc[0]
            assert (page.rect.width > page.rect.height) == (orientation == "landscape")
        finally:
            doc.close()

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


# ------------------------------------------------------------- 内容面识别（单页免双面）
class TestContentSideNumbers:
    """只有 1 面内容时不应再走翻面流程；补页/空槽位不算内容。"""

    def test_single_page_duplex_has_one_content_side(self):
        plan = build_layout_plan(1, LayoutOptions(manual_duplex=True))
        assert plan.output_count == 2          # 仍补配对空白页，保持整批偶数不变量
        assert content_side_numbers(plan) == [1]

    def test_three_pages_all_content_sides(self):
        plan = build_layout_plan(3, LayoutOptions(manual_duplex=True))
        assert content_side_numbers(plan) == [1, 2, 3]

    def test_nup_single_output_side(self):
        plan = build_layout_plan(2, LayoutOptions(manual_duplex=True, nup=2))
        assert content_side_numbers(plan) == [1]

    def test_booklet_single_page(self):
        plan = build_layout_plan(1, LayoutOptions(manual_duplex=True, booklet=True))
        assert content_side_numbers(plan) == [1]

    def test_single_sided_plan_unaffected(self):
        plan = build_layout_plan(2, LayoutOptions(manual_duplex=False))
        assert content_side_numbers(plan) == [1, 2]


# ------------------------------------------------------------- 横版纸张 + N-up
class TestLandscapeNup:
    """orientation=landscape 时纸张必须宽>高，槽位按横向纸张划分，
    且预览上报的纸张/缩放与真实合成页面完全一致（不是只看预览字符串）。"""

    PAGE = (400.0, 600.0)   # 源页：竖向 400×600

    def _plan_and_slots(self, nup, manual_duplex):
        options = LayoutOptions(orientation="landscape", nup=nup,
                                manual_duplex=manual_duplex, scale_mode="fit")
        plan = build_layout_plan(4, options)
        slots = slot_rects(plan.paper_pt[0], plan.paper_pt[1], plan.nup,
                           options, gap_mm=2.0)
        return options, plan, slots

    @pytest.mark.parametrize("nup", (2, 4))
    @pytest.mark.parametrize("manual_duplex", (True, False))
    def test_landscape_paper_and_slots(self, nup, manual_duplex):
        options, plan, slots = self._plan_and_slots(nup, manual_duplex)
        paper_w, paper_h = plan.paper_pt
        assert paper_w > paper_h                       # A4 横版：297×210mm
        assert paper_w == pytest.approx(841.89, abs=0.05)
        assert paper_h == pytest.approx(595.28, abs=0.05)
        assert len(slots) == nup
        # 槽位必须落在纸张内，且按横向纸张排布：nup=2 → 2 列 1 行；nup=4 → 2×2。
        for slot in slots:
            assert slot.x0 >= -0.01 and slot.y0 >= -0.01
            assert slot.x1 <= paper_w + 0.01 and slot.y1 <= paper_h + 0.01
        assert len({round(s.y0, 2) for s in slots}) == (1 if nup == 2 else 2)
        assert len({round(s.x0, 2) for s in slots}) == 2
        assert slots[0].width < paper_w / 1.5          # 槽位宽度按横向纸张均分

    @pytest.mark.parametrize("nup", (2, 4))
    @pytest.mark.parametrize("manual_duplex", (True, False))
    def test_compose_page_matches_plan_paper(self, tmp_path, nup, manual_duplex):
        src = tmp_path / "src.pdf"
        make = tmp_path / "out.pdf"
        _make_pdf(src, 8, size=self.PAGE)
        options = LayoutOptions(orientation="landscape", nup=nup,
                                manual_duplex=manual_duplex, scale_mode="fit")
        result = compose_layout_pdf(str(src), list(range(1, 9)), str(make), options)
        doc = fitz.open(str(make))
        try:
            assert doc.page_count == result.plan.output_count
            for page in doc:
                assert page.rect.width > page.rect.height
                assert page.rect.width == pytest.approx(result.plan.paper_pt[0], abs=0.05)
                assert page.rect.height == pytest.approx(result.plan.paper_pt[1], abs=0.05)
            assert len(result.side_scales) == doc.page_count
        finally:
            doc.close()

    def test_preview_metrics_paper_and_scale_match_compose(self, tmp_path):
        """预览 metrics 的纸张尺寸 = 合成页尺寸；缩放百分比 = 真实放置比例。"""
        src = tmp_path / "src.pdf"
        make = tmp_path / "out.pdf"
        _make_pdf(src, 8, size=self.PAGE)
        options = LayoutOptions(orientation="landscape", nup=2,
                                manual_duplex=True, scale_mode="fit")
        result = compose_layout_pdf(str(src), list(range(1, 9)), str(make), options)
        metrics = preview_metrics(result.plan, options, *self.PAGE)
        doc = fitz.open(str(make))
        try:
            for index, page in enumerate(doc):
                assert metrics["paper"] == pytest.approx(
                    (page.rect.width, page.rect.height), abs=0.05)
                scale, overflow = result.side_scales[index]
                if scale:
                    assert scale > 0
            assert result.side_scales[0][0] == pytest.approx(metrics["scale_percent"], abs=0.01)
            assert metrics["scale_percent"] < 100       # 400×600 竖页放进横向槽位会缩小
        finally:
            doc.close()

    def test_preview_metrics_rejects_paper_sized_source(self, tmp_path):
        """用纸张尺寸冒充源页尺寸会算出错误比例（回归防护：两者必须不同）。"""
        options = LayoutOptions(orientation="landscape", nup=2,
                                manual_duplex=True, scale_mode="fit")
        plan = build_layout_plan(4, options)
        wrong = preview_metrics(plan, options, *plan.paper_pt)["scale_percent"]
        right = preview_metrics(plan, options, *self.PAGE)["scale_percent"]
        assert abs(wrong - right) > 5

    def test_landscape_swaps_paper_dimensions(self):
        portrait = paper_size_pt(LayoutOptions(orientation="portrait"))
        landscape = paper_size_pt(LayoutOptions(orientation="landscape"))
        assert landscape == pytest.approx((portrait[1], portrait[0]), abs=0.01)
        assert landscape[0] > landscape[1]

    def test_booklet_auto_uses_landscape_sheet_with_portrait_slots(self):
        options = LayoutOptions(booklet=True, manual_duplex=True)
        plan = build_layout_plan(8, options)
        assert plan.paper_pt[0] > plan.paper_pt[1]
        slots = slot_rects(plan.paper_pt[0], plan.paper_pt[1], 2, options, gap_mm=0.0)
        assert len(slots) == 2
        assert all(slot.height > slot.width for slot in slots)

    def test_booklet_landscape_keeps_landscape_sheet(self):
        options = LayoutOptions(booklet=True, orientation="landscape", manual_duplex=True)
        plan = build_layout_plan(8, options)
        assert plan.paper_pt[0] > plan.paper_pt[1]
        assert "小册子建议" not in __import__("pdf_layout").describe_plan(plan)


# ------------------------------------------------------------- 方形 N-up 保持阅读方向
class TestSquareNupContentDirection:
    """竖版 A4 放到横版纸的方形网格：默认不旋转，显式开启后才旋转。"""

    A4 = (595.28, 841.89)
    GRIDS = {4: (2, 2), 9: (3, 3), 16: (4, 4)}

    @classmethod
    def _make_marked(cls, path, count):
        doc = fitz.open()
        for index in range(count):
            page = doc.new_page(width=cls.A4[0], height=cls.A4[1])
            label = f"P{index + 1:02d}"
            page.insert_text((36, 42), f"TOP-{label}", fontsize=14)
            page.insert_text((72, cls.A4[1] / 2), f"BODY-{label}", fontsize=18)
            page.insert_text((36, cls.A4[1] - 30), f"BOTTOM-{label}", fontsize=14)
        doc.save(str(path))
        doc.close()

    @staticmethod
    def _text_lines(page):
        result = []
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                text = "".join(span.get("text", "") for span in line.get("spans", [])).strip()
                result.append((text, fitz.Rect(line["bbox"]), tuple(line["dir"])))
        return result

    @classmethod
    def _assert_composition(cls, result, out_path, options, nup, expect_vertical):
        assert nup_grid(nup, landscape=True) == cls.GRIDS[nup]
        slots = slot_rects(result.plan.paper_pt[0], result.plan.paper_pt[1],
                           nup, options, gap_mm=2.0)
        assert len(slots) == nup
        # 横版 A4 的 4/9/16 方形网格产生横版槽位，正是本回归的触发条件。
        assert result.plan.paper_pt[0] > result.plan.paper_pt[1]
        assert all(slot.width > slot.height for slot in slots)
        assert len(result.side_scales) == 1
        assert result.side_scales[0][1] is False

        doc = fitz.open(str(out_path))
        try:
            assert doc.page_count == 1
            lines = cls._text_lines(doc[0])
            expected_all = {
                f"{edge}-P{index + 1:02d}"
                for index in range(nup) for edge in ("TOP", "BODY", "BOTTOM")
            }
            assert {text for text, _bbox, _direction in lines} == expected_all

            for index, slot in enumerate(slots):
                expected = {
                    f"{edge}-P{index + 1:02d}" for edge in ("TOP", "BODY", "BOTTOM")
                }
                assigned = []
                for text, bbox, direction in lines:
                    center_x = (bbox.x0 + bbox.x1) / 2
                    center_y = (bbox.y0 + bbox.y1) / 2
                    if slot.x0 <= center_x <= slot.x1 and slot.y0 <= center_y <= slot.y1:
                        assigned.append((text, bbox, direction))
                # 逐槽归属防止只检查全页文字集合而漏掉跨槽/错槽内容。
                assert {text for text, _bbox, _direction in assigned} == expected
                positions = {}
                for text, bbox, direction in assigned:
                    if expect_vertical:
                        assert abs(direction[0]) < 0.01 and abs(direction[1]) > 0.99, (
                            nup, text, direction)
                    else:
                        assert direction[0] > 0.99 and abs(direction[1]) < 0.01, (
                            nup, text, direction)
                    # 所有标记完整落在自己的槽位内，不能用页面级“不越界”代替。
                    assert bbox.x0 >= slot.x0 - 0.75 and bbox.y0 >= slot.y0 - 0.75
                    assert bbox.x1 <= slot.x1 + 0.75 and bbox.y1 <= slot.y1 + 0.75
                    positions[text] = bbox
                if not expect_vertical:
                    label = f"P{index + 1:02d}"
                    top_y = (positions[f"TOP-{label}"].y0 + positions[f"TOP-{label}"].y1) / 2
                    body_y = (positions[f"BODY-{label}"].y0 + positions[f"BODY-{label}"].y1) / 2
                    bottom_y = (positions[f"BOTTOM-{label}"].y0
                                + positions[f"BOTTOM-{label}"].y1) / 2
                    assert top_y < body_y < bottom_y
        finally:
            doc.close()

        metrics = preview_metrics(result.plan, options, *cls.A4)
        assert metrics["overflow"] is False
        assert result.side_scales[0][0] == pytest.approx(metrics["scale_percent"], abs=0.01)

    @pytest.mark.parametrize("nup", (4, 9, 16))
    def test_landscape_square_nup_default_keeps_text_horizontal(self, tmp_path, nup):
        src = tmp_path / f"portrait-a4-{nup}.pdf"
        out = tmp_path / f"landscape-{nup}-default.pdf"
        self._make_marked(src, nup)
        # 不显式传 auto_rotate：必须走 LayoutOptions 的新默认值 False。
        options = LayoutOptions(orientation="landscape", nup=nup,
                                manual_duplex=False, scale_mode="fit")
        assert options.auto_rotate is False
        result = compose_layout_pdf(
            str(src), list(range(1, nup + 1)), str(out), options)
        self._assert_composition(result, out, options, nup, expect_vertical=False)

    @pytest.mark.parametrize("nup", (4, 9, 16))
    def test_landscape_square_nup_explicit_auto_rotate_still_rotates(self, tmp_path, nup):
        src = tmp_path / f"portrait-a4-{nup}.pdf"
        out = tmp_path / f"landscape-{nup}-rotated.pdf"
        self._make_marked(src, nup)
        options = LayoutOptions(orientation="landscape", nup=nup,
                                manual_duplex=False, scale_mode="fit", auto_rotate=True)
        result = compose_layout_pdf(
            str(src), list(range(1, nup + 1)), str(out), options)
        self._assert_composition(result, out, options, nup, expect_vertical=True)


# ------------------------------------------------------------- auto 首页 + /Rotate
class TestAutoOrientationAndRotateBake:
    """auto 必须与合成页一致；`/Rotate` 必须先烘培再裁切，不能丢内容。

    这两类问题在诊断报告中是独立缺陷：
      * auto + 横版首页曾让 plan/preview 用竖版纸，而合成页是横版；
      * `/Rotate != 0` 的源页直接 show_pdf_page 会按未旋转坐标绘制并被裁切。
    """

    MEDIA = (595.0, 842.0)   # A4 竖版媒体框

    def _make_marked(self, path, rotation: int = 0):
        doc = fitz.open()
        page = doc.new_page(width=self.MEDIA[0], height=self.MEDIA[1])
        page.insert_text((40, 40), "TOP-EDGE", fontsize=14)
        page.insert_text((40, self.MEDIA[1] - 30), "BOTTOM-EDGE", fontsize=14)
        page.insert_text((200, 420), "CENTER-MARK", fontsize=20)
        if rotation:
            page.set_rotation(rotation)
        doc.save(str(path))
        doc.close()

    def test_auto_landscape_first_page_matches_composition(self, tmp_path):
        """auto + 横版首页：plan / 合成页 / preview_metrics 的纸面必须一致。"""
        src = tmp_path / "landscape-first.pdf"
        out = tmp_path / "landscape-first-out.pdf"
        doc = fitz.open()
        for index in range(2):
            page = doc.new_page(width=842, height=595)
            page.insert_text((40, 70), f"L{index + 1}", fontsize=20)
        doc.save(str(src))
        doc.close()

        options = LayoutOptions(orientation="auto", manual_duplex=False,
                                scale_mode="shrink")
        result = compose_layout_pdf(str(src), [1, 2], str(out), options)
        assert result.plan.paper_pt[0] > result.plan.paper_pt[1]

        source = fitz.open(str(src))
        try:
            first = source[0].rect
            metrics = preview_metrics(result.plan, options, first.width, first.height)
        finally:
            source.close()
        composed = fitz.open(str(out))
        try:
            for page in composed:
                assert page.rect.width > page.rect.height
                assert metrics["paper"] == pytest.approx(
                    (page.rect.width, page.rect.height), abs=0.05)
                assert metrics["paper"] == pytest.approx(result.plan.paper_pt, abs=0.05)
        finally:
            composed.close()

    @pytest.mark.parametrize("rotation", (90, 180, 270))
    def test_rotated_source_is_not_clipped(self, tmp_path, rotation):
        """`/Rotate=90/180/270`：烘培后合成，TOP/BOTTOM 标记都在页面内。"""
        src = tmp_path / f"rotate-{rotation}.pdf"
        out = tmp_path / f"rotate-{rotation}-out.pdf"
        self._make_marked(src, rotation)

        # auto 跟随首页显示方向：90/270 是横版显示，180 是竖版显示。
        options = LayoutOptions(orientation="auto", manual_duplex=False,
                                auto_rotate=True)
        result = compose_layout_pdf(str(src), [1], str(out), options)
        composed = fitz.open(str(out))
        try:
            assert composed.page_count == 1
            page = composed[0]
            assert page.rect.width == pytest.approx(result.plan.paper_pt[0], abs=0.05)
            assert page.rect.height == pytest.approx(result.plan.paper_pt[1], abs=0.05)
            words = page.get_text("words")
            texts = {word[4] for word in words}
            assert "TOP-EDGE" in texts, texts
            assert "BOTTOM-EDGE" in texts, texts
            assert "CENTER-MARK" in texts, texts
            # 未裁剪内容必须完整落在页面内（0.6pt 容差覆盖字形 bbox 取整）。
            for word in words:
                assert word[0] >= -0.6 and word[1] >= -0.6
                assert word[2] <= page.rect.width + 0.6
                assert word[3] <= page.rect.height + 0.6
        finally:
            composed.close()

    @pytest.mark.parametrize("rotation", (90, 180, 270))
    def test_rotated_source_with_crop_keeps_markers(self, tmp_path, rotation):
        """先烘培 /Rotate 再按显示坐标裁切：裁切不会吃掉整个页面或标记。"""
        src = tmp_path / f"rotate-crop-{rotation}.pdf"
        out = tmp_path / f"rotate-crop-{rotation}-out.pdf"
        doc = fitz.open()
        # 媒体框 595×842；90/270 显示为 842×595，180 仍为 595×842。
        page = doc.new_page(width=self.MEDIA[0], height=self.MEDIA[1])
        page.insert_text((40, 40), "TOP-EDGE", fontsize=14)
        page.insert_text((40, self.MEDIA[1] - 30), "BOTTOM-EDGE", fontsize=14)
        page.set_rotation(rotation)
        doc.save(str(src))
        doc.close()

        options = LayoutOptions(orientation="auto", manual_duplex=False,
                                auto_rotate=False,
                                crop_left_mm=2, crop_top_mm=2,
                                crop_right_mm=2, crop_bottom_mm=2)
        result = compose_layout_pdf(str(src), [1], str(out), options)
        composed = fitz.open(str(out))
        try:
            page = composed[0]
            assert page.rect.width == pytest.approx(result.plan.paper_pt[0], abs=0.05)
            texts = {word[4] for word in page.get_text("words")}
            assert "TOP-EDGE" in texts, texts
            assert "BOTTOM-EDGE" in texts, texts
        finally:
            composed.close()


# --------------------------------------- 单页保持阅读方向（nup=1 不旋转）
class TestSinglePageKeepsReadingDirection:
    """单页（nup=1）不允许自动旋转：横版内容+竖版纸只缩放居中并留白。

    只有拼接（N-up>1）或小册子才需要把内容转 90° 适配槽位；单页旋转会让
    用户看到的内容侧躺，与源 PDF 的阅读方向相反。
    """

    WIDE = (842.0, 595.0)   # 横版源页（A4 横向）

    def _make_wide_marked(self, path):
        doc = fitz.open()
        page = doc.new_page(width=self.WIDE[0], height=self.WIDE[1])
        # 左右两条文字尽量横跨整幅，便于断言缩放后仍横向铺开。
        page.insert_text((20, 100), "LEFT-EDGE", fontsize=20)
        page.insert_text((700, 100), "RIGHT-EDGE", fontsize=20)
        page.insert_text((20, 565), "WIDE-FOOTER", fontsize=14)
        doc.save(str(path))
        doc.close()

    @staticmethod
    def _line_dirs(page):
        return [tuple(round(v, 2) for v in line["dir"])
                for block in page.get_text("dict")["blocks"]
                for line in block.get("lines", [])]

    def test_auto_rotate_allowed_only_for_nup_or_booklet(self):
        assert auto_rotate_allowed(
            LayoutOptions(nup=1, auto_rotate=True), 1, False) is False
        assert auto_rotate_allowed(
            LayoutOptions(nup=2, auto_rotate=True), 2, False) is True
        assert auto_rotate_allowed(
            LayoutOptions(booklet=True, auto_rotate=True), 1, True) is True
        assert auto_rotate_allowed(
            LayoutOptions(nup=4, auto_rotate=False), 4, False) is False

    def test_single_page_landscape_on_portrait_paper_is_not_rotated(self, tmp_path):
        src = tmp_path / "wide.pdf"
        out = tmp_path / "wide-single.pdf"
        self._make_wide_marked(src)
        options = LayoutOptions(orientation="portrait", nup=1,
                                manual_duplex=False, auto_rotate=True,
                                scale_mode="fit")
        result = compose_layout_pdf(str(src), [1], str(out), options)
        assert result.plan.paper_pt[1] > result.plan.paper_pt[0]   # 竖版纸
        doc = fitz.open(str(out))
        try:
            page = doc[0]
            # 文字方向必须保持水平（旋转 90° 后 line["dir"] 会变成 (0, ±1)）。
            dirs = self._line_dirs(page)
            assert dirs and all(abs(d[1]) < 0.01 and d[0] > 0.99 for d in dirs), dirs
            words = page.get_text("words")
            assert {"LEFT-EDGE", "RIGHT-EDGE", "WIDE-FOOTER"} <= {w[4] for w in words}
            x0 = min(w[0] for w in words)
            x1 = max(w[2] for w in words)
            y0 = min(w[1] for w in words)
            y1 = max(w[3] for w in words)
            # 内容按纸宽铺开、水平居中，上下留出明显白边（不旋转、仅缩放）。
            assert x0 < page.rect.width * 0.12
            assert x1 > page.rect.width * 0.85
            assert abs((x0 + x1) / 2 - page.rect.width / 2) < 6
            assert y0 - page.rect.y0 > 36
            assert page.rect.y1 - y1 > 36
        finally:
            doc.close()
        # 预览 metrics 与实际放置比例一致（单页两处都不能旋转）。
        metrics = preview_metrics(result.plan, options, *self.WIDE)
        assert result.side_scales[0][0] == pytest.approx(
            metrics["scale_percent"], abs=0.01)

    def test_nup_two_still_rotates_landscape_to_portrait_slot(self, tmp_path):
        src = tmp_path / "wide-nup.pdf"
        out = tmp_path / "wide-nup2.pdf"
        self._make_wide_marked(src)
        options = LayoutOptions(orientation="portrait", nup=2,
                                manual_duplex=False, auto_rotate=True,
                                scale_mode="fit")
        result = compose_layout_pdf(str(src), [1], str(out), options)
        doc = fitz.open(str(out))
        try:
            dirs = {d for page in doc for d in self._line_dirs(page)}
            assert dirs
            # N-up 槽位是纵向的，横版内容必须旋转 90° 才能适配。
            assert all(abs(d[0]) < 0.01 and abs(d[1]) > 0.99 for d in dirs), dirs
        finally:
            doc.close()
        metrics = preview_metrics(result.plan, options, *self.WIDE)
        assert result.side_scales[0][0] == pytest.approx(
            metrics["scale_percent"], abs=0.01)
