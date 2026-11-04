# -*- coding: utf-8 -*-
"""PDF 排版与页面逻辑（核心业务规则都在这里，全部可脱离界面单测）。

核心规则（保持不变）：
  * 相对页码：以用户选定打印范围为基准，范围第 1 页 = 相对第 1 页；
    奇偶判断与打印顺序全部基于相对页码，而不是原 PDF 绝对页码。
  * 手动双面：先正序打印相对奇数页，用户整体翻面后逆序打印相对偶数页。
  * 范围总页数为奇数时，在范围末尾追加一张与原页面同尺寸同方向的空白页。
  * N-up / 小册子排版后，若输出面仍为奇数，再补一张空白输出面，保证双面配对。
"""

import math
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    fitz = None

from core_utils import (
    MM_PER_INCH, PdfError, ToolError, mm_to_pt, parse_page_range,
)

# 常用纸张尺寸（毫米，纵向）
PAPER_SIZES_MM = {
    "A4": (210.0, 297.0),
    "A5": (148.0, 210.0),
    "B5": (176.0, 250.0),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
}
CUSTOM_PAPER_NAME = "自定义"

# Windows DMPAPER_* 纸张代码（用于 DEVMODE.dmPaperSize）
PAPER_CODES = {
    "Letter": 1,
    "Legal": 5,
    "A4": 9,
    "A5": 11,
    "B5": 13,
    CUSTOM_PAPER_NAME: 256,   # DMPAPER_USER
}

NUP_CHOICES = (1, 2, 4, 6, 9, 16)
SCALE_MODES = ("fit", "shrink", "actual", "custom", "paper")
ORIENTATIONS = ("auto", "portrait", "landscape")
RANGE_MODES = ("all", "current", "custom", "odd", "even", "recent")


# ----------------------------------------------------------------------------
# PDF 基础操作
# ----------------------------------------------------------------------------
def open_pdf(path: str) -> "fitz.Document":
    """打开 PDF，并对常见异常给出友好提示。"""
    if fitz is None:
        raise PdfError("缺少 PyMuPDF 库，请执行：pip install PyMuPDF")
    if not os.path.isfile(path):
        raise PdfError(f"文件不存在：{path}")
    try:
        doc = fitz.open(path)
    except Exception as exc:
        raise PdfError(f"无法打开 PDF（文件可能损坏或被占用）：\n{path}\n{exc}") from exc

    if not doc.is_pdf:
        doc.close()
        raise PdfError(f"该文件不是有效的 PDF：\n{path}")
    if doc.needs_pass:
        doc.close()
        raise PdfError(f"该 PDF 已加密（需要密码），暂不支持：\n{path}")
    if doc.page_count == 0:
        doc.close()
        raise PdfError(f"该 PDF 没有任何页面：\n{path}")
    return doc


def get_pdf_page_count(path: str) -> int:
    """读取 PDF 页数（同时验证文件是否可读）。"""
    doc = open_pdf(path)
    try:
        return doc.page_count
    finally:
        doc.close()


def page_is_blank(page: "fitz.Page") -> bool:
    """判断页面是否为空白（无文字、无矢量图、无图片）。

    扫描件中的空白页仍带有图片，因此不会被误判为空白页，
    这是“忽略空白页”功能的一种保守降级实现。
    """
    try:
        if page.get_text().strip():
            return False
        if page.get_drawings():
            return False
        if page.get_images(full=True):
            return False
    except Exception:
        # 读取失败时按“非空白”处理，避免误删有效内容
        return False
    return True


def append_blank_page_like(doc: "fitz.Document", template_page: "fitz.Page") -> None:
    """在文档末尾追加一张与模板页“尺寸、方向完全一致”的空白页。

    MediaBox 是未考虑 /Rotate 的原始页面框，page.rect 才是显示尺寸。
    用 MediaBox 宽高建页并把 rotation 设为与模板页相同，
    这样空白页的显示尺寸、横竖方向和打印结果都与模板页完全一致。
    """
    rotation = template_page.rotation % 360
    mediabox = template_page.mediabox
    blank = doc.new_page(width=mediabox.width, height=mediabox.height)
    if rotation:
        blank.set_rotation(rotation)


def build_range_pdf(src_path: str, abs_pages: List[int], out_path: str) -> Tuple[int, bool]:
    """按绝对页码提取页面生成临时范围 PDF，并在需要时补一张空白页。

    返回 (范围 PDF 总页数, 是否追加了空白页)。
    ★ 补页判断：以“选定范围的总页数”为准，奇数页时末尾补 1 张空白页。
    """
    src = open_pdf(src_path)
    out = fitz.open()
    try:
        for abs_no in abs_pages:
            out.insert_pdf(src, from_page=abs_no - 1, to_page=abs_no - 1)

        total = out.page_count
        added_blank = False
        if total % 2 == 1:
            append_blank_page_like(out, out[total - 1])
            added_blank = True
            total = out.page_count

        out.save(out_path, garbage=3, deflate=True)
        return total, added_blank
    except ToolError:
        raise
    except Exception as exc:
        raise PdfError(f"处理 PDF 页面时出错（文件可能损坏或不支持）：\n{src_path}\n{exc}") from exc
    finally:
        out.close()
        src.close()


# ----------------------------------------------------------------------------
# 打印顺序（相对页码）
# ----------------------------------------------------------------------------
def compute_duplex_orders(
    total_sides: int, copies: int, collate: bool = True, reverse: bool = False
) -> Tuple[List[int], List[int]]:
    """计算手动双面两阶段的“输出面序号”顺序。

    前提：total_sides 为偶数（调用方已补页）。
    阶段一 = 奇数面；阶段二 = 偶数面逆序。

    份数：
      * collate=True（整份连续）：每份从第 1 面开始，[1,3,5, 1,3,5]；
      * collate=False（逐页重复）：同一面连续打印多份，[1,1,3,3,5,5]。
    反面顺序 = 正面顺序整体反转后 +1，保证每张纸正背配对。

    reverse=True（逆序打印）：把正面顺序整体反转，反面顺序随之重算，
    整体出纸顺序反向，但配对关系不变。
    """
    copies = max(1, int(copies))
    odd_sides = [p for p in range(1, total_sides + 1) if p % 2 == 1]
    if collate:
        fronts = odd_sides * copies
    else:
        fronts = [p for p in odd_sides for _ in range(copies)]
    if reverse:
        fronts = list(reversed(fronts))
    # 正面第 p 面与背面第 p+1 面在同一张纸上；翻转纸叠后顺序相反
    backs = [p + 1 for p in reversed(fronts)]
    return fronts, backs


def compute_print_orders(total_pages: int, copies: int) -> Tuple[List[int], List[int]]:
    """兼容旧接口：默认整份连续、正序。"""
    return compute_duplex_orders(total_pages, copies, collate=True, reverse=False)


# ----------------------------------------------------------------------------
# 整批一次翻面（多文件合并为一份总任务）
# ----------------------------------------------------------------------------
def merged_side_ranges(file_side_counts: List[int]) -> List[Tuple[int, int]]:
    """计算每个文件在“合并任务”中的全局页码范围（1 基，含两端）。

    每个文件各自的排版 PDF 已经是偶数页（见 build_layout_plan 的补页规则），
    所以合并后每个文件的起始全局页码一定是奇数，全局奇偶性与文件内相对
    奇偶性完全对齐——这是“整批一次翻面”能保持配对关系的前提。
    """
    ranges: List[Tuple[int, int]] = []
    start = 1
    for count in file_side_counts:
        count = max(0, int(count))
        if count <= 0:
            ranges.append((0, 0))
            continue
        ranges.append((start, start + count - 1))
        start += count
    return ranges


def compute_merged_duplex_orders(
    file_side_counts: List[int],
    copies: int = 1,
    collate: bool = True,
    reverse: bool = False,
) -> Tuple[List[int], List[int]]:
    """整批一次翻面的两阶段全局页序。

    ★ 为什么必须对“合并后的全局页序”整体倒序，而不是每个文件各自倒序：
      整叠纸翻面后，出纸堆上下顺序被整体颠倒。第二阶段必须从全局最后一页
      （其背面是合并任务最后一个偶数位页）开始依次回退，才能让每张纸的
      正面与背面一一配对。若按文件各自逆序，文件 B 的背面会被先送到
      “文件 A 最后一张纸”的背面，整叠顺序全部错位。

    ★ 份数：
      collate=True  → fronts = 全局奇数面 × 份数（[1,3,5, 1,3,5]）；
      collate=False → fronts = 每个奇数面连续重复（[1,1,3,3,5,5]）。
      两种情况都只在全部正面打完后翻一次面，反面 = 正面顺序整体反转后 +1。

    前提：合并后总页数为偶数（每个文件已各自补页保证）。
    """
    total = sum(max(0, int(count)) for count in file_side_counts)
    if total <= 0:
        raise PdfError("没有需要打印的页面。")
    if total % 2 != 0:
        raise PdfError("合并任务的总页数必须为偶数（每个文件应先补齐空白页）。")

    copies = max(1, int(copies))
    odd_sides = [p for p in range(1, total + 1) if p % 2 == 1]
    if collate:
        fronts = odd_sides * copies
    else:
        fronts = [p for p in odd_sides for _ in range(copies)]
    if reverse:
        fronts = list(reversed(fronts))
    # 每张纸正面为 p、背面为 p+1；整叠翻面后出纸顺序整体相反
    backs = [p + 1 for p in reversed(fronts)]
    return fronts, backs


def merge_layout_pdfs(paths: List[str], out_path: str) -> Tuple[int, List[int]]:
    """按列表顺序把多个排版 PDF 合并为一份“总任务”PDF。

    返回 (总页数, 每个文件的面数)。合并前会校验每个文件面数为偶数，
    因为整批一次翻面的配对关系依赖“每个文件都从奇数全局页开始”。
    """
    if fitz is None:
        raise PdfError("缺少 PyMuPDF 库，请执行：pip install PyMuPDF")
    if not paths:
        raise PdfError("没有可合并的 PDF。")

    counts: List[int] = []
    merged = fitz.open()
    try:
        for path in paths:
            src = fitz.open(path)
            try:
                if not src.is_pdf or src.page_count <= 0:
                    raise PdfError(f"无法合并该 PDF（没有任何页面）：{path}")
                if src.page_count % 2 != 0:
                    raise PdfError(
                        f"文件面数为奇数，无法参与整批双面：{path}\n"
                        "（正常情况下排版阶段应已补齐空白页，请反馈此问题）"
                    )
                merged.insert_pdf(src)
                counts.append(src.page_count)
            finally:
                src.close()
        merged.save(out_path, garbage=3, deflate=True)
        return merged.page_count, counts
    except ToolError:
        raise
    except Exception as exc:
        raise PdfError(f"合并打印任务时出错：{exc}") from exc
    finally:
        merged.close()


def merged_sheet_count(file_side_counts: List[int]) -> int:
    """合并任务的纸张数（每张纸 2 个输出面）。"""
    total = sum(max(0, int(count)) for count in file_side_counts)
    return total // 2


def compute_merged_per_copy_orders(
    file_side_counts: List[int],
    copies: int,
    reverse: bool = False,
) -> List[Tuple[int, List[int], List[int]]]:
    """份数>1 且选择“按份逐份执行”时，每一份的两阶段页序。

    返回 [(份号, 正面页序, 反面页序)]。
    每份独立成叠、各翻一次面，适合需要把每一份分开取走的场景；
    代价是翻面次数 = 份数（默认“整批一次翻面”只翻 1 次）。
    """
    copies = max(1, int(copies))
    fronts, backs = compute_merged_duplex_orders(
        file_side_counts, copies=1, collate=True, reverse=reverse
    )
    return [(index + 1, list(fronts), list(backs)) for index in range(copies)]


def plan_batch_summary(file_side_counts: List[int], copies: int = 1,
                       exec_mode: str = "batch", copy_exec: str = "all",
                       reverse: bool = False) -> dict:
    """生成整批执行计划摘要（供预览区与日志使用，纯数据便于单测）。"""
    copies = max(1, int(copies))
    total_pages = sum(max(0, int(count)) for count in file_side_counts)
    sheets = merged_sheet_count(file_side_counts)
    if exec_mode == "batch":
        if copies > 1 and copy_exec == "per_copy":
            items = compute_merged_per_copy_orders(file_side_counts, copies, reverse=reverse)
            fronts, backs = items[0][1], items[0][2]
            flip_count = copies
            note = f"按份逐份执行：每份正面→翻面→反面，共翻面 {flip_count} 次"
        else:
            fronts, backs = compute_merged_duplex_orders(
                file_side_counts, copies, collate=True, reverse=reverse)
            flip_count = 1
            note = "整批一次翻面：全部正面打完 → 翻面 1 次 → 全部反面打完"
    else:
        fronts, backs = compute_merged_duplex_orders(
            file_side_counts, 1, collate=True, reverse=reverse)
        flip_count = max(1, len([c for c in file_side_counts if int(c) > 0]))
        note = f"逐文件翻面：每个文件单独翻面，共翻面 {flip_count} 次"
    return {
        "total_pages": total_pages,
        "sheets": sheets,
        "fronts": fronts,
        "backs": backs,
        "flip_count": flip_count,
        "note": note,
    }


def compute_single_sided_orders(
    total_sides: int, copies: int, collate: bool = True, reverse: bool = False
) -> List[int]:
    """单面打印顺序：所有输出面按顺序打印。"""
    copies = max(1, int(copies))
    base = list(range(1, total_sides + 1))
    if collate:
        order = base * copies
    else:
        order = [p for p in base for _ in range(copies)]
    return list(reversed(order)) if reverse else order


# ----------------------------------------------------------------------------
# 打印范围模式 → 绝对页码
# ----------------------------------------------------------------------------
def resolve_range_pages(
    mode: str, custom_text: str, page_count: int, current_page: int = 1
) -> Tuple[List[int], List[int]]:
    """把“范围模式”解析为绝对页码。

    返回 (有效页码（升序去重）, 超出文件页数被忽略的页码)。
      all     全部页面
      current 预览中的当前页（多文件时按同一相对页号，越界则取最后一页）
      custom  自定义范围（支持 1-5,8,10-12 与中文标点）
      odd     仅原文档奇数页
      even    仅原文档偶数页
      recent  最近使用范围（等价于 custom，文本由界面传入）
    """
    if page_count < 1:
        raise PdfError("该 PDF 没有可打印页面。")

    if mode == "all":
        return list(range(1, page_count + 1)), []
    if mode == "current":
        page = max(1, min(int(current_page or 1), page_count))
        return [page], []
    if mode in ("custom", "recent"):
        pages = parse_page_range(custom_text)
        valid = [p for p in pages if p <= page_count]
        dropped = [p for p in pages if p > page_count]
        return valid, dropped
    if mode == "odd":
        return [p for p in range(1, page_count + 1) if p % 2 == 1], []
    if mode == "even":
        return [p for p in range(2, page_count + 1, 2)], []
    raise PdfError(f"未知的打印范围模式：{mode}")


# ----------------------------------------------------------------------------
# 版式选项
# ----------------------------------------------------------------------------
@dataclass
class LayoutOptions:
    """打印版式设置。

    页边距与裁切单位均为毫米；纸张方向 auto 表示跟随首页方向。
    """
    paper_name: str = "A4"
    paper_width_mm: float = 210.0
    paper_height_mm: float = 297.0
    orientation: str = "auto"        # auto / portrait / landscape
    source_name: str = "自动选择"
    source_code: int = 7             # DMBIN_AUTO
    auto_rotate: bool = True
    scale_mode: str = "fit"           # fit / shrink / actual / custom / paper
    scale_percent: int = 100
    nup: int = 1
    nup_border: bool = False
    page_order: str = "row"           # row（从左到右）/ column（从上到下）
    booklet: bool = False
    # 默认 0：与 settings_store 默认及 Acrobat“直接打印”口径一致——
    # 只在驱动返回的可打印区域内按需收缩，不额外加白边。
    margin_left_mm: float = 0.0
    margin_top_mm: float = 0.0
    margin_right_mm: float = 0.0
    margin_bottom_mm: float = 0.0
    # 驱动返回的可打印区域内缩（点，左/上/右/下）；None 表示未知。
    # 排版时与用户页边距取最大值，保证内容永远不会落到打印机的不可打印区，
    # 这也是 Acrobat“收缩到可打印区”的实际口径。
    printable_insets_pt: Optional[Tuple[float, float, float, float]] = None
    crop_left_mm: float = 0.0
    crop_top_mm: float = 0.0
    crop_right_mm: float = 0.0
    crop_bottom_mm: float = 0.0
    manual_duplex: bool = True       # True：手动双面流程；False：单面打印
    annots: bool = True              # 是否输出 PDF 批注/表单/戳记
    skip_blank: bool = False         # 忽略空白页

    def margins_pt(self) -> Tuple[float, float, float, float]:
        """返回（左、上、右、下）页边距，单位点。"""
        return (
            mm_to_pt(self.margin_left_mm),
            mm_to_pt(self.margin_top_mm),
            mm_to_pt(self.margin_right_mm),
            mm_to_pt(self.margin_bottom_mm),
        )

    def effective_margins_pt(self) -> Tuple[float, float, float, float]:
        """页边距与驱动可打印区的并集（取每边最大值）。

        用户页边距是“留白需求”，驱动内缩是“物理不可打印”，两者必须同时满足，
        因此逐边取最大值；驱动未返回可打印区时等同于用户页边距。
        """
        left, top, right, bottom = self.margins_pt()
        insets = self.printable_insets_pt
        if not insets:
            return left, top, right, bottom
        return (
            max(left, float(insets[0])),
            max(top, float(insets[1])),
            max(right, float(insets[2])),
            max(bottom, float(insets[3])),
        )

    def crop_pt(self) -> Tuple[float, float, float, float]:
        """返回（左、上、右、下）内容裁切量，单位点。"""
        return (
            mm_to_pt(self.crop_left_mm),
            mm_to_pt(self.crop_top_mm),
            mm_to_pt(self.crop_right_mm),
            mm_to_pt(self.crop_bottom_mm),
        )


def paper_size_pt(options: LayoutOptions, first_page_landscape: bool = False) -> Tuple[float, float]:
    """计算最终纸张尺寸（点），并应用方向设置。

    ★ 同一打印任务只能使用一种纸张尺寸，所以“自动”方向以首页为准；
      混合横竖页通过“自动旋转”在槽位内部转 90° 解决。
    """
    if options.paper_name in PAPER_SIZES_MM:
        width_mm, height_mm = PAPER_SIZES_MM[options.paper_name]
    else:
        width_mm = float(options.paper_width_mm or 210.0)
        height_mm = float(options.paper_height_mm or 297.0)

    width_pt, height_pt = mm_to_pt(width_mm), mm_to_pt(height_mm)
    orientation = options.orientation
    if orientation == "portrait":
        width_pt, height_pt = min(width_pt, height_pt), max(width_pt, height_pt)
    elif orientation == "landscape":
        width_pt, height_pt = max(width_pt, height_pt), min(width_pt, height_pt)
    else:  # auto：跟随首页
        if first_page_landscape and width_pt < height_pt:
            width_pt, height_pt = height_pt, width_pt
        elif not first_page_landscape and width_pt > height_pt:
            width_pt, height_pt = height_pt, width_pt
    return width_pt, height_pt


def nup_grid(nup: int) -> Tuple[int, int]:
    """每张多页的网格（列, 行）。"""
    return {
        1: (1, 1), 2: (2, 1), 4: (2, 2), 6: (3, 2), 9: (3, 3), 16: (4, 4),
    }.get(int(nup), (1, 1))


def slot_rects(
    paper_w: float,
    paper_h: float,
    nup: int,
    options: LayoutOptions,
    gap_mm: float = 0.0,
) -> List["fitz.Rect"]:
    """生成每个版式槽位的矩形（点）。

    槽位间距 gap 用于 N-up 分隔；四周使用用户设置的页边距。
    排列顺序：row = 从左到右、从上到下；column = 从上到下、从左到右。
    """
    cols, rows = nup_grid(nup)
    left, top, right, bottom = options.effective_margins_pt()
    gap = mm_to_pt(gap_mm)
    usable_w = max(1.0, paper_w - left - right - gap * (cols - 1))
    usable_h = max(1.0, paper_h - top - bottom - gap * (rows - 1))
    slot_w = usable_w / cols
    slot_h = usable_h / rows

    cells = [(r, c) for r in range(rows) for c in range(cols)]
    if options.page_order == "column":
        cells.sort(key=lambda rc: (rc[1], rc[0]))

    rects = []
    for row, col in cells[: max(1, int(nup))]:
        x0 = left + col * (slot_w + gap)
        y0 = top + row * (slot_h + gap)
        rects.append(fitz.Rect(x0, y0, x0 + slot_w, y0 + slot_h))
    return rects


def compute_desired_rect(
    src_rect: "fitz.Rect", slot: "fitz.Rect", scale_mode: str, percent: float
) -> "fitz.Rect":
    """计算页面在槽位中的“目标矩形”（居中、等比，未做裁切）。

    与 compute_placement 的区别：这里不裁掉超出槽位的部分，便于预览显示
    “内容实际会比可打印区更大（会被裁切）”的真实情况。
    """
    sw, sh = src_rect.width, src_rect.height
    if sw <= 0 or sh <= 0:
        return fitz.Rect(slot)
    if scale_mode == "paper":
        return fitz.Rect(slot)
    if scale_mode == "fit":
        scale = min(slot.width / sw, slot.height / sh)
    elif scale_mode == "shrink":
        scale = min(1.0, slot.width / sw, slot.height / sh)
    elif scale_mode == "actual":
        scale = 1.0
    elif scale_mode == "custom":
        scale = max(0.01, float(percent) / 100.0)
    else:
        scale = min(slot.width / sw, slot.height / sh)
    dest_w, dest_h = sw * scale, sh * scale
    return fitz.Rect(
        slot.x0 + (slot.width - dest_w) / 2,
        slot.y0 + (slot.height - dest_h) / 2,
        slot.x0 + (slot.width + dest_w) / 2,
        slot.y0 + (slot.height + dest_h) / 2,
    )


def compute_placement(
    src_rect: "fitz.Rect", slot: "fitz.Rect", scale_mode: str, percent: float
) -> Tuple[Optional["fitz.Rect"], Optional["fitz.Rect"], bool]:
    """计算页面在槽位中的目标矩形、源内容裁剪矩形与是否保持比例。

    返回 (dest, clip, keep_proportion)：
      * dest 为 None 表示完全在槽位外；
      * clip 为 None 表示无需裁剪；
      * 实际大小/自定义比例超出槽位时，对称裁掉超出部分（不压到相邻槽位）。
    """
    sw, sh = src_rect.width, src_rect.height
    if sw <= 0 or sh <= 0:
        return None, None, True

    if scale_mode == "paper":
        # 按纸张缩放：拉伸填满槽位（不保持宽高比）
        return fitz.Rect(slot), None, False

    if scale_mode == "fit":
        scale = min(slot.width / sw, slot.height / sh)
    elif scale_mode == "shrink":
        scale = min(1.0, slot.width / sw, slot.height / sh)
    elif scale_mode == "actual":
        scale = 1.0
    elif scale_mode == "custom":
        scale = max(0.01, float(percent) / 100.0)
    else:
        scale = min(slot.width / sw, slot.height / sh)

    dest_w, dest_h = sw * scale, sh * scale
    centered = fitz.Rect(
        slot.x0 + (slot.width - dest_w) / 2,
        slot.y0 + (slot.height - dest_h) / 2,
        slot.x0 + (slot.width + dest_w) / 2,
        slot.y0 + (slot.height + dest_h) / 2,
    )
    visible = centered & slot
    if visible.is_empty or visible.width <= 0.001 or visible.height <= 0.001:
        return None, None, True

    clip = fitz.Rect(
        src_rect.x0 + (visible.x0 - centered.x0) / scale,
        src_rect.y0 + (visible.y0 - centered.y0) / scale,
        src_rect.x0 + (visible.x1 - centered.x0) / scale,
        src_rect.y0 + (visible.y1 - centered.y0) / scale,
    )
    # 四边都完整显示时不需要 clip
    if (abs(clip.x0 - src_rect.x0) < 0.01 and abs(clip.y0 - src_rect.y0) < 0.01
            and abs(clip.x1 - src_rect.x1) < 0.01 and abs(clip.y1 - src_rect.y1) < 0.01):
        clip = None
    return visible, clip, True


def preview_metrics(
    plan: "LayoutPlan",
    options: LayoutOptions,
    page_w: float,
    page_h: float,
) -> Dict[str, object]:
    """计算“成纸预览”所需的几何数据（纯函数，便于单元测试）。

    返回：
      paper      纸张尺寸（点）
      slots      每个版式槽位的矩形 (x0, y0, x1, y1)，按页序排列
      nup        实际每张页数（小册子固定 2）
      scale_percent  当前缩放模式下页面的实际缩放百分比（如 90.5）
      overflow   内容是否超出槽位（实际大小/自定义比例过大时会被裁掉）
      centered   是否水平/垂直居中（缩放类模式都会居中）
      printable  可打印区域矩形（相对纸张左上角，点）
    """
    paper_w, paper_h = plan.paper_pt
    left, top, right, bottom = options.effective_margins_pt()
    nup = 2 if options.booklet else max(1, int(plan.nup))
    slots = slot_rects(paper_w, paper_h, nup, options)
    slot = slots[0] if slots else fitz.Rect(left, top, paper_w - right, paper_h - bottom)

    src = fitz.Rect(0, 0, max(1.0, page_w), max(1.0, page_h))
    desired = compute_desired_rect(src, slot, options.scale_mode, options.scale_percent)
    percent = 0.0
    if src.width > 0:
        percent = 100.0 * desired.width / src.width
    overflow = (desired.width > slot.width + 0.01
                or desired.height > slot.height + 0.01)
    printable = (left, top, paper_w - right, paper_h - bottom)
    return {
        "paper": (paper_w, paper_h),
        "slots": [(r.x0, r.y0, r.x1, r.y1) for r in slots],
        "nup": nup,
        "scale_percent": percent,
        "overflow": overflow,
        "centered": True,
        "printable": printable,
    }


# ----------------------------------------------------------------------------
# 版式计划（N-up / 小册子 / 双面补页）
# ----------------------------------------------------------------------------
def plan_nup_sides(count: int, nup: int) -> List[List[Optional[int]]]:
    """把源页序分配到 N-up 输出面。

    返回每个输出面上放置的源页序号（1 基），空白槽位为 None。
    最后一个面不足 N 页时用 None 补齐，保证槽位数量固定。
    注：槽位本身的行列顺序由 slot_rects 依据 page_order 决定，与这里无关。
    """
    nup = max(1, int(nup))
    pages: List[Optional[int]] = list(range(1, count + 1))
    sides: List[List[Optional[int]]] = []
    for start in range(0, len(pages), nup):
        chunk = pages[start:start + nup]
        chunk += [None] * (nup - len(chunk))
        sides.append(chunk)
    return sides


def plan_booklet_sides(count: int) -> List[List[Optional[int]]]:
    """计算小册子（骑马钉）的每个输出面。

    约定 count 为 4 的倍数。输出面顺序为：
      第 1 张纸正面、第 1 张纸背面、第 2 张纸正面、第 2 张纸背面 …
    这样正好可以套用手动双面流程：
      阶段一按正序打印所有“正面”，整叠翻面后阶段二逆序打印所有“背面”。

    每张纸的 4 个页面：
      正面左 = N-2i，正面右 = 1+2i；
      背面左 = 2+2i，背面右 = N-1-2i（i 从 0 开始）。
    """
    if count % 4 != 0:
        raise PdfError("小册子排版要求页数为 4 的倍数（调用方应先补空白页）。")
    sides: List[List[Optional[int]]] = []
    sheets = count // 4
    for i in range(sheets):
        front_left = count - 2 * i
        front_right = 1 + 2 * i
        back_left = 2 + 2 * i
        back_right = count - 1 - 2 * i
        sides.append([front_left, front_right])
        sides.append([back_left, back_right])
    return sides


@dataclass
class LayoutPlan:
    """一次打印的排版计划（纯数据，便于单元测试与界面预览）。"""
    total_selected: int
    padded_count: int
    source_blank_added: int
    duplex_blank_added: bool
    sides: List[List[Optional[int]]]
    paper_pt: Tuple[float, float]
    nup: int
    booklet: bool
    output_blank_added: int
    manual_duplex: bool

    @property
    def output_count(self) -> int:
        return len(self.sides)

    def side_label(self, side_no: int, abs_pages: Optional[List[int]] = None) -> str:
        """输出面的可读标签，例如 "1、2（原3、4）" 或 "空白(补)"。"""
        if side_no < 1 or side_no > len(self.sides):
            return "—"
        entries = self.sides[side_no - 1]
        labels = []
        for rel in entries:
            if rel is None:
                continue
            if abs_pages and rel > len(abs_pages):
                labels.append("空白(补)")
            elif abs_pages:
                labels.append(f"{rel}·原{abs_pages[rel - 1]}")
            else:
                labels.append(str(rel))
        if not labels:
            return "空白(补)"
        return "、".join(labels)

    def output_blank_text(self, side_no: int) -> bool:
        """该输出面是否全部为空白槽位。"""
        entries = self.sides[side_no - 1] if 1 <= side_no <= len(self.sides) else []
        return all(rel is None for rel in entries)


def build_layout_plan(
    selected_count: int, options: LayoutOptions, first_page_landscape: bool = False
) -> LayoutPlan:
    """根据选定源页数生成排版计划。

    补页顺序：
      1) 手动双面要求总页数为偶数 → 奇数时补 1 张空白（核心规则）；
      2) 小册子要求页数为 4 的倍数 → 再补 0-3 张；
      3) N-up 后若输出面仍为奇数（双面无法配对）→ 末尾再补 1 个空白输出面。
    """
    if selected_count < 1:
        raise PdfError("选定范围至少需要 1 页。")

    n = int(selected_count)
    duplex_pad = 1 if (options.manual_duplex and n % 2 == 1) else 0
    base = n + duplex_pad
    booklet_pad = (-base) % 4 if options.booklet else 0
    padded = base + booklet_pad

    if options.booklet:
        sides = plan_booklet_sides(padded)
    elif options.nup > 1:
        sides = plan_nup_sides(padded, options.nup)
    else:
        sides = [[i] for i in range(1, padded + 1)]

    output_blank = 0
    if options.manual_duplex and len(sides) % 2 == 1:
        # 例如 4 页文本按 4-up 排版只有 1 个输出面，双面需要再补 1 个空白面
        sides.append([None] * max(1, int(options.nup)))
        output_blank = 1

    width, height = paper_size_pt(options, first_page_landscape)
    return LayoutPlan(
        total_selected=n,
        padded_count=padded,
        source_blank_added=padded - n,
        duplex_blank_added=bool(duplex_pad),
        sides=sides,
        paper_pt=(width, height),
        nup=max(1, int(options.nup)),
        booklet=bool(options.booklet),
        output_blank_added=output_blank,
        manual_duplex=bool(options.manual_duplex),
    )


def describe_plan(plan: LayoutPlan) -> str:
    """生成适合写入日志的版式说明。"""
    parts = [f"选取 {plan.total_selected} 页"]
    if plan.source_blank_added:
        text = f"补 {plan.source_blank_added} 张空白页 → {plan.padded_count} 页"
        if plan.duplex_blank_added:
            text += "（其中 1 张为双面配对补页）"
        parts.append(text)
    if plan.booklet:
        parts.append(f"小册子排版 → {plan.output_count} 个输出面")
    elif plan.nup > 1:
        parts.append(f"{plan.nup} 合 1 → {plan.output_count} 个输出面")
    else:
        parts.append(f"{plan.output_count} 个输出面")
    if plan.output_blank_added:
        parts.append("另补 1 个空白输出面用于双面配对")
    return "，".join(parts)


# ----------------------------------------------------------------------------
# 裁切：显示坐标 → 原始页面坐标
# ----------------------------------------------------------------------------
def display_rect_to_media_rect(page: "fitz.Page", rect: "fitz.Rect") -> "fitz.Rect":
    """把“显示方向”的矩形换算成 PDF 原始（未旋转）坐标。

    页面的 /Rotate 会让显示宽高与 MediaBox 交换；裁切边距按用户看到的
    方向计算，因此需要按 90/180/270 做坐标映射。
    """
    media = page.mediabox
    width, height = media.width, media.height
    rotation = page.rotation % 360
    x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
    if rotation == 90:
        # 显示 x = H - y，显示 y = x
        return fitz.Rect(y0, height - x1, y1, height - x0)
    if rotation == 180:
        return fitz.Rect(width - x1, height - y1, width - x0, height - y0)
    if rotation == 270:
        # 显示 x = y，显示 y = W - x
        return fitz.Rect(width - y1, x0, width - y0, x1)
    return fitz.Rect(x0, y0, x1, y1)


def crop_rect_for_page(page: "fitz.Page", options: LayoutOptions) -> Optional["fitz.Rect"]:
    """按裁切设置计算原始坐标下的裁剪矩形；无裁切时返回 None。"""
    left, top, right, bottom = options.crop_pt()
    if max(left, top, right, bottom) <= 0.01:
        return None
    display = fitz.Rect(0, 0, page.rect.width, page.rect.height)
    inner = fitz.Rect(
        display.x0 + left, display.y0 + top, display.x1 - right, display.y1 - bottom
    )
    if inner.width <= 1 or inner.height <= 1:
        return None
    media_rect = display_rect_to_media_rect(page, inner)
    media = page.mediabox
    # 夹到 MediaBox 内，避免驱动/PyMuPDF 拒绝非法裁剪框
    return fitz.Rect(
        max(media.x0, media_rect.x0), max(media.y0, media_rect.y0),
        min(media.x1, media_rect.x1), min(media.y1, media_rect.y1),
    )


# ----------------------------------------------------------------------------
# 生成排版后的临时 PDF
# ----------------------------------------------------------------------------
@dataclass
class ComposeResult:
    out_path: str
    plan: LayoutPlan
    selected_count: int
    skipped_blank: int
    abs_pages: List[int] = field(default_factory=list)


def _standalone_page(src: "fitz.Document", pno: int, crop_rect: Optional["fitz.Rect"],
                     rotate: bool) -> "fitz.Document":
    """把源页复制成独立单页文档，可选裁切与 90° 旋转。

    先复制再裁切，避免修改用户原文件；裁切后 page.rect 以 (0,0) 为原点，
    后续 show_pdf_page 无需再关心原始 MediaBox 偏移。
    """
    tmp = fitz.open()
    tmp.insert_pdf(src, from_page=pno, to_page=pno)
    if crop_rect is not None:
        tmp[0].set_cropbox(crop_rect)
    if not rotate:
        return tmp

    rotated = fitz.open()
    page = rotated.new_page(width=tmp[0].rect.height, height=tmp[0].rect.width)
    page.show_pdf_page(page.rect, tmp, 0, rotate=90)
    tmp.close()
    return rotated


def _place_source(
    out_page: "fitz.Page",
    src: "fitz.Document",
    pno: int,
    slot: "fitz.Rect",
    options: LayoutOptions,
) -> None:
    """把源页放入槽位，处理裁切、自动旋转、缩放与超界裁剪。"""
    page = src[pno]
    crop_rect = crop_rect_for_page(page, options)

    rotate = False
    if options.auto_rotate:
        display_w = page.rect.width - sum(options.crop_pt()[0::2])
        display_h = page.rect.height - sum(options.crop_pt()[1::2])
        if display_w <= 0 or display_h <= 0:
            display_w, display_h = page.rect.width, page.rect.height
        if (display_w > display_h) != (slot.width > slot.height):
            rotate = True

    standalone = _standalone_page(src, pno, crop_rect, rotate)
    try:
        src_rect = standalone[0].rect
        dest, clip, keep_proportion = compute_placement(
            src_rect, slot, options.scale_mode, options.scale_percent
        )
        if dest is None:
            return
        kwargs = {"keep_proportion": keep_proportion}
        if clip is not None:
            kwargs["clip"] = clip
        out_page.show_pdf_page(dest, standalone, 0, **kwargs)
    finally:
        standalone.close()


def compose_layout_pdf(
    src_path: str,
    abs_pages: List[int],
    out_path: str,
    options: LayoutOptions,
) -> ComposeResult:
    """按版式设置生成用于打印的临时 PDF。

    流程：过滤空白页 → 补页计划 → 生成每个输出面 → 逐槽位放置源页。
    """
    src = open_pdf(src_path)
    out = fitz.open()
    try:
        pages = list(abs_pages)
        skipped = 0
        if options.skip_blank:
            filtered = []
            for abs_no in pages:
                if page_is_blank(src[abs_no - 1]):
                    skipped += 1
                    continue
                filtered.append(abs_no)
            pages = filtered
        if not pages:
            raise PdfError("选定范围内全部是空白页（已按“忽略空白页”设置跳过）。")

        first_landscape = src[pages[0] - 1].rect.width > src[pages[0] - 1].rect.height
        plan = build_layout_plan(len(pages), options, first_landscape)
        paper_w, paper_h = plan.paper_pt

        for side in plan.sides:
            out_page = out.new_page(width=paper_w, height=paper_h)
            if plan.booklet:
                slots = slot_rects(paper_w, paper_h, 2, options, gap_mm=0.0)
            else:
                slots = slot_rects(paper_w, paper_h, plan.nup, options, gap_mm=2.0)
            for index, rel in enumerate(side):
                if index >= len(slots):
                    break
                slot = slots[index]
                if options.nup_border:
                    out_page.draw_rect(slot, color=(0.55, 0.55, 0.55), width=0.6)
                if rel is None or rel > len(pages):
                    continue
                abs_no = pages[rel - 1]
                _place_source(out_page, src, abs_no - 1, slot, options)

        out.save(out_path, garbage=3, deflate=True)
        return ComposeResult(
            out_path=out_path,
            plan=plan,
            selected_count=len(pages),
            skipped_blank=skipped,
            abs_pages=pages,
        )
    except ToolError:
        raise
    except Exception as exc:
        raise PdfError(f"生成排版 PDF 时出错（文件可能损坏或不支持）：\n{src_path}\n{exc}") from exc
    finally:
        out.close()
        src.close()
