# -*- coding: utf-8 -*-
"""生成 Word 打印对比实测样本（v2：改用 Selection 逐段写入，避免 Range.Text 覆盖段落）。

覆盖：中英混排 / 纯中文 / 纯英文 / 表格 / 图片 / 公式与特殊符号 / 页眉页脚页码 /
旧格式 .doc / 缺失字体 / 35 页长文。生成后由 inspect_samples.py 统计页数与字体。
"""
import os
import time

import pythoncom
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE, "samples")
IMG_DIR = os.path.join(BASE, "assets")

WD_FORMAT_DOCX = 16
WD_FORMAT_DOC97 = 0
WD_ALIGN_CENTER = 1
WD_STORY_END = 6
WD_PAGE_BREAK = 7
WD_FIELD_PAGE = 33
WD_FIELD_NUMPAGES = 26


def make_image():
    from PIL import Image, ImageDraw
    path = os.path.join(IMG_DIR, "chart.png")
    os.makedirs(IMG_DIR, exist_ok=True)
    img = Image.new("RGB", (900, 560), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([20, 20, 880, 540], outline=(60, 90, 140), width=3)
    for i, h in enumerate([180, 320, 260, 420, 360, 480]):
        x0 = 90 + i * 120
        draw.rectangle([x0, 500 - h, x0 + 70, 500], fill=(70, 120, 200))
    draw.line([70, 500, 860, 500], fill=(40, 40, 40), width=2)
    draw.text((70, 30), "Sample chart 图表 sample", fill=(30, 30, 30))
    for i, label in enumerate(["Q1", "Q2", "Q3", "Q4", "Q5", "Q6"]):
        draw.text((100 + i * 120, 512), label, fill=(30, 30, 30))
    img.save(path)
    return path


def add_para(word, text, font="宋体", size=10.5, bold=False, align=None, space_after=6):
    sel = word.Selection
    sel.EndKey(WD_STORY_END)
    sel.TypeText(text)
    sel.TypeParagraph()
    doc = word.ActiveDocument
    para = doc.Content.Paragraphs(doc.Content.Paragraphs.Count - 1)
    para.Range.Font.Name = font
    para.Range.Font.Size = size
    para.Range.Font.Bold = bold
    para.Format.SpaceAfter = space_after
    if align is not None:
        para.Format.Alignment = align
    return para


def add_page_break(word):
    sel = word.Selection
    sel.EndKey(WD_STORY_END)
    sel.InsertBreak(WD_PAGE_BREAK)


def add_table(word, rows, cols, font="宋体", size=9):
    sel = word.Selection
    sel.EndKey(WD_STORY_END)
    table = word.ActiveDocument.Tables.Add(sel.Range, rows, cols)
    table.Borders.Enable = True
    for r in range(rows):
        for c in range(cols):
            cell = table.Cell(r + 1, c + 1)
            cell.Range.Text = f"第{r + 1}行 第{c + 1}列 Cell({r + 1},{c + 1})"
            cell.Range.Font.Name = font
            cell.Range.Font.Size = size
    sel.EndKey(WD_STORY_END)
    return table


def set_header_footer(doc, header_text):
    section = doc.Sections(1)
    hdr = section.Headers(1).Range
    hdr.Text = header_text
    hdr.Font.Name = "宋体"
    hdr.Font.Size = 9
    hdr.ParagraphFormat.Alignment = WD_ALIGN_CENTER
    ftr = section.Footers(1).Range
    ftr.Text = "第 "
    ftr.Font.Name = "宋体"
    ftr.Font.Size = 9
    ftr.ParagraphFormat.Alignment = WD_ALIGN_CENTER
    ftr.Collapse(0)
    ftr.Fields.Add(ftr, WD_FIELD_PAGE)
    ftr.InsertAfter(" 页 / 共 ")
    ftr.Collapse(0)
    ftr.Fields.Add(ftr, WD_FIELD_NUMPAGES)
    ftr.InsertAfter(" 页")


def new_doc(word):
    doc = word.Documents.Add()
    doc.Styles(-1).Font.Name = "宋体"
    doc.Styles(-1).Font.Size = 10.5
    return doc


def build_mixed_short(word, path):
    doc = new_doc(word)
    add_para(word, "中英文混排短文档 / Mixed Chinese-English Short Document", "微软雅黑", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "本文档用于验证 Word 转 PDF 再打印时中文与英文的排版是否一致。段落中同时出现汉字、"
                   "拉丁字母、数字与标点：The quick brown fox jumps over the lazy dog. 温度 25°C，"
                   "版本 v2.1.3，编号 A-2026-1104。", "宋体", 10.5)
    add_para(word, "第二段使用 Times New Roman 11pt：Font substitution may change glyph shapes; "
                   "kerning and line spacing can drift after export.", "Times New Roman", 11)
    add_para(word, "第三段使用微软雅黑 12pt：字体回退、字距与行距在导出后可能发生漂移，"
                   "需要用像素与字体两层指标交叉验证。", "微软雅黑", 12)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "中英混排短文档（10.5/11/12/16pt）"


def build_chinese_pure(word, path):
    doc = new_doc(word)
    add_para(word, "纯中文技术说明", "宋体", 18, True, WD_ALIGN_CENTER, 14)
    text = ("本软件在打印 Word 文档前，会先用办公软件把文档导出为 PDF，再按纸张尺寸排版，"
            "最后由打印驱动输出。这样做的好处是排版在导出阶段已经固定，打印阶段不再依赖"
            "打印驱动对字体的支持。若打印机驱动缺少某些字体，直接打印 Word 就可能出现方框"
            "或字形替换；先转 PDF 则把字体嵌入到文件中，输出更稳定。")
    for i in range(6):
        add_para(word, f"第{i + 1}节：" + text, "宋体", 12)
    add_page_break(word)
    add_para(word, "附录：常见问题", "黑体", 16, True, None, 10)
    for q in ("为什么打印前要转 PDF？", "打印出现方块字怎么办？", "双面打印如何避免错页？"):
        add_para(word, "问：" + q, "黑体", 12, True)
        add_para(word, "答：" + text, "楷体", 12)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "纯中文（宋体/黑体/楷体 12pt）"


def build_english_pure(word, path):
    doc = new_doc(word)
    add_para(word, "English Only Sample Document", "Calibri", 20, True, WD_ALIGN_CENTER, 14)
    body = ("This sample contains English text only. It is used to compare the Word-to-PDF "
            "conversion fidelity of the product pipeline against the built-in Word export. "
            "Paragraphs use Calibri, Times New Roman and Arial at different sizes. ")
    for i in range(8):
        add_para(word, f"Section {i + 1}. " + body * 3, "Calibri", 11)
    add_page_break(word)
    for i in range(6):
        add_para(word, f"Times New Roman part {i + 1}. " + body * 3, "Times New Roman", 12)
    add_page_break(word)
    for i in range(6):
        add_para(word, f"Arial part {i + 1}. " + body * 3, "Arial", 10)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "纯英文（Calibri/Times New Roman/Arial 10-12pt）"


def build_table(word, path):
    doc = new_doc(word)
    add_para(word, "表格与合并单元格", "微软雅黑", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "下表含 12 行 5 列，含表头加粗与中英混排内容。", "宋体", 10.5)
    table = add_table(word, 12, 5, "宋体", 9)
    table.Rows(1).Range.Font.Bold = True
    add_para(word, "第二张表：窄列宽，验证换行与列宽在导出后是否保持一致。", "宋体", 10.5)
    add_table(word, 8, 4, "Calibri", 9)
    add_page_break(word)
    add_para(word, "表格之后的长段落，用于检查分页位置。", "宋体", 10.5)
    for i in range(6):
        add_para(word, ("表格排版依赖列宽、内边距与字体度量，任何一项变化都可能改变换行位置。"
                        "Table layout depends on column width, cell padding and font metrics. ") * 4, "宋体", 10.5)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "表格（12x5、8x4）9pt"


def build_image(word, path, img_path):
    doc = new_doc(word)
    add_para(word, "图片与图文混排", "微软雅黑", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "下图为本机生成的 PNG，插入宽度 12cm，用于验证图片在 PDF 与打印稿中的清晰度。", "宋体", 10.5)
    sel = word.Selection
    sel.EndKey(WD_STORY_END)
    shape = doc.InlineShapes.AddPicture(img_path, False, True, sel.Range)
    shape.Width = 340
    add_para(word, "图像下的说明文字（图 1：示例柱状图）。", "宋体", 9, False, WD_ALIGN_CENTER)
    add_page_break(word)
    add_para(word, "第二页继续放一张缩小的图片，验证缩放插值差异。", "宋体", 10.5)
    sel.EndKey(WD_STORY_END)
    shape2 = doc.InlineShapes.AddPicture(img_path, False, True, sel.Range)
    shape2.Width = 200
    add_para(word, "图片会以位图或矢量形式嵌入 PDF，打印时再按 300 DPI 栅格化。", "宋体", 10.5)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "图片（PNG 12cm/7cm）"


def build_formula(word, path):
    doc = new_doc(word)
    add_para(word, "公式与特殊符号", "微软雅黑", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "下列符号覆盖数学、希腊字母、箭头、圈号与货币符号，全部使用 Cambria Math / 宋体。",
             "宋体", 10.5)
    add_para(word, "∑ ∏ ∫ √ ± ≤ ≥ ≠ ≈ ∞ ∂ ∇ α β γ δ λ μ π σ ω Δ Ω → ← ↑ ↓ ↔ ⇒ ∈ ∉ ⊂ ⊃ ∪ ∩",
             "Cambria Math", 14)
    add_para(word, "圈号与单位：①②③④⑤ ℃ ℉ ™ © ® § ¥ € £ ‰", "宋体", 14)
    add_para(word, "线性公式：E = mc²，a² + b² = c²，x = (-b ± √(b² - 4ac)) / 2a", "Cambria Math", 13)
    note = "公式与特殊符号（特殊符号 + 线性公式），Cambria Math 13-14pt"
    try:
        para = add_para(word, "y = (-b + √(b² - 4ac)) / 2a", "Cambria Math", 14)
        rng = para.Range
        doc.OMaths.Add(rng)
        doc.OMaths(doc.OMaths.Count).BuildUp()
        note = "公式文档（含 OMML 公式与特殊符号），Cambria Math 13-14pt"
    except Exception as exc:
        print("OMML 插入失败，保留线性公式:", exc, flush=True)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return note


def build_header_footer(word, path):
    doc = new_doc(word)
    set_header_footer(doc, "内部资料 · 页眉示例 Internal Use Only")
    add_para(word, "页眉页脚与页码", "微软雅黑", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "本文档正文 5 页，每页都有页眉、页脚和“第 N 页 / 共 M 页”页码域。", "宋体", 11)
    for i in range(4):
        add_page_break(word)
        add_para(word, f"第 {i + 2} 页正文：页眉与页码是打印稿上最容易发生位移的元素，"
                       "需要在像素层比较其位置是否一致。", "宋体", 11)
        for j in range(6):
            add_para(word, f"填充段落 {j + 1}：中文与 English 混排，验证页眉页脚不随正文缩放。", "宋体", 11)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "页眉/页脚/页码 5 页 11pt"


def build_legacy_doc(word, path):
    doc = new_doc(word)
    add_para(word, "旧格式 .doc 样本（Word 97-2003）", "黑体", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "本文件以 Word 97-2003 二进制格式保存，用于验证旧格式在转换链路中的表现。", "宋体", 10.5)
    add_para(word, "Legacy binary format may behave differently from OOXML. 中文与 English 混排。", "Times New Roman", 11)
    add_page_break(word)
    add_table(word, 6, 4, "宋体", 9)
    add_para(word, "第二页：表格 + 正文，验证 .doc 的分页与字体嵌入。", "宋体", 10.5)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOC97)
    doc.Close(0)
    return "旧格式 .doc 2 页"


def build_missing_font(word, path):
    doc = new_doc(word)
    add_para(word, "缺失字体样本 NoSuchFontABC", "宋体", 16, True, WD_ALIGN_CENTER, 12)
    add_para(word, "以下段落被显式设置为本机不存在的字体 NoSuchFontABC，用于验证字体回退。", "宋体", 10.5)
    add_para(word, "NoSuchFontABC 渲染：中文汉字与 Latin ABC 0123 都会经过字体替换。", "宋体", 12)
    add_para(word, "第二段同样使用 NoSuchFontABC：符号 ①②③ ∑ √ 与标点“引号”也要检查。", "宋体", 14)
    add_para(word, "结尾段落恢复宋体，作为对照。", "宋体", 10.5)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "缺失字体 NoSuchFontABC 12/14pt（生成后由 fix_missing_font.py 写入字体名）"


def build_long(word, path):
    doc = new_doc(word)
    set_header_footer(doc, "长文档性能样本 Long Document")
    add_para(word, "35 页长文档（性能样本）", "微软雅黑", 18, True, WD_ALIGN_CENTER, 14)
    base = ("本段用于制造分页压力：中文段落与 English paragraph 交替出现，包含数字 2026、"
            "标点符号、括号与百分比 12.5%。The purpose is to measure conversion and printing "
            "time on a document long enough to amortize cold start. ")
    for page in range(35):
        if page:
            add_page_break(word)
        add_para(word, f"第 {page + 1} 页标题 Page {page + 1}", "黑体", 12, True)
        for i in range(3):
            font = "Times New Roman" if i % 3 == 0 else "宋体"
            add_para(word, base * 3, font, 10.5)
    doc.SaveAs2(path, FileFormat=WD_FORMAT_DOCX)
    doc.Close(0)
    return "35 页长文档 10.5-12pt"


def main():
    os.makedirs(SAMPLE_DIR, exist_ok=True)
    img = make_image()
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    builders = [
        ("01_mixed_short.docx", build_mixed_short, ()),
        ("02_chinese_pure.docx", build_chinese_pure, ()),
        ("03_english_pure.docx", build_english_pure, ()),
        ("04_table.docx", build_table, ()),
        ("05_image.docx", build_image, (img,)),
        ("06_formula.docx", build_formula, ()),
        ("07_header_footer.docx", build_header_footer, ()),
        ("08_legacy.doc", build_legacy_doc, ()),
        ("09_missing_font.docx", build_missing_font, ()),
        ("10_long_35p.docx", build_long, ()),
    ]
    for name, builder, extra in builders:
        path = os.path.join(SAMPLE_DIR, name)
        if os.path.exists(path):
            os.remove(path)
        started = time.perf_counter()
        note = builder(word, path, *extra)
        print("生成", name, note, "%.1fs" % (time.perf_counter() - started), flush=True)
    try:
        word.Quit(0)
    except Exception:
        pass
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
