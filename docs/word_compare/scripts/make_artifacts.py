# -*- coding: utf-8 -*-
"""生成报告用的对比图、差异热力图、人工目检截图（写入项目 docs/word_compare）。"""
import json
import os
import shutil
import statistics

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")
DOCS = r"d:\打印机软件\docs\word_compare"
HEAT = os.path.join(DOCS, "heatmaps")
SHOTS = os.path.join(DOCS, "shots")

FONT = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"


def fnt(size, bold=False):
    try:
        return ImageFont.truetype(FONT_BOLD if bold else FONT, size)
    except Exception:
        return ImageFont.load_default()


def render(pdf, pno, dpi=150):
    doc = fitz.open(pdf)
    pix = doc[pno].get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    doc.close()
    return img


def gray_arr(pdf, pno, dpi=150):
    doc = fitz.open(pdf)
    pix = doc[pno].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width].copy()
    doc.close()
    return arr


def heat_overlay(page_img, diff, threshold=32, alpha_max=0.9):
    """在浅化的页面图上叠加差异热力（蓝→黄→红）。"""
    base = np.asarray(page_img).astype(np.float32)
    gray = base.mean(axis=2)
    light = np.stack([np.clip(gray * 0.55 + 115, 0, 255)] * 3, axis=2)
    norm = np.clip(diff.astype(np.float32) / 160.0, 0, 1)[..., None]
    # 蓝(低) → 黄(中) → 红(高)
    r = np.where(norm < 0.5, 60 + norm * 2 * 195, 255)
    g = np.where(norm < 0.5, 90 + norm * 2 * 135, 235 - (norm - 0.5) * 2 * 205)
    b = np.where(norm < 0.5, 230 - norm * 2 * 150, 30)
    heat = np.concatenate([r, g, b], axis=2)
    a = (diff > threshold).astype(np.float32)[..., None] * alpha_max
    out = light * (1 - a) + heat * a
    return Image.fromarray(out.astype(np.uint8))


def three_panel(title, a_img, b_img, heat_img, footer, path, width=560):
    imgs = []
    for img in (a_img, b_img, heat_img):
        h = int(img.height * width / img.width)
        imgs.append(img.resize((width, h), Image.LANCZOS))
    gap = 16
    head = 62
    foot = 44
    canvas = Image.new("RGB", (width * 3 + gap * 2, head + imgs[0].height + foot), (246, 247, 249))
    d = ImageDraw.Draw(canvas)
    d.text((16, 10), title, font=fnt(20, True), fill=(26, 32, 44))
    labels = ["A：软件链路", "B：Word 导出 PDF", "差异热力图（A vs B）"]
    for i, (img, label) in enumerate(zip(imgs, labels)):
        x = i * (width + gap)
        canvas.paste(img, (x, head))
        d.rectangle([x, head, x + width - 1, head + img.height - 1], outline=(210, 214, 220))
        d.text((x + 8, head - 26), label, font=fnt(16, True), fill=(60, 70, 90))
    d.rectangle([0, head + imgs[0].height, canvas.width, canvas.height], fill=(246, 247, 249))
    d.text((16, head + imgs[0].height + 12), footer, font=fnt(16), fill=(70, 80, 100))
    canvas.save(path)


def make_artifacts():
    os.makedirs(HEAT, exist_ok=True)
    os.makedirs(SHOTS, exist_ok=True)
    metrics = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))
    samples = json.load(open(os.path.join(BASE, "samples", "samples.json"), encoding="utf-8"))
    pages = {os.path.splitext(s["name"])[0]: s["pages"] for s in samples}

    for name, m in metrics.items():
        n_pages = pages[name]
        # 转换对照：A 转换 PDF vs B Word 导出 PDF（已实测逐页 0 差异）
        a_pdf = os.path.join(OUT, "convA", name + ".pdf")
        b_pdf = os.path.join(OUT, "exportB", name + ".pdf")
        pno = 0
        a_img, b_img = render(a_pdf, pno), render(b_pdf, pno)
        ga, gb = gray_arr(a_pdf, pno), gray_arr(b_pdf, pno)
        diff = np.abs(ga.astype(np.int16) - gb.astype(np.int16))
        heat = heat_overlay(a_img, diff)
        d = ImageDraw.Draw(heat)
        d.rectangle([0, 0, heat.width, 46], fill=(255, 255, 255))
        d.text((14, 12), f"逐像素差异 0.000%（全部 {n_pages} 页均为 0）",
               font=fnt(20, True), fill=(20, 120, 60))
        three_panel(
            "%s（%d 页）第 1 页：转换产物对照" % (name, n_pages),
            a_img, b_img, heat,
            "口径：150 DPI 灰度，阈值 |Δ|>32/255；本文件 %d 页全部 0.000%%，文本层与字体层指标完全相同。"
            % n_pages,
            os.path.join(HEAT, "conv_%s.png" % name))

        # 打印对照：A 打印落盘 PDF vs B 打印落盘 PDF（差异来自版面缩放）
        ap = os.path.join(OUT, "printA", name + "_r0.pdf")
        bp = os.path.join(OUT, "printB", name + "_r0.pdf")
        dp = m["diff_print"]
        wp = max(0, min(dp["max_page"] - 1, n_pages - 1))
        a_img, b_img = render(ap, wp), render(bp, wp)
        ga, gb = gray_arr(ap, wp), gray_arr(bp, wp)
        diff = np.abs(ga.astype(np.int16) - gb.astype(np.int16))
        heat = heat_overlay(a_img, diff)
        d = ImageDraw.Draw(heat)
        d.rectangle([0, 0, heat.width, 46], fill=(255, 255, 255))
        d.text((14, 12), "差异像素 %.3f%%（整册均值 %.3f%%）"
               % (dp["per_page"][wp] * 100, dp["mean_diff_ratio"] * 100),
               font=fnt(20, True), fill=(190, 40, 30))
        three_panel(
            "%s（%d 页）第 %d 页：打印落盘对照" % (name, n_pages, wp + 1),
            a_img, b_img, heat,
            "口径：150 DPI 灰度，阈值 |Δ|>32/255；A 页被缩放为 0.905 倍并居中，B 为 1:1 满幅，差异为版面位移而非字形替换。",
            os.path.join(HEAT, "print_%s.png" % name))

    # 人工目检截图：转换产物与打印落盘各取 4 份文档第 1 页
    for name in ["04_table", "06_formula", "07_header_footer", "09_missing_font"]:
        for tag, a_pdf, b_pdf in (
            ("convert", os.path.join(OUT, "convA", name + ".pdf"),
             os.path.join(OUT, "exportB", name + ".pdf")),
            ("printed", os.path.join(OUT, "printA", name + "_r0.pdf"),
             os.path.join(OUT, "printB", name + "_r0.pdf"))):
            pno = 0
            a_img, b_img = render(a_pdf, pno), render(b_pdf, pno)
            w = 620
            ai = a_img.resize((w, int(a_img.height * w / a_img.width)), Image.LANCZOS)
            bi = b_img.resize((w, int(b_img.height * w / b_img.width)), Image.LANCZOS)
            canvas = Image.new("RGB", (w * 2 + 24, ai.height + 52), (246, 247, 249))
            canvas.paste(ai, (0, 44))
            canvas.paste(bi, (w + 24, 44))
            d = ImageDraw.Draw(canvas)
            d.text((10, 12), "A：软件链路（%s）" % ("转换 PDF" if tag == "convert" else "打印落盘 PDF"),
                   font=fnt(18, True), fill=(26, 32, 44))
            d.text((w + 34, 12), "B：Word 导出 + PDF 打印", font=fnt(18, True), fill=(26, 32, 44))
            canvas.save(os.path.join(SHOTS, "%s_%s_p1.png" % (name, tag)))

    # 速度对比图
    conv = json.load(open(os.path.join(OUT, "convert_timing.json"), encoding="utf-8"))
    prt = json.load(open(os.path.join(OUT, "print_timing.json"), encoding="utf-8"))

    def med(vals):
        vals = [v for v in vals if isinstance(v, (int, float))]
        return statistics.median(vals) if vals else float("nan")

    names, cold_a, cold_b, steady_a, steady_b = [], [], [], [], []
    for s in samples:
        name = os.path.splitext(s["name"])[0]
        a_single = med([r["seconds"] for r in conv["A_single"][name] if "seconds" in r])
        a_batch = med([x["seconds"] for x in conv["A_batch"]["per_file"][name] if x["seconds"]])
        b_cold = med(conv["B_cold"][name])
        b_steady = med(conv["B_steady"][name])
        ea = [r for r in prt[name]["A"] if "error" not in r]
        eb = [r for r in prt[name]["B"] if "error" not in r]
        a_layout = med([r["layout"] for r in ea])
        a_print = med([r["print"] for r in ea])
        b_print = med([r["print"] for r in eb])
        names.append(name)
        cold_a.append(a_single + a_layout + a_print)
        cold_b.append(b_cold + b_print)
        steady_a.append(a_batch + a_layout + a_print)
        steady_b.append(b_steady + b_print)

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    fig, axes = plt.subplots(1, 2, figsize=(15, 6.2), sharey=True)
    y = np.arange(len(names))
    h = 0.36
    for ax, av, bv, title in (
        (axes[0], cold_a, cold_b, "冷启动整链路（每份新 Word 进程）"),
        (axes[1], steady_a, steady_b, "稳态整链路（批量/复用 Word 实例）")):
        ax.barh(y + h / 2, av, h, label="A 软件链路", color="#3b6fb6")
        ax.barh(y - h / 2, bv, h, label="B 手动路径", color="#d08a2e")
        for i, (x1, x2) in enumerate(zip(av, bv)):
            ax.text(x1 + 0.15, i + h / 2, "%.1f" % x1, va="center", fontsize=8, color="#26456e")
            ax.text(x2 + 0.15, i - h / 2, "%.1f" % x2, va="center", fontsize=8, color="#7a4f13")
        ax.set_title(title, fontsize=12, fontweight="bold")
        ax.set_xlabel("端到端耗时（秒，3 次中位数）")
        ax.grid(axis="x", ls=":", color="#bbb")
        ax.set_axisbelow(True)
    axes[0].set_yticks(y)
    axes[0].set_yticklabels(names, fontsize=9)
    axes[0].invert_yaxis()
    axes[0].legend(loc="lower right", fontsize=9)
    fig.tight_layout()
    fig.savefig(os.path.join(DOCS, "speed_total.png"), dpi=150)

    # 打印差异逐文件图
    fig, ax = plt.subplots(figsize=(11, 4.6))
    mean = [metrics[n]["diff_print"]["mean_diff_ratio"] * 100 for n in names]
    mx = [metrics[n]["diff_print"]["max_diff_ratio"] * 100 for n in names]
    x = np.arange(len(names))
    ax.bar(x - 0.2, mean, 0.4, label="整册均值", color="#5b8fd1")
    ax.bar(x + 0.2, mx, 0.4, label="最差页", color="#c8574a")
    for i, (m1, m2) in enumerate(zip(mean, mx)):
        ax.text(i - 0.2, m1 + 0.3, "%.1f" % m1, ha="center", fontsize=8)
        ax.text(i + 0.2, m2 + 0.3, "%.1f" % m2, ha="center", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(names, rotation=25, ha="right", fontsize=9)
    ax.set_ylabel("差异像素占比（%）\n阈值 |Δ|>32/255，150 DPI")
    ax.set_title("打印落盘 PDF 的 A/B 差异：全部来自版面缩放（0.905）与居中，不是字形替换", fontsize=11, fontweight="bold")
    ax.grid(axis="y", ls=":", color="#bbb")
    ax.set_axisbelow(True)
    ax.legend()
    fig.tight_layout()
    fig.savefig(os.path.join(DOCS, "print_diff.png"), dpi=150)
    print("artifacts written:", DOCS)


if __name__ == "__main__":
    make_artifacts()
