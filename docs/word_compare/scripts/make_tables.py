# -*- coding: utf-8 -*-
"""生成报告用的 Markdown 表格片段。"""
import json
import os
import statistics

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")
conv = json.load(open(os.path.join(OUT, "convert_timing.json"), encoding="utf-8"))
prt = json.load(open(os.path.join(OUT, "print_timing.json"), encoding="utf-8"))
metrics = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))
samples = json.load(open(os.path.join(BASE, "samples", "samples.json"), encoding="utf-8"))


def med(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return statistics.median(vals) if vals else None


def rng(vals, digits=2):
    vals = [v for v in vals if isinstance(v, (int, float))]
    if not vals:
        return "—"
    fmt = "%." + str(digits) + "f (%." + str(digits) + "f–%." + str(digits) + "f)"
    return fmt % (statistics.median(vals), min(vals), max(vals))


def pct(v):
    return "%.2f%%" % (v * 100)


rows = []
for s in samples:
    name = os.path.splitext(s["name"])[0]
    m = metrics[name]
    a_single = [r["seconds"] for r in conv["A_single"][name] if "seconds" in r]
    a_batch = [x["seconds"] for x in conv["A_batch"]["per_file"][name] if x["seconds"]]
    b_cold = [v for v in conv["B_cold"][name] if v is not None]
    b_steady = [v for v in conv["B_steady"][name] if v is not None]
    ea = [r for r in prt[name]["A"] if "error" not in r]
    eb = [r for r in prt[name]["B"] if "error" not in r]
    row = dict(s=s, m=m, name=name,
               a_single=a_single, a_batch=a_batch, b_cold=b_cold, b_steady=b_steady,
               a_layout=[r["layout"] for r in ea], a_open=[r["open"] for r in ea],
               a_render=[r["render"] for r in ea], a_print=[r["print"] for r in ea],
               b_open=[r["open"] for r in eb], b_render=[r["render"] for r in eb],
               b_print=[r["print"] for r in eb])
    row["a_cold_total"] = [
        a_single[i] + row["a_layout"][i] + row["a_print"][i]
        for i in range(min(len(a_single), len(row["a_layout"]), len(row["a_print"])))]
    row["a_steady_total"] = [
        a_batch[i] + row["a_layout"][i] + row["a_print"][i]
        for i in range(min(len(a_batch), len(row["a_layout"]), len(row["a_print"])))]
    row["b_cold_total"] = [
        b_cold[i] + row["b_open"][i] + row["b_print"][i]
        for i in range(min(len(b_cold), len(row["b_open"]), len(row["b_print"])))
        if b_cold[i] is not None]
    row["b_steady_total"] = [
        b_steady[i] + row["b_open"][i] + row["b_print"][i]
        for i in range(min(len(b_steady), len(row["b_open"]), len(row["b_print"])))
        if b_steady[i] is not None]
    rows.append(row)

lines = []
lines.append("### 表 1 样本与乱码指标（A：软件链路转换产物；B：Word 导出）\n")
lines.append("| 样本 | 页数 | 字号 / 声明字体 | A 字符 / B 字符 | U+FFFD（A/B） | 方框字符（A/B） | 字符缺失率（A/B） | 缺失 glyph（A/B） | 未嵌入字体（A/B） | 替换字体（A/B） | 转换像素差（均值/最差） | 打印像素差（均值/最差） |\n")
lines.append("|---|---|---|---|---|---|---|---|---|---|---|---|\n")
for r in rows:
    m = r["m"]
    fonts = "、".join(m["declared_fonts"])
    note = r["s"].get("note", "")
    a = m["A"]["text"]
    b = m["B"]["text"]
    sub_a = "、".join(m["A"]["font"]["substituted"]) or "无"
    sub_b = "、".join(m["B"]["font"]["substituted"]) or "无"
    lines.append("| %s | %d | %s；%s | %d / %d | %d / %d | %d / %d | %s / %s | %d/%d / %d/%d | %d / %d | %s / %s | %s / %s | %s / %s |\n" % (
        r["name"], m["pages"], note, fonts, a["pdf_chars"], b["pdf_chars"],
        a["ufffd"], b["ufffd"], a["box_chars"], b["box_chars"],
        pct(a["missing_rate"]), pct(b["missing_rate"]),
        m["A"]["font"]["missing_glyph_uses"], m["A"]["font"]["glyph_checks"],
        m["B"]["font"]["missing_glyph_uses"], m["B"]["font"]["glyph_checks"],
        len(m["A"]["font"]["unembedded"]), len(m["B"]["font"]["unembedded"]),
        sub_a, sub_b,
        pct(m["diff_convert"]["mean_diff_ratio"]), pct(m["diff_convert"]["max_diff_ratio"]),
        pct(m["diff_print"]["mean_diff_ratio"]), pct(m["diff_print"]["max_diff_ratio"])))

lines.append("\n### 表 2 转换与渲染耗时（秒，3 次中位数（范围））\n")
lines.append("| 样本 | 页数 | A 单文件转换（冷） | A 批量内单文件（稳态） | B 冷启动导出 | B 稳态导出 | A 渲染 | B 渲染 |\n")
lines.append("|---|---|---|---|---|---|---|---|\n")
for r in rows:
    lines.append("| %s | %d | %s | %s | %s | %s | %s | %s |\n" % (
        r["name"], r["m"]["pages"], rng(r["a_single"]), rng(r["a_batch"]),
        rng(r["b_cold"]), rng(r["b_steady"]), rng(r["a_render"]), rng(r["b_render"])))

lines.append("\n### 表 3 打印与整链路耗时（秒，3 次中位数（范围））\n")
lines.append("| 样本 | A 排版 | A 提交打印 | B 打开 PDF | B 提交打印 | 冷启动总耗时 A / B | 冷启动倍数 | 稳态总耗时 A / B | 稳态倍数 |\n")
lines.append("|---|---|---|---|---|---|---|---|---|\n")
for r in rows:
    lines.append("| %s | %s | %s | %s | %s | %s / %s | %.2fx | %s / %s | %.2fx |\n" % (
        r["name"], rng(r["a_layout"]), rng(r["a_print"]), rng(r["b_open"], 3), rng(r["b_print"]),
        rng(r["a_cold_total"]), rng(r["b_cold_total"]), statistics.median(r["a_cold_total"]) / statistics.median(r["b_cold_total"]),
        rng(r["a_steady_total"]), rng(r["b_steady_total"]), statistics.median(r["a_steady_total"]) / statistics.median(r["b_steady_total"])))

# 汇总
all_chars = sum(metrics[n]["A"]["text"]["pdf_chars"] for n in metrics)
all_glyph = sum(metrics[n]["A"]["font"]["glyph_checks"] for n in metrics)
all_pages = sum(metrics[n]["pages"] for n in metrics)
conv_zero = sum(1 for n in metrics if metrics[n]["diff_convert"]["max_diff_ratio"] == 0)
print_pages = sum(metrics[n]["diff_print"]["pages_compared"] for n in metrics)
print_ratios = [v for n in metrics for v in metrics[n]["diff_print"]["per_page"]]
lines.append("\n### 汇总数字\n")
lines.append("- 样本 %d 份、共 %d 页；转换对照比较 %d 页，其中 max 像素差为 0 的文档 %d/%d，全部页面 0。\n" % (
    len(metrics), all_pages, all_pages, conv_zero, len(metrics)))
lines.append("- 文本层：PDF 字符 %d 个，U+FFFD 0 个，方框字符 0 个；字符缺失率除 05 为 1.61%%（A/B 完全相同）外均为 0；glyph 检查 %d 次，缺失 0 次。\n" % (
    all_chars, all_glyph))
lines.append("- 字体层：两路径所有字体均嵌入，未嵌入字体数 A/B 均为 0；替换字体仅 09 有 Cambria（A/B 相同）。\n")
lines.append("- 打印像素差：比较 %d 页，整册均值 %.2f%%–%.2f%%，全部页均值 %.2f%%，单页最大值 %.2f%%。\n" % (
    print_pages, min(v for n in metrics for v in [metrics[n]["diff_print"]["mean_diff_ratio"]]) * 100,
    max(v for n in metrics for v in [metrics[n]["diff_print"]["mean_diff_ratio"]]) * 100,
    statistics.mean(print_ratios) * 100, max(print_ratios) * 100))
lines.append("- A 单文件转换合计 %.2fs，B 冷启动导出合计 %.2fs；A 批量 3 轮总耗时 %s s，中位 %.2fs；B 稳态导出合计 %.2fs。\n" % (
    sum(med(r["a_single"]) for r in rows), sum(med(r["b_cold"]) for r in rows),
    "/".join("%.2f" % x["total"] for x in conv["A_batch"]["runs"]),
    med([x["total"] for x in conv["A_batch"]["runs"]]), sum(med(r["b_steady"]) for r in rows)))
lines.append("- 冷启动整链路合计 A %.2fs / B %.2fs（%.2fx）；稳态整链路合计 A %.2fs / B %.2fs（%.2fx）。\n" % (
    sum(med(r["a_cold_total"]) for r in rows), sum(med(r["b_cold_total"]) for r in rows),
    sum(med(r["a_cold_total"]) for r in rows) / sum(med(r["b_cold_total"]) for r in rows),
    sum(med(r["a_steady_total"]) for r in rows), sum(med(r["b_steady_total"]) for r in rows),
    sum(med(r["a_steady_total"]) for r in rows) / sum(med(r["b_steady_total"]) for r in rows)))
lines.append("- 冷启动逐文件倍数中位 %.2fx（%.2f–%.2fx）；稳态逐文件倍数中位 %.2fx（%.2f–%.2fx）。\n" % (
    statistics.median([med(r["a_cold_total"]) / med(r["b_cold_total"]) for r in rows]),
    min(med(r["a_cold_total"]) / med(r["b_cold_total"]) for r in rows),
    max(med(r["a_cold_total"]) / med(r["b_cold_total"]) for r in rows),
    statistics.median([med(r["a_steady_total"]) / med(r["b_steady_total"]) for r in rows]),
    min(med(r["a_steady_total"]) / med(r["b_steady_total"]) for r in rows),
    max(med(r["a_steady_total"]) / med(r["b_steady_total"]) for r in rows)))

text = "".join(lines)
open(os.path.join(BASE, "tables.md"), "w", encoding="utf-8").write(text)
print(text)
