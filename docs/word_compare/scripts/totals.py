# -*- coding: utf-8 -*-
"""计算报告要用的逐文件与汇总数字。"""
import json
import os
import statistics

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")
conv = json.load(open(os.path.join(OUT, "convert_timing.json"), encoding="utf-8"))
prt = json.load(open(os.path.join(OUT, "print_timing.json"), encoding="utf-8"))
samples = json.load(open(os.path.join(BASE, "samples", "samples.json"), encoding="utf-8"))


def med(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return statistics.median(vals) if vals else float("nan")


rows = []
for s in samples:
    name = os.path.splitext(s["name"])[0]
    a_single = med([r["seconds"] for r in conv["A_single"][name] if "seconds" in r])
    a_batch = med([x["seconds"] for x in conv["A_batch"]["per_file"][name] if x["seconds"]])
    b_cold = med(conv["B_cold"][name])
    b_steady = med(conv["B_steady"][name])
    ea = [r for r in prt[name]["A"] if "error" not in r]
    eb = [r for r in prt[name]["B"] if "error" not in r]
    a_layout = med([r["layout"] for r in ea])
    a_open = med([r["open"] for r in ea])
    a_print = med([r["print"] for r in ea])
    b_open = med([r["open"] for r in eb])
    b_print = med([r["print"] for r in eb])
    pages = s["pages"]
    pad = 0 if pages % 2 == 0 else 1
    row = dict(name=name, pages=pages, pad=pad, a_single=a_single, a_batch=a_batch,
               b_cold=b_cold, b_steady=b_steady, a_layout=a_layout, a_open=a_open,
               a_print=a_print, b_open=b_open, b_print=b_print)
    row["A_cold_total"] = a_single + a_layout + a_open + a_print
    row["A_steady_total"] = a_batch + a_layout + a_open + a_print
    row["B_cold_total"] = b_cold + b_open + b_print
    row["B_steady_total"] = b_steady + b_open + b_print
    row["ratio_cold"] = row["A_cold_total"] / row["B_cold_total"]
    row["ratio_steady"] = row["A_steady_total"] / row["B_steady_total"]
    rows.append(row)

print(f"{'name':22s} {'p':>3s} {'A单转':>6s} {'A批':>5s} {'B冷':>5s} {'B稳':>5s} "
      f"{'A打印':>6s} {'B打印':>6s} {'A冷总':>6s} {'B冷总':>6s} {'比':>5s} "
      f"{'A稳总':>6s} {'B稳总':>6s} {'比':>5s}")
for r in rows:
    print(f"{r['name']:22s} {r['pages']:3d} {r['a_single']:6.2f} {r['a_batch']:5.2f} "
          f"{r['b_cold']:5.2f} {r['b_steady']:5.2f} {r['a_print']:6.2f} {r['b_print']:6.2f} "
          f"{r['A_cold_total']:6.2f} {r['B_cold_total']:6.2f} {r['ratio_cold']:5.2f} "
          f"{r['A_steady_total']:6.2f} {r['B_steady_total']:6.2f} {r['ratio_steady']:5.2f}")

print()
print("批量总耗时（3 次）:", [round(x["total"], 2) for x in conv["A_batch"]["runs"]])
print("B 稳态导出合计(中位数相加):", round(sum(r["b_steady"] for r in rows), 2))
print("A 批量总链路合计:", round(sum(r["A_steady_total"] for r in rows), 2))
print("B 稳态总链路合计:", round(sum(r["B_steady_total"] for r in rows), 2))
print("A 冷启动总链路合计:", round(sum(r["A_cold_total"] for r in rows), 2))
print("B 冷启动总链路合计:", round(sum(r["B_cold_total"] for r in rows), 2))
print("冷启动逐文件倍数 中位数:", round(statistics.median([r["ratio_cold"] for r in rows]), 2),
      "范围", round(min(r["ratio_cold"] for r in rows), 2), round(max(r["ratio_cold"] for r in rows), 2))
print("稳态逐文件倍数 中位数:", round(statistics.median([r["ratio_steady"] for r in rows]), 2),
      "范围", round(min(r["ratio_steady"] for r in rows), 2), round(max(r["ratio_steady"] for r in rows), 2))
print("转换环节单文件/冷启动 倍数 中位数:",
      round(statistics.median([r["a_single"] / r["b_cold"] for r in rows]), 2),
      "范围", round(min(r["a_single"] / r["b_cold"] for r in rows), 2),
      round(max(r["a_single"] / r["b_cold"] for r in rows), 2))
print("转换环节批量/稳态 倍数 中位数:",
      round(statistics.median([r["a_batch"] / r["b_steady"] for r in rows]), 2),
      "范围", round(min(r["a_batch"] / r["b_steady"] for r in rows), 2),
      round(max(r["a_batch"] / r["b_steady"] for r in rows), 2))
print("打印环节 A/B 倍数 中位数:",
      round(statistics.median([r["a_print"] / r["b_print"] for r in rows]), 2),
      "范围", round(min(r["a_print"] / r["b_print"] for r in rows), 2),
      round(max(r["a_print"] / r["b_print"] for r in rows), 2))
