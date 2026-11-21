# -*- coding: utf-8 -*-
"""汇总速度数据，打印关键统计，供报告撰写。"""
import json
import os
import statistics

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
conv = json.load(open(os.path.join(OUT, "convert_timing.json"), encoding="utf-8"))
prt = json.load(open(os.path.join(OUT, "print_timing.json"), encoding="utf-8"))
samples = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      "samples", "samples.json"), encoding="utf-8"))


def med(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return round(statistics.median(vals), 2) if vals else None


def rng(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return (round(min(vals), 2), round(max(vals), 2)) if vals else (None, None)


print("=== A single convert (cold, 新 Word 进程) ===")
for s in samples:
    name = os.path.splitext(s["name"])[0]
    runs = conv["A_single"].get(name, [])
    vals = [r["seconds"] for r in runs if "seconds" in r]
    errs = sum(1 for r in runs if "error" in r)
    print(name, "pages", s["pages"], "median", med(vals), "range", rng(vals), "err", errs)

print("=== A batch ===")
for run in conv["A_batch"]["runs"]:
    print(" rep", run["rep"], "total", round(run["total"], 2), "ok", run["ok"])
print(" batch per-file:")
for s in samples:
    name = os.path.splitext(s["name"])[0]
    vals = [x["seconds"] for x in conv["A_batch"]["per_file"].get(name, []) if x["seconds"]]
    print("  ", name, "median", med(vals), "range", rng(vals))

print("=== B cold / steady ===")
for s in samples:
    name = os.path.splitext(s["name"])[0]
    for key in ("B_cold", "B_steady"):
        vals = conv[key].get(name, [])
        print(name, key, "median", med(vals), "range", rng(vals), "vals", [round(v, 2) if v else v for v in vals])

print("=== printing ===")
for s in samples:
    name = os.path.splitext(s["name"])[0]
    entry = prt.get(name, {})
    for tag in ("A", "B"):
        runs = [r for r in entry.get(tag, []) if "error" not in r]
        if not runs:
            print(name, tag, "NO RUN")
            continue
        parts = {}
        for key in ("layout", "open", "render", "print"):
            if key in runs[0]:
                vals = [r[key] for r in runs]
                parts[key] = (med(vals), rng(vals))
        print(name, tag, "pages", runs[0]["pages"], parts)
