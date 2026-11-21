# -*- coding: utf-8 -*-
"""修正后的字体/glyph 指标重算（不重跑像素比较，直接更新 metrics.json）。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import measure

OUT = measure.OUT
metrics = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))

for name, m in metrics.items():
    declared = m["declared_fonts"]
    for tag, sub in (("A", "convA"), ("B", "exportB")):
        pdf = os.path.join(OUT, sub, name + ".pdf")
        m[tag]["font"] = measure.font_metrics(pdf, declared)
    a, b = m["A"]["text"], m["B"]["text"]
    m["text_identical"] = all(a[k] == b[k] for k in a)
    print(name, "A_fonts", list(m["A"]["font"]["fonts"]),
          "subst", m["A"]["font"]["substituted"],
          "glyph", m["A"]["font"]["missing_glyph_uses"], "/", m["A"]["font"]["glyph_checks"],
          "| B_subst", m["B"]["font"]["substituted"],
          "B_glyph", m["B"]["font"]["missing_glyph_uses"], "/", m["B"]["font"]["glyph_checks"],
          "| text_identical", m["text_identical"], flush=True)

with open(os.path.join(OUT, "metrics.json"), "w", encoding="utf-8") as fh:
    json.dump(metrics, fh, ensure_ascii=False, indent=2)
