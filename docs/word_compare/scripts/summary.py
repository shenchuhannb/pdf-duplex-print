# -*- coding: utf-8 -*-
"""把 metrics.json 汇总成便于核对的中文摘要。"""
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")
metrics = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))

print("name | pages | A_chars A_fffd A_missing_rate A_glyph_miss A_subst | B_chars B_fffd B_missing_rate B_glyph_miss B_subst | convdiff printdiff")
for name, m in metrics.items():
    a, b = m["A"], m["B"]
    at, bt = a["text"], b["text"]
    af, bf = a["font"], b["font"]
    print(f"{name} | {m['pages']} | {at['pdf_chars']} {at['ufffd']} {at['missing_rate']} "
          f"{af['missing_glyph_uses']}/{af['glyph_checks']} {af['substituted']} | "
          f"{bt['pdf_chars']} {bt['ufffd']} {bt['missing_rate']} "
          f"{bf['missing_glyph_uses']}/{bf['glyph_checks']} {bf['substituted']} | "
          f"{m['diff_convert']['mean_diff_ratio']} {m['diff_print']['mean_diff_ratio']}")
print()
for name, m in metrics.items():
    print(name, "A fonts", m["A"]["font"]["fonts"], "unembedded", m["A"]["font"]["unembedded"])
    print(name, "B fonts", m["B"]["font"]["fonts"], "unembedded", m["B"]["font"]["unembedded"])
    print(name, "printed A pages/text", m["printA"], "B", m["printB"])
    print(name, "diff_print pages", m["diff_print"]["pages_compared"], m["diff_print"]["per_page"])
    print()
