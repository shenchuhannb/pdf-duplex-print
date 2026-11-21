# -*- coding: utf-8 -*-
"""只重跑 B 稳态导出，并合并进 convert_timing.json。"""
import json
import os

import measure

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")


def main():
    samples = measure.sample_list()
    timing_b_steady = measure.convert_b_steady(samples)
    path = os.path.join(OUT, "convert_timing.json")
    data = json.load(open(path, encoding="utf-8"))
    data["B_steady"] = timing_b_steady
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    print("B 稳态重跑完成")


if __name__ == "__main__":
    main()
