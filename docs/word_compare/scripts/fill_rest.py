# -*- coding: utf-8 -*-
"""补跑其余 9 份样本的 A 单文件转换与 B 冷启动导出，并合并进 convert_timing.json。"""
import json
import os

import measure

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")


def main():
    samples = [s for s in measure.sample_list() if measure.safe_name(s["name"]) != "07_header_footer"]
    a_single = measure.convert_a_single(samples)
    b_cold = measure.convert_b_cold(samples)
    path = os.path.join(OUT, "convert_timing.json")
    data = json.load(open(path, encoding="utf-8"))
    data["A_single"].update(a_single)
    data["B_cold"].update(b_cold)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    print("A 单文件与 B 冷启动补齐完成")


if __name__ == "__main__":
    main()
