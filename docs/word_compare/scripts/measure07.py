# -*- coding: utf-8 -*-
"""修复 07 页脚后，只重测 07（A 单文件、B 冷启动、打印、分析），A 批量与 B 稳态整体重跑。"""
import json
import os

import measure

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")


def main():
    all_samples = measure.sample_list()
    one = [s for s in all_samples if measure.safe_name(s["name"]) == "07_header_footer"]
    timing_a = measure.convert_a_single(one)
    timing_batch = measure.convert_a_batch(all_samples)
    timing_b_cold = measure.convert_b_cold(one)
    timing_b_steady = measure.convert_b_steady(all_samples)
    with open(os.path.join(OUT, "convert_timing.json"), "w", encoding="utf-8") as fh:
        json.dump({"A_single": timing_a, "A_batch": timing_batch,
                   "B_cold": timing_b_cold, "B_steady": timing_b_steady},
                  fh, ensure_ascii=False, indent=2)

    timing_print = measure.run_printing(one)
    with open(os.path.join(OUT, "print_timing.json"), "w", encoding="utf-8") as fh:
        json.dump(timing_print, fh, ensure_ascii=False, indent=2)

    metrics = measure.analyze(one)
    path = os.path.join(OUT, "metrics.json")
    old = json.load(open(path, encoding="utf-8"))
    old.update(metrics)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(old, fh, ensure_ascii=False, indent=2)
    print("07 重测完成")


if __name__ == "__main__":
    main()
