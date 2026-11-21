# -*- coding: utf-8 -*-
"""重测 B 路径打印（提交打印计时包含落盘等待），更新 print_timing.json 的 B 记录。"""
import json
import os

import measure

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "out")


def main():
    samples = measure.sample_list()
    timing = json.load(open(os.path.join(OUT, "print_timing.json"), encoding="utf-8"))
    for rep in range(measure.REPEATS):
        for item in samples:
            name = measure.safe_name(item["name"])
            exp_pdf = os.path.join(OUT, "exportB", name + ".pdf")
            out_b = os.path.join(OUT, "printB", f"{name}_r{rep}.pdf")
            t_open = measure.open_seconds(exp_pdf)
            t_render = measure.render_seconds(exp_pdf)
            t_print = measure.print_pdf_to_pdf(exp_pdf, out_b, f"B-{name}-r{rep}")
            timing[name]["B"][rep] = {"open": t_open, "render": t_render,
                                      "print": t_print, "pages": item["pages"]}
            print("B print %s rep%d %.2fs" % (name, rep, t_print))
    with open(os.path.join(OUT, "print_timing.json"), "w", encoding="utf-8") as fh:
        json.dump(timing, fh, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
