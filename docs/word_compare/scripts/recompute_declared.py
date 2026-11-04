# -*- coding: utf-8 -*-
"""用 OOXML 字体表（.docx）与 Word 逐字符采样（.doc）重建“文档声明的字体清单”，
再重算字体替换指标。"""
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import measure

import pythoncom
import win32com.client

SAMPLE_DIR = os.path.join(measure.BASE, "samples")
OUT = measure.OUT


def fonts_from_docx(path):
    names = set()
    with zipfile.ZipFile(path) as z:
        parts = [n for n in z.namelist()
                 if re.match(r"word/(document|styles|header\d*|footer\d*)\.xml$", n)]
        for part in parts:
            xml = z.read(part).decode("utf-8")
            for tag in re.findall(r"<w:rFonts[^>]*>", xml):
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    names |= set(re.findall(r'w:%s="([^"]+)"' % attr, tag))
    names.discard("")
    return names


def fonts_from_doc(path):
    names = set()
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
        total = int(doc.Characters.Count)
        step = max(1, total // 200)
        for i in range(1, total + 1, step):
            name = str(doc.Characters(i).Font.Name or "")
            if name:
                names.add(name)
        doc.Close(0)
    finally:
        try:
            app.Quit(0)
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return names


def main():
    samples_path = os.path.join(SAMPLE_DIR, "samples.json")
    samples = json.load(open(samples_path, encoding="utf-8"))
    for item in samples:
        path = item["path"]
        try:
            if path.lower().endswith(".docx"):
                declared = fonts_from_docx(path)
            else:
                declared = fonts_from_doc(path)
        except Exception as exc:
            print("字体清点失败", item["name"], exc)
            declared = set(item["fonts"])
        item["declared_fonts"] = sorted(declared)
        print(item["name"], "declared", item["declared_fonts"], flush=True)
    json.dump(samples, open(samples_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    metrics = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))
    by_name = {os.path.splitext(item["name"])[0]: item for item in samples}
    for name, m in metrics.items():
        declared = by_name[name]["declared_fonts"]
        m["declared_fonts"] = declared
        for tag, sub in (("A", "convA"), ("B", "exportB")):
            m[tag]["font"] = measure.font_metrics(os.path.join(OUT, sub, name + ".pdf"), declared)
        print(name, "A_subst", m["A"]["font"]["substituted"],
              "B_subst", m["B"]["font"]["substituted"], flush=True)
    json.dump(metrics, open(os.path.join(OUT, "metrics.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
