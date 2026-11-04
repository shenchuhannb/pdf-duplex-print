# -*- coding: utf-8 -*-
"""补齐 B 路径缺失的导出文件，并比较 A/B 转换产物是否逐字节相同。"""
import hashlib
import json
import os

import pythoncom
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(BASE, "samples")
OUT = os.path.join(BASE, "out")


def md5(path):
    with open(path, "rb") as fh:
        return hashlib.md5(fh.read()).hexdigest()


def main():
    samples = json.load(open(os.path.join(SAMPLES, "samples.json"), encoding="utf-8"))
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for item in samples:
            name = os.path.splitext(item["name"])[0]
            dst = os.path.join(OUT, "exportB", name + ".pdf")
            if os.path.exists(dst):
                continue
            doc = app.Documents.Open(item["path"], ReadOnly=True, AddToRecentFiles=False, Visible=False)
            doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17, OpenAfterExport=False)
            doc.Close(0)
            print("补齐", dst, flush=True)
    finally:
        app.Quit(0)
    pythoncom.CoUninitialize()

    same = 0
    for item in samples:
        name = os.path.splitext(item["name"])[0]
        a = os.path.join(OUT, "convA", name + ".pdf")
        b = os.path.join(OUT, "exportB", name + ".pdf")
        equal = os.path.exists(a) and os.path.exists(b) and md5(a) == md5(b)
        same += 1 if equal else 0
        print(f"{name}: md5_equal={equal} sizeA={os.path.getsize(a)} sizeB={os.path.getsize(b)}", flush=True)
    print("逐字节相同:", same, "/", len(samples))


if __name__ == "__main__":
    main()
