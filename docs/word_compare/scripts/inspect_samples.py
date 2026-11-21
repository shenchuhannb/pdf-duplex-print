# -*- coding: utf-8 -*-
"""检查样本：页数、字号、字体清单、正文参考文本（增量写盘，避免 Word 崩溃丢结果）。"""
import json
import os
import traceback

import pythoncom
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE, "samples")
NAMES = [
    "01_mixed_short.docx", "02_chinese_pure.docx", "03_english_pure.docx",
    "04_table.docx", "05_image.docx", "06_formula.docx", "07_header_footer.docx",
    "08_legacy.doc", "09_missing_font.docx", "10_long_35p.docx",
]
NOTES = {
    "01_mixed_short.docx": "中英混排短文档（10.5/11/12/16pt）",
    "02_chinese_pure.docx": "纯中文（宋体/黑体/楷体 12pt）",
    "03_english_pure.docx": "纯英文（Calibri/Times New Roman/Arial 10-12pt）",
    "04_table.docx": "表格（12x5、8x4）9pt",
    "05_image.docx": "图片（PNG 12cm/7cm）",
    "06_formula.docx": "公式与特殊符号（Cambria Math 13-14pt）",
    "07_header_footer.docx": "页眉/页脚/页码 5 页 11pt",
    "08_legacy.doc": "旧格式 .doc 2 页",
    "09_missing_font.docx": "缺失字体 NoSuchFontABC 12/14pt",
    "10_long_35p.docx": "35 页长文档 10.5-12pt",
}

WD_STAT_PAGES = 2


def main():
    out_path = os.path.join(SAMPLE_DIR, "samples.json")
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    meta = []
    for name in NAMES:
        path = os.path.join(SAMPLE_DIR, name)
        print("inspect", name, flush=True)
        item = {"name": name, "path": path, "note": NOTES[name]}
        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
            item["pages"] = int(doc.ComputeStatistics(WD_STAT_PAGES))
            fonts = {}
            count = int(doc.Content.Paragraphs.Count)
            step = max(1, count // 60)
            for i in range(1, count + 1, step):
                rng = doc.Content.Paragraphs(i).Range
                fname = str(rng.Font.Name or "")
                fsize = round(float(rng.Font.Size), 1)
                if fname:
                    fonts.setdefault(fname, set()).add(fsize)
            item["fonts"] = {k: sorted(v) for k, v in sorted(fonts.items())}
            text = doc.Content.Text.replace("\x07", "\t")
            ref_path = os.path.join(SAMPLE_DIR, os.path.splitext(name)[0] + ".source.txt")
            with open(ref_path, "w", encoding="utf-8") as handle:
                handle.write(text)
            item["source_text"] = ref_path
            doc.Close(0)
        except Exception as exc:
            traceback.print_exc()
            item["error"] = repr(exc)
        item["size_bytes"] = os.path.getsize(path)
        meta.append(item)
        with open(out_path, "w", encoding="utf-8") as handle:
            json.dump(meta, handle, ensure_ascii=False, indent=2)
    try:
        word.Quit(0)
    except Exception:
        pass
    pythoncom.CoUninitialize()
    print("done", flush=True)


if __name__ == "__main__":
    main()
