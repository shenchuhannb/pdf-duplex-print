# -*- coding: utf-8 -*-
"""把 09_missing_font.docx 的字体名直接写入 OOXML（Word 在输入时会静默替换不存在的字体名，
只有直接改 rFonts 才能模拟“别的机器做好的、本机缺字体”的文档）。"""
import os
import re
import shutil
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(BASE, "samples", "09_missing_font.docx")
TAG = ('<w:rFonts w:ascii="NoSuchFontABC" w:hAnsi="NoSuchFontABC" '
       'w:eastAsia="NoSuchFontABC" w:cs="NoSuchFontABC"/>')


def patch(xml):
    xml = re.sub(r"<w:rFonts[^>]*/>", TAG, xml)

    def run_sub(match):
        body = match.group(1)
        if "<w:rPr>" in body:
            return match.group(0).replace("<w:rPr>", "<w:rPr>" + TAG, 1)
        if "<w:rPr/>" in body:
            return match.group(0).replace("<w:rPr/>", "<w:rPr>" + TAG + "</w:rPr>", 1)
        return match.group(0).replace(body, "<w:rPr>" + TAG + "</w:rPr>" + body, 1)

    return re.sub(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", run_sub, xml, flags=re.S)


def main():
    backup = PATH + ".bak"
    shutil.copy2(PATH, backup)
    with zipfile.ZipFile(backup) as src:
        items = [(info, src.read(info.filename)) for info in src.infolist()]
    with zipfile.ZipFile(PATH, "w", zipfile.ZIP_DEFLATED) as dst:
        for info, data in items:
            if info.filename == "word/document.xml":
                data = patch(data.decode("utf-8")).encode("utf-8")
            dst.writestr(info, data)
    os.remove(backup)
    with zipfile.ZipFile(PATH) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    print("rFonts 次数:", xml.count("NoSuchFontABC"))


if __name__ == "__main__":
    main()
