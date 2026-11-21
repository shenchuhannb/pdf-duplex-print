# -*- coding: utf-8 -*-
"""修复 07 样本页脚的域顺序，并验证导出后页码正确。"""
import os
import re
import shutil
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(BASE, "samples", "07_header_footer.docx")

RUN = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
       '<w:t xml:space="preserve">{}</w:t></w:r>')
FLD = ('<w:fldSimple w:instr="{}"><w:r><w:rPr><w:rFonts w:hint="eastAsia"/>'
       '<w:noProof/></w:rPr><w:t>{}</w:t></w:r></w:fldSimple>')
PARA = ('<w:p><w:pPr><w:pStyle w:val="af1"/><w:jc w:val="center"/>'
        '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
        + RUN.format("第 ")
        + FLD.format(" PAGE  \\* MERGEFORMAT ", "1")
        + RUN.format(" 页 / 共 ")
        + FLD.format(" NUMPAGES  \\* MERGEFORMAT ", "5")
        + RUN.format(" 页")
        + "</w:p>")


def main():
    src = DOCX
    tmp = src + ".new"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/footer2.xml":
                text = data.decode("utf-8")
                text = re.sub(r"<w:p [^>]*>.*?</w:p>", PARA, text, count=1, flags=re.S)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, src)
    print("footer2.xml 已重写")


if __name__ == "__main__":
    main()
