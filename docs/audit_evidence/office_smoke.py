# -*- coding: utf-8 -*-
"""可选的真实 Word 隐藏转换冒烟：仅操作临时文件，不执行打印或全局进程清理。"""
import argparse
import json
import pathlib
import sys
import tempfile
import zipfile

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

import fitz
from word_convert import convert_file


def make_docx(path: pathlib.Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>CODEARTS_OFFICE_SMOKE_审计转换</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document)


parser = argparse.ArgumentParser()
parser.add_argument("--python-exe", default="", help="可选：冻结 exe 或其他 worker 解释器")
parser.add_argument("--batch", action="store_true", help="验证两个文件的批量转换")
args = parser.parse_args()
with tempfile.TemporaryDirectory(prefix="office-smoke-") as temp:
    root = pathlib.Path(temp)
    if args.batch:
        from word_convert import convert_batch

        items = []
        for index in (1, 2):
            src = root / f"smoke-{index}.docx"
            make_docx(src)
            items.append({"iid": str(index), "src": str(src),
                          "dst": str(root / f"smoke-{index}.pdf"), "kind": "word"})
        callbacks = []
        results = convert_batch(items, timeout_per_file=90, on_result=callbacks.append,
                                python_exe=args.python_exe or None)
        print(json.dumps({"results": results, "callback_count": len(callbacks),
                          "python_exe": args.python_exe or "source"},
                         ensure_ascii=False, indent=2))
        raise SystemExit(0)
    src = root / "smoke.docx"
    dst = root / "smoke.pdf"
    make_docx(src)
    info = convert_file(str(src), str(dst), timeout=90,
                        python_exe=args.python_exe or None)
    with fitz.open(dst) as doc:
        info["text_preserved"] = "CODEARTS_OFFICE_SMOKE" in doc[0].get_text()
        info["pdf_pages"] = doc.page_count
    info["python_exe"] = args.python_exe or "source"
    print(json.dumps(info, ensure_ascii=False, indent=2))
