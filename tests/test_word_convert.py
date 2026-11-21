# -*- coding: utf-8 -*-
"""Word 转换子进程协议回归。"""

import json
import os
from pathlib import Path
import subprocess
import sys
import zipfile

import word_convert


def test_main_reports_chinese_font_names_as_utf8():
    script = """
import word_convert
word_convert.inspect_document_fonts = lambda path: {
    "checked": True, "missing": ["方正小标宋简体"], "fonts": ["方正小标宋简体"],
}
word_convert.convert_inprocess = lambda src, dst, font_info=None: "test"
raise SystemExit(word_convert._main(["--src", "sample.docx", "--dst", "out.pdf"]))
"""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "gbk"
    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=Path(__file__).resolve().parents[1], env=env,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30, check=False,
    )

    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    lines = result.stdout.decode("utf-8").splitlines()
    values = dict(line.split("=", 1) for line in lines)
    assert values["converter"] == "test"
    assert json.loads(values["font_info"])["missing"] == ["方正小标宋简体"]


def test_docx_ignores_unused_font_table_entries(tmp_path):
    path = tmp_path / "fonts.docx"
    namespace = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "word/document.xml",
            f'<w:document {namespace}><w:body><w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="黑体"/>'
            '</w:rPr><w:t>正文</w:t></w:r></w:p></w:body></w:document>',
        )
        archive.writestr(
            "word/styles.xml",
            f'<w:styles {namespace}><w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
        )
        archive.writestr(
            "word/fontTable.xml",
            f'<w:fontTable {namespace}><w:font w:name="MS Mincho"/>'
            '<w:font w:name="方正小标宋简体"/><w:font w:name="黑体"/></w:fontTable>',
        )

    result = word_convert.inspect_document_fonts(
        str(path), installed={"Calibri", "黑体"})

    assert result["checked"] is True
    assert result["missing"] == ["方正小标宋简体"]
    assert "MS Mincho" not in result["fonts"]

    result = word_convert.inspect_document_fonts(
        str(path), installed={"Calibri", "黑体", "方正小标宋简体"})
    assert result["missing"] == []
