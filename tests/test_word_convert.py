# -*- coding: utf-8 -*-
"""Word 转换子进程协议回归。"""

import json
import os
from pathlib import Path
import subprocess
import sys


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
