# -*- coding: utf-8 -*-
"""拆解子进程转换内部耗时（字体预检、Word 启动、导出、退出）。"""
import os
import sys
import time

import pythoncom
import win32com.client

sys.path.insert(0, r"d:\打印机软件")
import word_convert

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "01_mixed_short.docx")
DST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "profile_inproc.pdf")

started = time.perf_counter()
print("converter 探测:", word_convert.available_converters(), "%.2fs" % (time.perf_counter() - started))

started = time.perf_counter()
fonts = word_convert.installed_font_families(refresh=True)
print("字体族枚举:", len(fonts), "%.2fs" % (time.perf_counter() - started))

started = time.perf_counter()
check = word_convert.precheck_docx_fonts(SRC) if hasattr(word_convert, "precheck_docx_fonts") else None
print("docx 字体预检:", check, "%.2fs" % (time.perf_counter() - started))

pythoncom.CoInitialize()
started = time.perf_counter()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
t_start = time.perf_counter() - started
started = time.perf_counter()
doc = app.Documents.Open(SRC, ReadOnly=True, AddToRecentFiles=False, Visible=False)
doc.ExportAsFixedFormat(OutputFileName=DST, ExportFormat=17, OpenAfterExport=False)
doc.Close(0)
t_export = time.perf_counter() - started
started = time.perf_counter()
app.Quit(0)
t_quit = time.perf_counter() - started
pythoncom.CoUninitialize()
print("Word 启动: %.2fs, 打开+导出+关闭: %.2fs, 退出: %.2fs, 合计 %.2fs"
      % (t_start, t_export, t_quit, t_start + t_export + t_quit))
