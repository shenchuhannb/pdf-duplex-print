# -*- coding: utf-8 -*-
"""拆解 A 路径单文件转换 7～9 秒的构成（一次抽样，仅用于解释耗时）。"""
import os
import subprocess
import sys
import time

sys.path.insert(0, r"d:\打印机软件")
import word_convert

PROD_PY = r"D:\\打印机软件\.venv\Scripts\python.exe"
SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "01_mixed_short.docx")

started = time.perf_counter()
before = word_convert.conversion_processes()
t_enum = time.perf_counter() - started

started = time.perf_counter()
word_convert.conversion_processes()
after = word_convert.conversion_processes()
t_enum2 = time.perf_counter() - started
started = time.perf_counter()
word_convert._terminate_new_office(after)
t_term = time.perf_counter() - started

started = time.perf_counter()
subprocess.run([PROD_PY, "-c", "import sys; sys.path.insert(0, r'd:\\打印机软件'); import word_convert"],
               capture_output=True)
t_import = time.perf_counter() - started

cmd = [PROD_PY, os.path.abspath(word_convert.__file__), "--src", SRC,
       "--dst", os.path.join(os.path.dirname(os.path.abspath(__file__)), "profile_out.pdf")]
started = time.perf_counter()
proc = subprocess.run(cmd, capture_output=True)
t_subproc = time.perf_counter() - started

print("conversion_processes 枚举: %.2fs (第二次 %.2fs, 清理 %.2fs)" % (t_enum, t_enum2, t_term))
print("子进程启动 + import word_convert: %.2fs" % t_import)
print("子进程完整转换（CLI 到退出）: %.2fs, rc=%s" % (t_subproc, proc.returncode))
print("父进程 convert_file 额外开销估计: %.2fs" % (7.48 - t_subproc))
