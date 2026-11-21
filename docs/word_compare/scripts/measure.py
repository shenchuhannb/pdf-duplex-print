# -*- coding: utf-8 -*-
"""Word 打印两条路径对比实测：转换 / 排版 / 打印 / 乱码指标。

路径 A（软件路径）：word_convert.convert_file/convert_batch → pdf_layout.compose_layout_pdf
                     → printer_core.GdiPrinter → Microsoft Print to PDF 落盘
路径 B（手动路径）：Word COM ExportAsFixedFormat（默认参数）→ 独立 GDI 打印脚本
                     （300 DPI 栅格 + StretchDIBits）→ 同一虚拟打印机落盘

不改动业务代码；对 A 的唯一注入是通过 DOCINFO.lpszOutput 指定输出文件
（产品代码原本不设置该字段，虚拟打印机需要它才能自动落盘）。
"""
import ctypes
import difflib
import json
import os
import re
import shutil
import statistics
import sys
import threading
import time
import traceback
from collections import Counter
from ctypes import wintypes

import fitz
import numpy as np
import pythoncom
import win32com.client

PROJECT = r"d:\打印机软件"
sys.path.insert(0, PROJECT)

import pdf_layout          # noqa: E402
import printer_core        # noqa: E402
import word_convert        # noqa: E402

PRINTER = "Microsoft Print to PDF"
DPI = 300
REPEATS = 3
BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE, "samples")
OUT = os.path.join(BASE, "out")
PROD_PY = r"D:\\打印机软件\.venv\Scripts\python.exe"
if not os.path.isfile(PROD_PY):
    PROD_PY = sys.executable

for sub in ("convA", "exportB", "layoutA", "printA", "printB",
            "heat_conv", "heat_print", "shots"):
    os.makedirs(os.path.join(OUT, sub), exist_ok=True)


def log(msg):
    print(time.strftime("%H:%M:%S"), msg, flush=True)


def sample_list():
    with open(os.path.join(SAMPLE_DIR, "samples.json"), encoding="utf-8") as fh:
        return json.load(fh)


def safe_name(name):
    return os.path.splitext(name)[0]


# ---------------------------------------------------------------- A 路径打印注入
_current_output = {"path": None}
_orig_startdoc = printer_core._gdi32.StartDocW


def _startdoc_with_output(hdc, ptr):
    info = ctypes.cast(ptr, ctypes.POINTER(printer_core._DOCINFOW)).contents
    info.lpszOutput = _current_output["path"]
    return _orig_startdoc(hdc, ctypes.byref(info))


printer_core._gdi32.StartDocW = _startdoc_with_output


def wait_file(path, timeout=120.0):
    """等待虚拟打印机写完文件（大小稳定 1 秒）。"""
    deadline = time.time() + timeout
    last = -1
    stable = 0
    while time.time() < deadline:
        if os.path.exists(path):
            size = os.path.getsize(path)
            if size > 0 and size == last:
                stable += 1
                if stable >= 2:
                    return size
            else:
                stable = 0
            last = size
        time.sleep(0.5)
    return os.path.getsize(path) if os.path.exists(path) else 0


# ---------------------------------------------------------------- B 路径独立 GDI 打印
class _DOCINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_int),
        ("lpszDocName", wintypes.LPCWSTR),
        ("lpszOutput", wintypes.LPCWSTR),
        ("lpszDatatype", wintypes.LPCWSTR),
        ("fwType", wintypes.DWORD),
    ]


class _BMIH(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
gdi32.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
gdi32.CreateDCW.restype = wintypes.HDC
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
gdi32.GetDeviceCaps.restype = ctypes.c_int
gdi32.StartDocW.argtypes = [wintypes.HDC, ctypes.POINTER(_DOCINFOW)]
gdi32.StartDocW.restype = ctypes.c_int
gdi32.StartPage.argtypes = [wintypes.HDC]
gdi32.EndPage.argtypes = [wintypes.HDC]
gdi32.EndDoc.argtypes = [wintypes.HDC]
gdi32.AbortDoc.argtypes = [wintypes.HDC]
gdi32.SetStretchBltMode.argtypes = [wintypes.HDC, ctypes.c_int]
gdi32.StretchDIBits.argtypes = [
    wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_void_p, ctypes.POINTER(_BMIH), wintypes.DWORD, wintypes.DWORD,
]
gdi32.StretchDIBits.restype = ctypes.c_int


def _fit_rect(src_w, src_h, dst_w, dst_h):
    scale = min(dst_w / src_w, dst_h / src_h)
    width, height = max(1, int(src_w * scale)), max(1, int(src_h * scale))
    return (max(0, (dst_w - width) // 2), max(0, (dst_h - height) // 2), width, height)


def print_pdf_to_pdf(pdf_path, out_path, job_name, dpi=DPI, split=False):
    """独立 GDI 打印（模拟 PDF 阅读器“以图像方式打印”），输出到虚拟打印机文件。"""
    if os.path.exists(out_path):
        os.remove(out_path)
    doc = fitz.open(pdf_path)
    started = time.perf_counter()
    hdc = gdi32.CreateDCW("WINSPOOL", PRINTER, None, None)
    if not hdc:
        doc.close()
        raise RuntimeError("CreateDC 失败")
    try:
        hdc = hdc  # type: ignore
        printable_w = gdi32.GetDeviceCaps(hdc, 8)
        printable_h = gdi32.GetDeviceCaps(hdc, 10)
        info = _DOCINFOW(ctypes.sizeof(_DOCINFOW), job_name, out_path, None, 0)
        if gdi32.StartDocW(hdc, ctypes.byref(info)) <= 0:
            raise RuntimeError("StartDoc 失败")
        for page in doc:
            pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
            width, height = pix.width, pix.height
            stride = (width * 3 + 3) & ~3
            buffer = ctypes.create_string_buffer(stride * height)
            src_ref = ctypes.c_char_p(pix.samples)
            src = ctypes.cast(src_ref, ctypes.c_void_p).value
            for row in range(height):
                ctypes.memmove(ctypes.addressof(buffer) + row * stride,
                               src + row * pix.stride, pix.stride)
            bmi = _BMIH(ctypes.sizeof(_BMIH), width, -height, 1, 24, 0,
                        stride * height, 0, 0, 0, 0)
            dest = _fit_rect(width, height, printable_w, printable_h)
            if gdi32.StartPage(hdc) <= 0:
                raise RuntimeError("StartPage 失败")
            gdi32.StretchDIBits(hdc, dest[0], dest[1], dest[2], dest[3],
                                0, 0, width, height, buffer, ctypes.byref(bmi), 0, 0x00CC0020)
            gdi32.EndPage(hdc)
        gdi32.EndDoc(hdc)
    finally:
        gdi32.DeleteDC(hdc)
        doc.close()
    submit = time.perf_counter() - started
    flush_started = time.perf_counter()
    wait_file(out_path, timeout=120)
    flush = time.perf_counter() - flush_started
    return (submit, flush) if split else submit + flush


# ---------------------------------------------------------------- 转换
def convert_a_single(samples):
    """A 路径单文件转换（每次新起 Word 进程，等价冷启动）。"""
    out = {}
    job = os.path.join(OUT, "convA")
    for rep in range(REPEATS):
        for item in samples:
            name = safe_name(item["name"])
            dst = os.path.join(job, f"{name}.pdf") if rep == 0 else os.path.join(job, f"{name}_r{rep}.pdf")
            started = time.perf_counter()
            try:
                info = word_convert.convert_file(item["path"], dst, python_exe=PROD_PY)
                entry = {"seconds": time.perf_counter() - started, "info": info}
            except Exception as exc:
                entry = {"seconds": time.perf_counter() - started, "error": repr(exc)}
            out.setdefault(name, []).append(entry)
            log(f"A single convert {name} rep{rep} {entry['seconds']:.2f}s")
    return out


def convert_a_batch(samples):
    runs = []
    per_file = {}
    for rep in range(REPEATS):
        batch_dir = os.path.join(OUT, "convA", f"batch_r{rep}")
        os.makedirs(batch_dir, exist_ok=True)
        items = [{"iid": safe_name(s["name"]), "src": s["path"],
                  "dst": os.path.join(batch_dir, safe_name(s["name"]) + ".pdf"),
                  "kind": "word"} for s in samples]
        started = time.perf_counter()
        results = word_convert.convert_batch(items, python_exe=PROD_PY)
        total = time.perf_counter() - started
        runs.append({"rep": rep, "total": total, "ok": sum(1 for r in results.values() if r.get("ok"))})
        for iid, res in results.items():
            per_file.setdefault(iid, []).append(
                {"seconds": res.get("seconds"), "ok": res.get("ok"), "pages": res.get("pages")})
        log(f"A batch rep{rep} {total:.2f}s ok={runs[-1]['ok']}/{len(items)}")
    return {"runs": runs, "per_file": per_file}


def word_app():
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def export_b(app, src, dst):
    started = time.perf_counter()
    doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        if os.path.exists(dst):
            os.remove(dst)
        doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17, OpenAfterExport=False)
    finally:
        doc.Close(0)
    return time.perf_counter() - started


def convert_b_cold(samples):
    """B 路径冷启动：每次新建 Word 实例并导出第一份文档。"""
    out = {}
    job = os.path.join(OUT, "exportB", "cold")
    os.makedirs(job, exist_ok=True)
    for rep in range(REPEATS):
        for item in samples:
            name = safe_name(item["name"])
            dst = os.path.join(job, f"{name}_r{rep}.pdf")
            pythoncom.CoInitialize()
            app = word_app()
            try:
                seconds = export_b(app, item["path"], dst)
            except Exception as exc:
                seconds = None
                log(f"B cold export 失败 {name}: {exc}")
            finally:
                try:
                    app.Quit(0)
                except Exception:
                    pass
                pythoncom.CoUninitialize()
            out.setdefault(name, []).append(seconds)
            log(f"B cold export {name} rep{rep} {seconds and round(seconds, 2)}s")
    return out


def convert_b_steady(samples):
    """B 路径稳态：同一 Word 实例连续导出。"""
    out = {}
    job = os.path.join(OUT, "exportB")
    os.makedirs(job, exist_ok=True)
    pythoncom.CoInitialize()
    app = word_app()
    try:
        for rep in range(REPEATS):
            for item in samples:
                name = safe_name(item["name"])
                dst = os.path.join(job, f"{name}.pdf") if rep == 0 else os.path.join(job, f"{name}_r{rep}.pdf")
                try:
                    seconds = export_b(app, item["path"], dst)
                except Exception as exc:
                    seconds = None
                    log(f"B steady export 失败 {name}: {exc}")
                out.setdefault(name, []).append(seconds)
                log(f"B steady export {name} rep{rep} {seconds and round(seconds, 2)}s")
    finally:
        try:
            app.Quit(0)
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return out


# ---------------------------------------------------------------- 排版与打印
def run_layout(name, conv_pdf, out_pdf):
    doc = fitz.open(conv_pdf)
    pages = list(range(1, doc.page_count + 1))
    doc.close()
    started = time.perf_counter()
    pdf_layout.compose_layout_pdf(conv_pdf, pages, out_pdf, pdf_layout.LayoutOptions())
    return time.perf_counter() - started, len(pages)


def print_a(layout_pdf, out_pdf, job_name, split=False):
    if os.path.exists(out_pdf):
        os.remove(out_pdf)
    _current_output["path"] = out_pdf
    printer = printer_core.GdiPrinter(PRINTER, printer_core.PrinterOptions(dpi=DPI))
    doc = fitz.open(layout_pdf)
    pages = list(range(1, doc.page_count + 1))
    doc.close()
    started = time.perf_counter()
    printer.print_pages(layout_pdf, pages, job_name, threading.Event())
    submit = time.perf_counter() - started
    flush_started = time.perf_counter()
    wait_file(out_pdf)
    flush = time.perf_counter() - flush_started
    return (submit, flush) if split else submit + flush


def render_seconds(pdf_path, dpi=DPI):
    started = time.perf_counter()
    doc = fitz.open(pdf_path)
    try:
        for page in doc:
            page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
    finally:
        doc.close()
    return time.perf_counter() - started


def open_seconds(pdf_path):
    """脚本等价的“打开 PDF”成本：解析文件并读取页数，不含阅读器 GUI 启动。"""
    started = time.perf_counter()
    doc = fitz.open(pdf_path)
    try:
        doc.page_count
    finally:
        doc.close()
    return time.perf_counter() - started


def run_printing(samples):
    timing = {}
    for rep in range(REPEATS):
        for item in samples:
            name = safe_name(item["name"])
            pages = item["pages"]
            conv_pdf = os.path.join(OUT, "convA", name + ".pdf")
            exp_pdf = os.path.join(OUT, "exportB", name + ".pdf")
            layout_pdf = os.path.join(OUT, "layoutA", f"{name}_r{rep}.pdf")
            out_a = os.path.join(OUT, "printA", f"{name}_r{rep}.pdf")
            out_b = os.path.join(OUT, "printB", f"{name}_r{rep}.pdf")
            entry = timing.setdefault(name, {"A": [], "B": []})
            # A
            try:
                t_layout, _ = run_layout(name, conv_pdf, layout_pdf)
                t_open = open_seconds(layout_pdf)
                t_render = render_seconds(layout_pdf)
                t_submit, t_flush = print_a(layout_pdf, out_a, f"A-{name}-r{rep}", split=True)
                entry["A"].append({"layout": t_layout, "render": t_render,
                                   "open": t_open, "submit": t_submit, "flush": t_flush,
                                   "print": t_submit + t_flush, "pages": pages})
            except Exception as exc:
                traceback.print_exc()
                entry["A"].append({"error": repr(exc)})
            # B
            try:
                t_open = open_seconds(exp_pdf)
                t_render = render_seconds(exp_pdf)
                t_submit, t_flush = print_pdf_to_pdf(exp_pdf, out_b, f"B-{name}-r{rep}", split=True)
                entry["B"].append({"open": t_open, "render": t_render,
                                   "submit": t_submit, "flush": t_flush,
                                   "print": t_submit + t_flush, "pages": pages})
            except Exception as exc:
                traceback.print_exc()
                entry["B"].append({"error": repr(exc)})
            log(f"print {name} rep{rep} A={entry['A'][-1].get('print') and round(entry['A'][-1]['print'], 2)} "
                f"B={entry['B'][-1].get('print') and round(entry['B'][-1]['print'], 2)}")
    return timing


# ---------------------------------------------------------------- 乱码分析
BOX_CHARS = set("\ufffd\u25a0\u25a1\u25af\u2b1c\u25fe")


def normalize(text):
    return re.sub(r"\s+", "", text or "")


def text_metrics(pdf_text, source_text):
    pdf = normalize(pdf_text)
    src = normalize(source_text)
    fffd = pdf.count("\ufffd")
    boxes = sum(1 for ch in pdf if ch in BOX_CHARS)
    ratio = difflib.SequenceMatcher(None, src, pdf, autojunk=False).ratio() if src else 1.0
    src_counter = Counter(src)
    pdf_counter = Counter(pdf)
    missing = sum(max(0, n - pdf_counter.get(ch, 0)) for ch, n in src_counter.items())
    return {
        "pdf_chars": len(pdf), "source_chars": len(src),
        "ufffd": fffd, "box_chars": boxes,
        "seq_ratio_vs_source": round(ratio, 4),
        "missing_chars": missing,
        "missing_rate": round(missing / len(src), 4) if src else 0.0,
    }


ALIASES = {
    "simsun": "宋体", "nsimsun": "宋体", "songti": "宋体",
    "microsoftyahei": "微软雅黑", "microsoftyaheiuibold": "微软雅黑",
    "microsoftyaheiui": "微软雅黑", "kaiti": "楷体", "kaitisimplified": "楷体",
    "simhei": "黑体", "hei": "黑体", "fangsong": "仿宋", "fangsong_gb2312": "仿宋",
    "timesnewromanpsmt": "Times New Roman", "timesnewroman": "Times New Roman",
    "arialmt": "Arial", "arial": "Arial", "calibri": "Calibri",
    "cambriacambriamath": "Cambria Math", "cambriamath": "Cambria Math",
    "cambria": "Cambria", "nimbussansl": "宋体", "nimbusroman": "Times New Roman",
}


def base_font(name):
    text = str(name or "").split("+")[-1]
    text = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", text)
    return text.lower()


def strip_style(name):
    text = name
    changed = True
    while changed:
        changed = False
        for suffix in ("bolditalic", "boldoblique", "bold", "italic", "oblique",
                       "regular", "light", "medium", "psmt", "mt"):
            if text.endswith(suffix) and len(text) > len(suffix):
                text = text[: -len(suffix)]
                changed = True
    return text


def font_metrics(pdf_path, declared):
    doc = fitz.open(pdf_path)
    try:
        fonts = {}
        for pno in range(doc.page_count):
            for item in doc[pno].get_fonts():
                xref, ext, typ, basefont, name, enc = item
                key = basefont
                entry = fonts.setdefault(key, {"type": typ, "pages": set(),
                                               "embedded": bool(ext) and xref > 0})
                entry["pages"].add(pno + 1)
        declared_keys = {base_font(v) for v in declared}
        substituted = []
        for key in fonts:
            normalized = base_font(key)
            alias = ALIASES.get(normalized) or ALIASES.get(strip_style(normalized))
            if alias and base_font(alias) in declared_keys:
                continue
            if normalized in declared_keys or strip_style(normalized) in declared_keys:
                continue
            if normalized in ("symbol", "zapfdingbats"):
                continue
            substituted.append(key)
        missing_glyphs = count_missing_glyphs(doc)
        return {
            "fonts": {k: {"type": v["type"], "embedded": v["embedded"],
                          "pages": sorted(v["pages"])} for k, v in fonts.items()},
            "unembedded": [k for k, v in fonts.items() if not v["embedded"]],
            "substituted": substituted,
            "missing_glyph_uses": missing_glyphs["uses"],
            "glyph_checks": missing_glyphs["checked"],
            "pages": doc.page_count,
        }
    finally:
        doc.close()


def count_missing_glyphs(doc):
    buffers = {}
    for pno in range(doc.page_count):
        for item in doc[pno].get_fonts():
            xref, ext, typ, basefont, name, enc = item
            key = basefont.split("+")[-1]
            if key in buffers or not (xref and ext):
                continue
            try:
                buf = doc.extract_font(xref)[3]
                if buf:
                    buffers[key] = fitz.Font(fontbuffer=buf)
            except Exception:
                continue
    checked = 0
    uses = 0
    for pno in range(doc.page_count):
        for block in doc[pno].get_text("dict").get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    font = buffers.get(str(span.get("font", "")).split("+")[-1])
                    if font is None:
                        continue
                    for ch in span.get("text", ""):
                        if ord(ch) <= 32:
                            continue
                        checked += 1
                        try:
                            if not font.has_glyph(ord(ch)):
                                uses += 1
                        except Exception:
                            pass
    return {"checked": checked, "uses": uses}


def render_gray(pdf_path, pno, dpi=150):
    doc = fitz.open(pdf_path)
    try:
        page = doc[pno]
        pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]
        return np.asarray(arr, dtype=np.int16)
    finally:
        doc.close()


def compare_pdfs(pdf_a, pdf_b, prefix, heat_dir, dpi=150, threshold=32, save_heat=True):
    da, db = fitz.open(pdf_a), fitz.open(pdf_b)
    ratios = []
    worst = (0.0, 0, None)
    try:
        for pno in range(min(da.page_count, db.page_count)):
            pa = da[pno].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
            pb = db[pno].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
            if (pa.width, pa.height) != (pb.width, pb.height):
                ratios.append(1.0)
                continue
            ia = np.frombuffer(pa.samples, dtype=np.uint8).reshape(pa.height, pa.stride)[:, :pa.width]
            ib = np.frombuffer(pb.samples, dtype=np.uint8).reshape(pb.height, pb.stride)[:, :pb.width]
            diff = np.abs(ia.astype(np.int16) - ib.astype(np.int16))
            ratio = float((diff > threshold).mean())
            ratios.append(round(ratio, 5))
            if ratio >= worst[0]:
                worst = (ratio, pno, diff)
        if save_heat and worst[2] is not None:
            from PIL import Image
            diff = worst[2]
            rgb = np.zeros((diff.shape[0], diff.shape[1], 3), dtype=np.uint8)
            rgb[..., 0] = 255
            rgb[..., 1] = np.clip(255 - diff * 3, 0, 255)
            rgb[..., 2] = np.clip(255 - diff * 3, 0, 255)
            Image.fromarray(rgb).save(os.path.join(heat_dir, f"{prefix}_p{worst[1] + 1}.png"))
        return {
            "pages_compared": len(ratios),
            "mean_diff_ratio": round(float(np.mean(ratios)), 5) if ratios else 0.0,
            "max_diff_ratio": max(ratios) if ratios else 0.0,
            "max_page": worst[1] + 1,
            "per_page": ratios,
            "threshold": threshold,
            "dpi": dpi,
        }
    finally:
        da.close()
        db.close()


def save_shot(pdf_a, pdf_b, pno, path, dpi=150):
    from PIL import Image, ImageDraw
    imgs = []
    for pdf in (pdf_a, pdf_b):
        doc = fitz.open(pdf)
        pix = doc[pno].get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        doc.close()
        imgs.append(img)
    width = imgs[0].width + imgs[1].width + 20
    height = max(i.height for i in imgs) + 30
    canvas = Image.new("RGB", (width, height), (255, 255, 255))
    canvas.paste(imgs[0], (0, 30))
    canvas.paste(imgs[1], (imgs[0].width + 20, 30))
    draw = ImageDraw.Draw(canvas)
    draw.text((10, 8), "A: product pipeline", fill=(0, 0, 0))
    draw.text((imgs[0].width + 30, 8), "B: Word export + PDF print", fill=(0, 0, 0))
    canvas.save(path)


def analyze(samples):
    metrics = {}
    for item in samples:
        name = safe_name(item["name"])
        declared = list(item["fonts"])
        source_text = open(item["source_text"], encoding="utf-8").read()
        conv_pdf = os.path.join(OUT, "convA", name + ".pdf")
        exp_pdf = os.path.join(OUT, "exportB", name + ".pdf")
        print_a_pdf = os.path.join(OUT, "printA", name + "_r0.pdf")
        print_b_pdf = os.path.join(OUT, "printB", name + "_r0.pdf")
        entry = {"name": item["name"], "pages": item["pages"], "declared_fonts": declared}
        for tag, pdf in (("A", conv_pdf), ("B", exp_pdf)):
            doc = fitz.open(pdf)
            text = "\n".join(doc[p].get_text() for p in range(doc.page_count))
            doc.close()
            entry[tag] = {
                "text": text_metrics(text, source_text),
                "font": font_metrics(pdf, declared),
                "sampled_text": text[:300],
            }
        entry["diff_convert"] = compare_pdfs(conv_pdf, exp_pdf, name, os.path.join(OUT, "heat_conv"))
        entry["diff_print"] = compare_pdfs(print_a_pdf, print_b_pdf, name, os.path.join(OUT, "heat_print"))
        for tag, pdf in (("printA", print_a_pdf), ("printB", print_b_pdf)):
            doc = fitz.open(pdf)
            entry[tag] = {
                "pages": doc.page_count,
                "text_chars": len(normalize("".join(doc[p].get_text() for p in range(doc.page_count)))),
                "font_count": sum(len(doc[p].get_fonts()) for p in range(doc.page_count)),
                "image_pages": sum(1 for p in range(doc.page_count) if doc[p].get_images(full=True)),
            }
            doc.close()
        metrics[name] = entry
        log(f"analyze {name} convert_diff={entry['diff_convert']['mean_diff_ratio']} "
            f"print_diff={entry['diff_print']['mean_diff_ratio']}")
    return metrics


def shots(samples):
    picks = [("09_missing_font", 0), ("06_formula", 0), ("04_table", 0), ("07_header_footer", 0)]
    done = []
    for name, pno in picks:
        a = os.path.join(OUT, "printA", name + "_r0.pdf")
        b = os.path.join(OUT, "printB", name + "_r0.pdf")
        conv_a = os.path.join(OUT, "convA", name + ".pdf")
        conv_b = os.path.join(OUT, "exportB", name + ".pdf")
        if not (os.path.exists(a) and os.path.exists(b)):
            continue
        save_shot(a, b, pno, os.path.join(OUT, "shots", f"{name}_printed_p{pno + 1}.png"))
        save_shot(conv_a, conv_b, pno, os.path.join(OUT, "shots", f"{name}_convert_p{pno + 1}.png"))
        done.append(name)
    return done


# ---------------------------------------------------------------- 主流程
def summarize(timing):
    result = {}
    for name, entry in timing.items():
        row = {}
        for tag in ("A", "B"):
            runs = [e for e in entry[tag] if "error" not in e]
            if not runs:
                row[tag] = {"error": "no successful run"}
                continue
            def stats(key):
                vals = [r[key] for r in runs if key in r]
                return {"median": round(statistics.median(vals), 3),
                        "min": round(min(vals), 3), "max": round(max(vals), 3)} if vals else None
            row[tag] = {"pages": runs[0]["pages"]}
            for key in ("layout", "open", "render", "print"):
                if key in runs[0]:
                    row[tag][key] = stats(key)
            total = [sum(r.get(k, 0) for k in ("layout", "print")) for r in runs]
            row[tag]["print_total"] = {"median": round(statistics.median(total), 3),
                                       "min": round(min(total), 3), "max": round(max(total), 3)}
        result[name] = row
    return result


def main():
    samples = sample_list()
    env = {
        "printer": PRINTER,
        "dpi": DPI,
        "repeats": REPEATS,
        "python": sys.executable,
        "convert_python": PROD_PY,
        "fitz": fitz.VersionBind,
        "numpy": np.__version__,
        "samples": len(samples),
    }
    with open(os.path.join(OUT, "env.json"), "w", encoding="utf-8") as fh:
        json.dump(env, fh, ensure_ascii=False, indent=2)

    stage = os.environ.get("STAGE", "all")
    if stage in ("all", "convert"):
        timing_a = convert_a_single(samples)
        timing_batch = convert_a_batch(samples)
        timing_b_cold = convert_b_cold(samples)
        timing_b_steady = convert_b_steady(samples)
        with open(os.path.join(OUT, "convert_timing.json"), "w", encoding="utf-8") as fh:
            json.dump({"A_single": timing_a, "A_batch": timing_batch,
                       "B_cold": timing_b_cold, "B_steady": timing_b_steady},
                      fh, ensure_ascii=False, indent=2)
        log("转换阶段完成")
    if stage in ("all", "print"):
        timing_print = run_printing(samples)
        with open(os.path.join(OUT, "print_timing.json"), "w", encoding="utf-8") as fh:
            json.dump(timing_print, fh, ensure_ascii=False, indent=2)
        log("打印阶段完成")
    if stage in ("all", "analyze"):
        metrics = analyze(samples)
        with open(os.path.join(OUT, "metrics.json"), "w", encoding="utf-8") as fh:
            json.dump(metrics, fh, ensure_ascii=False, indent=2)
        log("分析阶段完成")
        log("截图: " + ",".join(shots(samples)))
    if stage == "all":
        with open(os.path.join(OUT, "print_summary.json"), "w", encoding="utf-8") as fh:
            json.dump(summarize(timing_print), fh, ensure_ascii=False, indent=2)
    log("全部完成")


if __name__ == "__main__":
    main()
