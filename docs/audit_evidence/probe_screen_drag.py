# -*- coding: utf-8 -*-
r"""屏幕级日志拖动残影探针：独立线程高频抓窗口，快速改变日志高度。

用法：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_screen_drag.py
"""
import ctypes
import os
import pathlib
import sys
import tempfile
import threading
import time

REPO = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="drag-probe-"))

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.GetWindowRect.restype = ctypes.c_int
user32.GetDC.restype = ctypes.c_void_p
user32.GetDC.argtypes = [ctypes.c_void_p]
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
gdi32.BitBlt.restype = ctypes.c_int
gdi32.BitBlt.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                         ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                         ctypes.c_uint]
gdi32.GetDIBits.restype = ctypes.c_int
gdi32.GetDIBits.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint,
                            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint), ("biWidth", ctypes.c_int),
                ("biHeight", ctypes.c_int), ("biPlanes", ctypes.c_ushort),
                ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
                ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_int),
                ("biYPelsPerMeter", ctypes.c_int), ("biClrUsed", ctypes.c_uint),
                ("biClrImportant", ctypes.c_uint)]


def capture_window(screen_dc, hwnd):
    rect = RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None, None, None
    left, top = rect.left, rect.top
    w, h = rect.right - left, rect.bottom - top
    if w <= 0 or h <= 0:
        return None, None, None
    memdc = gdi32.CreateCompatibleDC(screen_dc)
    bmp = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
    old = gdi32.SelectObject(memdc, bmp)
    gdi32.BitBlt(memdc, 0, 0, w, h, screen_dc, left, top, 0x00CC0020)
    bmi = BITMAPINFOHEADER()
    bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.biWidth = w
    bmi.biHeight = -h
    bmi.biPlanes = 1
    bmi.biBitCount = 32
    bmi.biCompression = 0
    buf = ctypes.create_string_buffer(w * h * 4)
    taken = gdi32.GetDIBits(memdc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.SelectObject(memdc, old)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(memdc)
    return (w, h, buf.raw) if taken else (None, None, None)


def dark_ratio(raw):
    if not raw or len(raw) < 8:
        return 0.0, 0.0
    dark = dim = sampled = 0
    for off in range(0, len(raw) - 3, 8):
        b, g, r = raw[off], raw[off + 1], raw[off + 2]
        sampled += 1
        if r < 16 and g < 16 and b < 16:
            dark += 1
        elif r < 64 and g < 64 and b < 64:
            dim += 1
    ratio = dark / max(1, sampled)
    if ratio > 1.0:
        return 1.0, 0.0
    return ratio, dim / max(1, sampled)


def save_png(raw, w, h, path):
    from PIL import Image
    img = Image.frombytes("RGBA", (w, h), raw, "raw", "BGRA")
    img.convert("RGB").save(path)


class WindowRecorder(threading.Thread):
    def __init__(self, hwnd):
        super().__init__(daemon=True)
        self.hwnd = hwnd
        self.stop_flag = threading.Event()
        self.frames = []
        self.worst_raw = None
        self.worst_size = None

    def run(self):
        screen_dc = user32.GetDC(None)
        t0 = time.perf_counter()
        while not self.stop_flag.is_set():
            t = time.perf_counter() - t0
            w, h, raw = capture_window(screen_dc, self.hwnd)
            if raw:
                d, m = dark_ratio(raw)
                self.frames.append((t, d, m))
                if not self.frames or d >= max(f[1] for f in self.frames):
                    self.worst_raw = raw
                    self.worst_size = (w, h)
            time.sleep(1 / 90)
        user32.ReleaseDC(None, screen_dc)

    def stop(self):
        self.stop_flag.set()
        self.join(timeout=3.0)


from ui_main import DuplexPrintApp, _create_root  # noqa: E402

root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+80+80")
root.update()
time.sleep(0.6)
root.update()

hwnd = ctypes.c_void_p(int(root.wm_frame(), 16))

# 基线
base = WindowRecorder(hwnd)
base.start()
time.sleep(0.8)
base.stop()
steady = [d for _, d, _ in base.frames]
print("STEADY dark avg=%.4f max=%.4f frames=%d" % (
    sum(steady) / max(1, len(steady)), max(steady or [0]), len(steady)))

# 快速拖动：与真实 <B1-Motion> 相同路径
rec = WindowRecorder(hwnd)
rec.start()
time.sleep(0.2)
heights = list(range(134, 330, 4)) + list(range(330, 134, -4)) + \
    list(range(134, 330, 8)) + list(range(330, 134, -8))
for value in heights:
    app._set_log_height_during_drag(value)
    root.update()
    time.sleep(1 / 100)
rec.stop()
print("DRAG frames:", len(rec.frames))
worst = max(rec.frames, key=lambda f: f[1], default=(0, 0, 0))
print("WORST drag dark=%.4f dim=%.4f at t=%.3f" % worst)
for t, d, m in rec.frames:
    if d > 0.03:
        print("  HIGH t=%.3f dark=%.4f dim=%.4f" % (t, d, m))
if rec.worst_raw and rec.worst_size:
    worst_path = os.path.join(tempfile.gettempdir(), "codearts-drag-worst.png")
    save_png(rec.worst_raw, *rec.worst_size, worst_path)
    print("WORST_FRAME_SAVED:", worst_path)

app.shutting_down = True
root.destroy()
print("DONE")
