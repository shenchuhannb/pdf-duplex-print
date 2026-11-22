# -*- coding: utf-8 -*-
r"""屏幕级黑块/残影复现探针 v2：独立线程高频抓屏。

黑闪预计发生在主线程同步处理 WM_SIZE / 布局期间——同线程探针在
root.update() 阻塞时无法抓帧。本探针用独立线程以 ~60fps 连续抓取
窗口区域，覆盖整个瞬态窗口期。

运行：
  D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\python.exe docs\audit_evidence\probe_screen_flicker2.py
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
os.environ.setdefault("APPDATA", tempfile.mkdtemp(prefix="flicker-screen2-"))

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
user32.PostMessageW.restype = ctypes.c_int
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                ctypes.c_ssize_t]


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


def capture_region(screen_dc, left, top, w, h):
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
    return buf.raw if taken else None


def dark_ratio(raw):
    dark = dim = sampled = 0
    for off in range(0, len(raw) - 3, 16):  # 每 4 像素采样
        b, g, r = raw[off], raw[off + 1], raw[off + 2]
        sampled += 1
        if r < 16 and g < 16 and b < 16:
            dark += 1
        elif r < 64 and g < 64 and b < 64:
            dim += 1
    return dark / max(1, sampled), dim / max(1, sampled)


class ScreenRecorder(threading.Thread):
    """后台高频抓屏（全屏），记录时间戳与暗像素比例。"""

    def __init__(self):
        super().__init__(daemon=True)
        self.stop_flag = threading.Event()
        self.frames = []  # (t, dark, dim)
        self.lock = threading.Lock()
        self.screen_dc = None

    def run(self):
        self.screen_dc = user32.GetDC(None)
        w = user32.GetSystemMetrics(0)
        h = user32.GetSystemMetrics(1)
        while not self.stop_flag.is_set():
            t = time.perf_counter()
            raw = capture_region(self.screen_dc, 0, 0, w, h)
            if raw:
                d, m = dark_ratio(raw)
                with self.lock:
                    self.frames.append((t, d, m))
            del raw
        user32.ReleaseDC(None, self.screen_dc)

    def stop(self):
        self.stop_flag.set()
        self.join(timeout=3.0)

    def window_of(self, t0, t1):
        with self.lock:
            return [(t - t0, d, m) for t, d, m in self.frames if t0 <= t <= t1]


# ---------------------------------------------------------------- 启动真实应用
from ui_main import DuplexPrintApp, _create_root  # noqa: E402

root = _create_root()
app = DuplexPrintApp(root, ui_scale_override=1.0)
app.preview_enabled = False
for attr in ("preview_after_id", "ui_poll_id", "printer_poll_id"):
    pid_ = getattr(app, attr, None)
    if pid_ is not None:
        root.after_cancel(pid_)
        setattr(app, attr, None)
root.geometry("1100x760+60+60")
root.update()
time.sleep(0.5)
root.update()

rec = ScreenRecorder()
rec.start()
time.sleep(0.5)

# ------------------------------------------------------------- 1) 最大化黑块
WM_SYSCOMMAND = 0x0112
SC_MAXIMIZE = 0xF030
SC_RESTORE = 0xF120
hwnd_c = ctypes.c_void_p(int(root.wm_frame(), 16))
t0 = time.perf_counter()
user32.PostMessageW(hwnd_c, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
end = t0 + 2.0
while time.perf_counter() < end:
    root.update()
    time.sleep(0.01)
rec.stop()
max_frames = rec.window_of(t0, t0 + 2.0)
print("MAXIMIZE total_frames:", len(max_frames))
baseline = [d for _, d, _ in max_frames if d > 0]
worst = max(max_frames, key=lambda f: f[1], default=(0, 0, 0))
print("WORST dark=%.4f dim=%.4f at t=%.3f" % worst)
for t, d, m in max_frames:
    if d > 0.02 or m > 0.02:
        print("  HIGH t=%.3f dark=%.4f dim=%.4f" % (t, d, m))

time.sleep(0.5)
root.update()

# --------------------------------------------------- 2) 日志分隔条快速拖动
rec2 = ScreenRecorder()
rec2.start()
time.sleep(0.3)
t1 = time.perf_counter()
heights = list(range(134, 320, 4)) + list(range(320, 134, -4))
for value in heights:
    app._set_log_height_during_drag(value)
    root.update()
    time.sleep(1 / 120)
rec2.stop()
drag_frames = rec2.window_of(t1, t1 + 10.0)
print("DRAG total_frames:", len(drag_frames))
worst_drag = max(drag_frames, key=lambda f: f[1], default=(0, 0, 0))
print("WORST drag dark=%.4f dim=%.4f at t=%.3f" % worst_drag)
for t, d, m in drag_frames:
    if d > 0.02 or m > 0.02:
        print("  HIGH t=%.3f dark=%.4f dim=%.4f" % (t, d, m))

app.shutting_down = True
root.destroy()
print("DONE")
