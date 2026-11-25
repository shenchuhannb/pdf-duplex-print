# -*- coding: utf-8 -*-
"""真实窗口几何记忆验证：启动 → 移动/缩放 → 关闭 → 重启 → 对比坐标。

约束：
  * 使用临时 APPDATA，设置与日志都落在临时目录，不污染用户配置；
  * 只操作窗口位置/尺寸，绝不点击打印或触发真实打印任务；
  * 通过 Win32 GetWindowRect / MoveWindow 取得物理像素坐标，并与
    window_geometry 的 DPI 换算结果对比，验证 150% DPI 下逻辑像素记忆。
"""
import ctypes
import json
import os
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import win32con  # noqa: E402
import win32gui  # noqa: E402

import window_geometry  # noqa: E402
from webapp.bridge import WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE  # noqa: E402

TITLE = "PDF 批量双面打印工具"
PYTHON = os.path.join(ROOT, ".venv", "Scripts", "pythonw.exe")
GEOMETRY_TOLERANCE = 4  # 允许 pywebview int() 截断与 DWM 边框带来的 1~2px 抖动


def set_dpi_awareness() -> None:
    """验证脚本自身声明 DPI 感知，才能看到与目标进程一致的物理像素。"""
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def find_windows(title_prefix):
    handles = []

    def _callback(hwnd, _extra):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowText(hwnd).startswith(title_prefix):
                handles.append(hwnd)
        return True

    win32gui.EnumWindows(_callback, None)
    return handles


def wait_new_window(before, timeout=60.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for hwnd in find_windows(TITLE):
            if hwnd not in before:
                return hwnd
        time.sleep(0.2)
    raise RuntimeError("等待应用窗口超时")


def get_rect(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return (left, top, right - left, bottom - top)


def settings_path(appdata):
    return os.path.join(appdata, "PDFDuplexPrint", "settings.json")


def read_settings(appdata):
    path = settings_path(appdata)
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return {}


def wait_geometry(appdata, predicate, timeout=20.0):
    deadline = time.monotonic() + timeout
    last = ""
    while time.monotonic() < deadline:
        last = str(read_settings(appdata).get("window_geometry") or "")
        parsed = window_geometry.parse_geometry(last)
        if parsed is not None and predicate(parsed):
            return last, parsed
        time.sleep(0.2)
    return last, window_geometry.parse_geometry(last)


def launch(appdata):
    env = dict(os.environ, APPDATA=appdata)
    before = find_windows(TITLE)
    process = subprocess.Popen([PYTHON, "webapp/main.py"], cwd=ROOT, env=env)
    hwnd = wait_new_window(before)
    # 等 WebView2 完成首屏与 Form 尺寸初始化，避免读到中间态。
    time.sleep(2.0)
    return process, hwnd


def close_window(process, hwnd, timeout=20.0):
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if process.poll() is not None:
            return True
        time.sleep(0.2)
    try:
        process.terminate()
        process.wait(timeout=5)
    except Exception:
        pass
    return False


def pywebview_physical(logical, scale):
    """复现 pywebview 的换算：位置/尺寸 = int(逻辑值 * 当前窗口 DPI 缩放)。"""
    width, height, x, y = logical
    return (int(x * scale), int(y * scale), int(width * scale), int(height * scale))


def main():
    if os.name != "nt":
        print("SKIP: 仅支持 Windows")
        return 2
    set_dpi_awareness()
    appdata = tempfile.mkdtemp(prefix="pdf-geometry-settings-")
    result = {"appdata": appdata, "pass": False}
    process = None
    process2 = None
    try:
        process, hwnd = launch(appdata)
        initial = get_rect(hwnd)
        scale = window_geometry.window_dpi_scale(hwnd)
        areas = window_geometry.list_work_areas()
        work_left, work_top, work_right, work_bottom = areas[0]
        print(f"DPI_SCALE: {scale}", flush=True)
        print(f"WORK_AREA: {(work_left, work_top, work_right, work_bottom)}", flush=True)
        print(f"INITIAL_RECT: {initial}", flush=True)

        # 移动 + 缩放为目标矩形（避开任务栏、大于最小尺寸）。
        min_w = int(WINDOW_MIN_SIZE[0] * scale)
        min_h = int(WINDOW_MIN_SIZE[1] * scale)
        target_w = min(min_w + 400, (work_right - work_left) - 240)
        target_h = min(min_h + 240, (work_bottom - work_top) - 240)
        target = (work_left + 120, work_top + 100, target_w, target_h)
        win32gui.MoveWindow(hwnd, target[0], target[1], target[2], target[3], True)
        print(f"TARGET_RECT: {target}", flush=True)

        # 目标逻辑几何：物理 → 逻辑 → 可见性/最小尺寸整理（与 bridge 完全一致）。
        # window_geometry 的几何元组是 (宽, 高, x, y)，MoveWindow 用的是
        # (x, y, 宽, 高)，这里显式换序，避免把 x/y 当成长宽。
        logical_raw = window_geometry.physical_geometry_to_logical(
            (target[2], target[3], target[0], target[1]), scale)
        expected_saved = window_geometry.format_geometry(
            *window_geometry.choose_logical_geometry(
                window_geometry.format_geometry(*logical_raw), areas, scale,
                WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE))
        expected = window_geometry.parse_geometry(expected_saved)
        print(f"EXPECTED_SAVED: {expected_saved}", flush=True)

        # 移动/缩放事件经 0.6s 去抖后落盘：关闭前必须已经保存。
        saved_at_move, parsed_move = wait_geometry(
            appdata, lambda value: abs(value[0] - expected[0]) <= GEOMETRY_TOLERANCE
            and abs(value[1] - expected[1]) <= GEOMETRY_TOLERANCE)
        result["saved_at_move"] = saved_at_move
        print(f"SAVED_AT_MOVE: {saved_at_move}", flush=True)
        if parsed_move is None:
            raise RuntimeError(f"移动后未保存几何：{saved_at_move!r}")

        moved_rect = get_rect(hwnd)
        result["moved_rect"] = moved_rect
        print(f"MOVED_RECT: {moved_rect}", flush=True)

        if not close_window(process, hwnd):
            raise RuntimeError("关闭应用超时")
        process = None
        final_saved = str(read_settings(appdata).get("window_geometry") or "")
        result["final_saved"] = final_saved
        print(f"FINAL_SAVED_AFTER_CLOSE: {final_saved}", flush=True)
        if final_saved != saved_at_move:
            raise RuntimeError("关闭时保存的最终几何与移动后不一致")

        # 重启：应以保存的逻辑几何恢复同等物理矩形。
        process2, hwnd2 = launch(appdata)
        restored = get_rect(hwnd2)
        restored_scale = window_geometry.window_dpi_scale(hwnd2)
        expected_rect = pywebview_physical(expected, restored_scale)
        result.update({
            "restored_rect": restored,
            "restored_scale": restored_scale,
            "expected_rect": expected_rect,
        })
        print(f"RESTORED_SCALE: {restored_scale}", flush=True)
        print(f"RESTORED_RECT: {restored}", flush=True)
        print(f"EXPECTED_RECT: {expected_rect}", flush=True)

        # 恢复位置必须来自保存记录，而不是巧合地停在默认居中位置。
        default_rect = pywebview_physical(
            window_geometry.centered_in(
                window_geometry.logical_work_areas(areas, restored_scale)[0],
                WINDOW_DEFAULT_SIZE, WINDOW_MIN_SIZE),
            restored_scale)
        differs_from_default = any(
            abs(restored[index] - default_rect[index]) > GEOMETRY_TOLERANCE
            for index in range(4))
        restored_ok = all(
            abs(restored[index] - expected_rect[index]) <= GEOMETRY_TOLERANCE
            for index in range(4))
        result["default_rect"] = default_rect
        result["differs_from_default"] = differs_from_default
        print(f"DEFAULT_RECT: {default_rect}", flush=True)
        print(f"DIFFERS_FROM_DEFAULT: {differs_from_default}", flush=True)

        # 保存的几何必须仍落在当前可见工作区内（多屏断开后的回退目标）。
        saved_visible = window_geometry.geometry_visible(
            expected, window_geometry.logical_work_areas(areas, restored_scale))
        result["saved_geometry_visible"] = saved_visible
        print(f"SAVED_GEOMETRY_VISIBLE: {saved_visible}", flush=True)

        # 最大化/最小化不得把整屏尺寸或 -32000 离屏坐标写进配置：
        # 两种状态下保存值都必须保持还原态。
        before_state = str(read_settings(appdata).get("window_geometry") or "")
        win32gui.ShowWindow(hwnd2, win32con.SW_MAXIMIZE)
        time.sleep(1.6)
        max_saved = str(read_settings(appdata).get("window_geometry") or "")
        maximized_rect = get_rect(hwnd2)
        win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
        time.sleep(1.2)
        win32gui.ShowWindow(hwnd2, win32con.SW_MINIMIZE)
        time.sleep(1.6)
        min_saved = str(read_settings(appdata).get("window_geometry") or "")
        win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
        time.sleep(1.2)
        state_ok = max_saved == before_state and min_saved == before_state
        result.update({
            "saved_before_maximize": before_state,
            "saved_while_maximized": max_saved,
            "saved_while_minimized": min_saved,
            "maximized_rect": maximized_rect,
            "state_restore_ok": state_ok,
        })
        print(f"MAXIMIZED_RECT: {maximized_rect}", flush=True)
        print(f"SAVED_WHILE_MAXIMIZED: {max_saved}", flush=True)
        print(f"SAVED_WHILE_MINIMIZED: {min_saved}", flush=True)
        print(f"STATE_RESTORE_OK: {state_ok}", flush=True)

        if not close_window(process2, hwnd2):
            raise RuntimeError("关闭重启后的应用超时")
        process2 = None
        result["pass"] = bool(
            restored_ok and differs_from_default and state_ok and saved_visible)
        print(f"GEOMETRY_RESTORE: {'PASS' if result['pass'] else 'FAIL'}", flush=True)
        print("GEOMETRY_RESULT:", json.dumps(result, ensure_ascii=False), flush=True)
        return 0 if result["pass"] else 1
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
        print(f"GEOMETRY_RESULT: {json.dumps(result, ensure_ascii=False)}", flush=True)
        print(f"GEOMETRY_RESTORE: FAIL ({result['error']})", flush=True)
        log_path = os.path.join(appdata, "PDFDuplexPrint", "logs", "app.log")
        if os.path.isfile(log_path):
            with open(log_path, "r", encoding="utf-8", errors="replace") as handle:
                print("APP_LOG_TAIL:", handle.read()[-2000:], flush=True)
        return 1
    finally:
        for current in (process, process2):
            if current is not None and current.poll() is None:
                try:
                    current.terminate()
                    current.wait(timeout=5)
                except Exception:
                    pass
        print(f"ISOLATED_APPDATA: {appdata}", flush=True)


if __name__ == "__main__":
    sys.exit(main())
