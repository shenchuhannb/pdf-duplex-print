# -*- coding: utf-8 -*-
"""打印设置的本地记忆（JSON）。

保存位置：%APPDATA%\\PDFDuplexPrint\\settings.json
关闭窗口、开始打印、添加文件时自动保存；启动时恢复；“恢复默认”可重置。
"""

import json
import os
import tempfile
from typing import Any, Dict

from core_utils import SettingsError

APP_DATA_FOLDER = "PDFDuplexPrint"
SETTINGS_FILENAME = "settings.json"
# 打印流程三选一：单一枚举，避免“整批翻面 + 每文件暂停”这类互相矛盾的组合。
PRINT_FLOW_CHOICES = ("batch_all", "batch_take", "per_file")
DEFAULT_PRINT_FLOW = "batch_all"

DEFAULT_SETTINGS: Dict[str, Any] = {
    # 打印机
    "printer": "",
    # 打印范围
    "range_mode": "all",          # all/current/custom/odd/even/recent
    "range_text": "1-5,8,10-12",
    "last_range": "",
    "per_file": False,            # 逐文件单独设置范围与份数
    # 份数 / 分页
    "copies": 1,
    "collate": "collate",         # collate（整份连续）/ uncollated（逐页重复）
    # 打印方式：首次启动默认为“手动双面”（本工具核心能力）。
    # 只有“恢复 Acrobat 等效默认”会切回单面（Acrobat 基线本身是单面输出）。
    "manual_duplex": True,        # True = 手动双面（默认）；False = 单面
    # 打印流程：batch_all=整批翻面·打完全部再取纸（默认）；
    # batch_take=整批翻面·逐文件取纸；per_file=逐文件翻面·逐文件取纸
    "print_flow": DEFAULT_PRINT_FLOW,
    "copy_exec_on": False,        # 份数>1 时按份逐份执行（每份各翻一次）
    # 图像与质量
    "color_mode": "color",        # color/gray/mono
    "ink_saver": False,
    "print_quality": "normal",    # draft/normal/high
    "dpi": 300,
    # Acrobat 的“Comments and Forms”默认是 Document（不含批注），这里保持一致
    "annots": False,
    # 缩放与版式
    "scale_mode": "shrink",       # shrink（Acrobat 默认“收缩到可打印区”）/ fit / actual / custom / paper
    "scale_percent": 100,
    "nup": 1,
    "nup_border": False,
    "page_order": "row",          # row/column
    "booklet": False,
    # 纸张与方向
    "paper_name": "A4",
    "paper_width_mm": 210.0,
    "paper_height_mm": 297.0,
    "source_name": "自动选择",
    "orientation": "auto",        # auto/portrait/landscape
    "auto_rotate": True,
    # 边距与裁切（毫米）
    # 默认 0：与 Acrobat“直接打印”一致——只在驱动返回的可打印区域内按需收缩，
    # 不额外加白边。需要装订留白时在“高级设置”里自行加大。
    "margin_left_mm": 0.0,
    "margin_top_mm": 0.0,
    "margin_right_mm": 0.0,
    "margin_bottom_mm": 0.0,
    "crop_left_mm": 0.0,
    "crop_top_mm": 0.0,
    "crop_right_mm": 0.0,
    "crop_bottom_mm": 0.0,
    # 页面处理
    "reverse": False,
    "skip_blank": False,
    # 打印完成判断
    "wait_mode": "queue",         # queue/delay
    "delay_seconds": 15.0,
    # 输出
    "auto_open": False,
    "keep_output": False,
    "output_dir": "",
    # 其他
    "recent_files": [],
    "recent_ranges": [],
    "ui_style": "A",              # 仅保留 A 浅色专业风（B/C 定义保留但不再暴露）
    # 窗口几何："WxH+X+Y"，启动时会校验是否仍在可见工作区内，否则居中
    "window_geometry": "",
    "preview_view": "single",     # single（单页）/ spread（双页并排）
}

# 数值型字段的允许范围（防止损坏的配置文件导致异常）
_NUMERIC_LIMITS = {
    "copies": (1, 99),
    "dpi": (72, 1200),
    "scale_percent": (10, 500),
    "nup": (1, 16),
    "paper_width_mm": (10.0, 2000.0),
    "paper_height_mm": (10.0, 2000.0),
    "margin_left_mm": (0.0, 100.0),
    "margin_top_mm": (0.0, 100.0),
    "margin_right_mm": (0.0, 100.0),
    "margin_bottom_mm": (0.0, 100.0),
    "crop_left_mm": (0.0, 100.0),
    "crop_top_mm": (0.0, 100.0),
    "crop_right_mm": (0.0, 100.0),
    "crop_bottom_mm": (0.0, 100.0),
    "delay_seconds": (0.0, 600.0),
}


def settings_path() -> str:
    """返回配置文件路径（优先 APPDATA，失败时退回临时目录）。"""
    base = os.environ.get("APPDATA") or tempfile.gettempdir()
    return os.path.join(base, APP_DATA_FOLDER, SETTINGS_FILENAME)


def default_settings() -> Dict[str, Any]:
    """返回默认设置的深拷贝（避免调用方修改全局常量）。"""
    return json.loads(json.dumps(DEFAULT_SETTINGS))


def acrobat_defaults() -> Dict[str, Any]:
    """与 Adobe Acrobat / Reader “直接点打印、不做任何修改”等效的基线。

    核对来源（Adobe 官方帮助）：
      * 页面缩放默认 = Shrink To Printable Area（收缩到可打印区，只缩不放）；
      * “自动旋转并居中”默认开启；
      * “Comments and Forms”默认 = Document（不打印批注/表单/戳记）；
      * “Choose paper source by PDF page size”默认关闭 → 使用打印机默认纸张；
      * 份数 1、自动分页（整份连续）、彩色、标准质量。
      * 边距 0mm：Acrobat 不额外加白边，只在驱动返回的可打印区域内按需收缩。
    与本工具核心差异：本工具首次启动默认手动双面；Acrobat 基线相当于
    单面，所以“恢复 Acrobat 等效默认”会把输出方式切回单面。
    """
    data = default_settings()
    data.update({
        "range_mode": "all",
        "copies": 1,
        "collate": "collate",
        "manual_duplex": False,
        "print_flow": DEFAULT_PRINT_FLOW,
        "copy_exec_on": False,
        "color_mode": "color",
        "ink_saver": False,
        "print_quality": "normal",
        "annots": False,
        "scale_mode": "shrink",
        "scale_percent": 100,
        "nup": 1,
        "booklet": False,
        "nup_border": False,
        "page_order": "row",
        "paper_name": "A4",
        "source_name": "自动选择",
        "orientation": "auto",
        "auto_rotate": True,
        "reverse": False,
        "skip_blank": False,
        "per_file": False,
        "wait_mode": "queue",
        "auto_open": False,
        "keep_output": False,
    })
    return data


def migrate_print_flow(data: Dict[str, Any]) -> str:
    """把旧配置的 batch_mode/exec_mode + per_file_pause 映射为单一 print_flow。

    映射规则（旧键缺省时按旧版行为处理）：
      * exec_mode/batch_mode == "by_file"            -> per_file
      * exec_mode/batch_mode == "batch" + pause=true -> batch_take
      * exec_mode/batch_mode == "batch" + pause=false-> batch_all
      * 无旧键或值未知                                -> batch_all
    新键 print_flow 合法时优先使用新键，保证老文件读起来不崩。
    """
    if not isinstance(data, dict):
        return DEFAULT_PRINT_FLOW
    flow = str(data.get("print_flow") or "")
    if flow in PRINT_FLOW_CHOICES:
        return flow
    legacy = data.get("exec_mode", data.get("batch_mode"))
    legacy = str(legacy or "")
    if legacy == "by_file":
        return "per_file"
    if legacy == "batch":
        # 旧版 _print_batch_backs 在缺省 per_file_pause 时按 True 处理
        paused = bool(data.get("per_file_pause", True))
        return "batch_take" if paused else "batch_all"
    return DEFAULT_PRINT_FLOW


def _sanitize(data: Dict[str, Any]) -> Dict[str, Any]:
    """把读取到的配置夹到合法范围，未知字段忽略。"""
    result = default_settings()
    if not isinstance(data, dict):
        return result
    # 先做流程枚举迁移，避免旧 settings.json 里的 exec_mode/per_file_pause
    # 因为新键不存在而被静默丢弃。
    result["print_flow"] = migrate_print_flow(data)
    for key in list(result.keys()):
        if key not in data:
            continue
        value = data[key]
        default = result[key]
        if isinstance(default, bool):
            result[key] = bool(value)
        elif isinstance(default, (int, float)):
            try:
                number = float(value)
            except (TypeError, ValueError):
                continue
            low, high = _NUMERIC_LIMITS.get(key, (None, None))
            if low is not None:
                number = min(max(number, low), high)
            result[key] = int(number) if isinstance(default, int) else number
        elif isinstance(default, str):
            result[key] = str(value)
        elif isinstance(default, list):
            if isinstance(value, list):
                # 与 add_recent 的 limit 保持一致，避免读入比写出更宽松
                result[key] = [str(item) for item in value][:10]
    # 非法 print_flow 字符串在通用循环里会被原样保留，这里再校正一次枚举。
    result["print_flow"] = migrate_print_flow(data)
    return result


def load_settings(path: str = "") -> Dict[str, Any]:
    """读取配置；文件不存在/损坏时返回默认设置并保证不崩溃。"""
    target = path or settings_path()
    try:
        with open(target, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except FileNotFoundError:
        return default_settings()
    except (OSError, ValueError):
        return default_settings()
    return _sanitize(data)


def save_settings(data: Dict[str, Any], path: str = "") -> str:
    """保存配置（先写临时文件再替换，避免写一半导致文件损坏）。"""
    target = path or settings_path()
    try:
        os.makedirs(os.path.dirname(target), exist_ok=True)
        fd, temp_path = tempfile.mkstemp(prefix="settings_", suffix=".json",
                                         dir=os.path.dirname(target))
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(_sanitize(data), handle, ensure_ascii=False, indent=2)
            os.replace(temp_path, target)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
        return target
    except OSError as exc:
        raise SettingsError(f"无法保存设置：{target}\n{exc}") from exc


def add_recent(items: list, value: str, limit: int = 10) -> list:
    """把 value 放到最近列表首位并去重、截断。"""
    value = str(value or "").strip()
    if not value:
        return list(items)[:limit]
    result = [item for item in items if item != value]
    result.insert(0, value)
    return result[:limit]
