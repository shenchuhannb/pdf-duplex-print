# -*- coding: utf-8 -*-
"""通用工具：异常、页码解析、拖拽数据解析、单位换算。

本模块不依赖 GUI / 打印机，便于单元测试。
"""

import re
from typing import List

MM_PER_INCH = 25.4
PT_PER_INCH = 72.0


def mm_to_pt(value: float) -> float:
    """毫米 → PDF 点（72 dpi 下的 1/72 英寸）。"""
    return float(value) * PT_PER_INCH / MM_PER_INCH


def pt_to_mm(value: float) -> float:
    """PDF 点 → 毫米。"""
    return float(value) * MM_PER_INCH / PT_PER_INCH


# ----------------------------------------------------------------------------
# 自定义异常：message 直接展示给用户
# ----------------------------------------------------------------------------
class ToolError(Exception):
    """本工具的通用异常基类。"""


class PdfError(ToolError):
    """PDF 解析、读取相关错误。"""


class PrintError(ToolError):
    """打印任务相关错误。"""


class PrinterOfflineError(PrintError):
    """打印机离线。"""


class PrintTimeoutError(PrintError):
    """监听打印队列超时（调用方可用固定延时兜底）。"""


class CancelledByUser(ToolError):
    """用户中止。"""


class SettingsError(ToolError):
    """设置读写错误。"""


# ----------------------------------------------------------------------------
# 页码范围解析（仅语法层，不关心具体文件页数）
# ----------------------------------------------------------------------------
def parse_page_range(text: str) -> List[int]:
    """解析页码范围，例如 "1-5,8,10-12"。

    兼容中文标点：，；、；－—–～~；允许空格。
    返回升序、去重后的页码列表；超出文件页数的裁剪由调用方负责。
    """
    if not text or not str(text).strip():
        raise PdfError("请输入页码范围，例如：1-5,8,10-12")

    normalized = str(text)
    for src, dst in (
        ("，", ","), ("；", ","), ("、", ","), (";", ","),
        ("－", "-"), ("—", "-"), ("–", "-"), ("～", "-"), ("~", "-"),
        ("．", "."), ("　", " "),
    ):
        normalized = normalized.replace(src, dst)
    normalized = re.sub(r"\s+", "", normalized)

    pages = set()
    for token in normalized.split(","):
        if not token:
            continue
        match = re.fullmatch(r"(\d+)-(\d+)", token)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            if start > end:
                start, end = end, start
            if start < 1:
                raise PdfError(f"页码范围不能小于 1：{token}")
            # 单段区间上限，防止误输入（如 1-99999）卡死程序
            if end - start + 1 > 20000:
                raise PdfError(f"页码范围过大：{token}")
            pages.update(range(start, end + 1))
            continue
        if re.fullmatch(r"\d+", token):
            number = int(token)
            if number < 1:
                raise PdfError(f"页码范围不能小于 1：{token}")
            pages.add(number)
            continue
        raise PdfError(f"页码范围格式有误：{token}（示例：1-5,8,10-12）")

    if not pages:
        raise PdfError("没有解析到有效页码，请检查输入。")
    return sorted(pages)


def format_page_list(pages: List[int], limit: int = 24) -> str:
    """把页码列表格式化为便于阅读的字符串，过长时截断。"""
    if not pages:
        return "（无）"
    if len(pages) <= limit:
        return "、".join(str(p) for p in pages)
    head = "、".join(str(p) for p in pages[:limit])
    return f"{head} …（共 {len(pages)} 页）"


# ----------------------------------------------------------------------------
# 拖拽数据解析
# ----------------------------------------------------------------------------
def split_drop_data(data: str) -> List[str]:
    """解析 tkinterdnd2 / TkDND 的拖拽数据串为路径列表。

    ★ 关键修复：不能使用 Tcl 的 splitlist 直接解析。
      Windows 路径中的反斜杠（如 C:\\Users\\a b\\1.pdf）会被 Tcl 当作
      转义字符吃掉，导致路径损坏（这正是上一轮“大括号含空格/多文件混合”
      两个用例失败的原因）。这里手写解析：
        - 大括号 {} 内视为一个整体，内部空格、反斜杠原样保留；
        - 双引号 "..." 内同样视为一个整体（兼容部分拖拽源）；
        - 大括号外以空白分隔；
        - 不做任何反斜杠转义处理。
    """
    text = (data or "").strip()
    if not text:
        return []

    paths: List[str] = []
    token: List[str] = []
    in_brace = in_quote = False
    has_token = False

    def flush() -> None:
        nonlocal token, has_token
        if has_token:
            value = "".join(token)
            if value.strip():
                paths.append(value)
        token = []
        has_token = False

    for ch in text:
        if in_brace:
            if ch == "}":
                in_brace = False
            else:
                token.append(ch)
            continue
        if in_quote:
            if ch == '"':
                in_quote = False
            else:
                token.append(ch)
            continue
        if ch == "{":
            in_brace = True
            has_token = True
            continue
        if ch == '"':
            in_quote = True
            has_token = True
            continue
        if ch.isspace():
            flush()
            continue
        token.append(ch)
        has_token = True
    flush()
    return paths


def is_pdf_path(path: str) -> bool:
    """判断路径是否为 PDF（不区分大小写）。"""
    return str(path).strip().lower().endswith(".pdf")


def is_word_path(path: str) -> bool:
    """判断路径是否为 Office/文本文档（.doc/.docx/.rtf/.txt/.odt）。"""
    return str(path).strip().lower().endswith(
        (".doc", ".docx", ".rtf", ".txt", ".odt"))


def is_wps_path(path: str) -> bool:
    """判断路径是否为 WPS 文字文档（.wps）。"""
    return str(path).strip().lower().endswith(".wps")
