# -*- coding: utf-8 -*-
"""界面视觉主题（原生 ttk 自绘，不依赖第三方主题库）。

设计目标：
  * 现代、干净、有整体感：统一的 8px 网格间距、单一强调色 + 中性灰阶；
  * 去掉 ttk 默认的灰色隆起按钮：全部改为 flat + 明确的悬停/按下/禁用三态；
  * 输入框、下拉框、卡片统一 1px 浅灰描边与内边距；
  * 字体统一 Microsoft YaHei UI，四档字号（标题/小标题/正文/辅助说明）。

唯一风格：A 浅色专业风（用户定稿）。
B/C 色板的字典定义保留在本文件中仅作历史资料，不再对外暴露；
命令行 --style 仅接受 A，旧配置里的 ui_style=B/C 会被 resolve_style 收敛回 A。
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict

FONT_FAMILY = "Microsoft YaHei UI"
# 四档字号阶梯（100% DPI 下的实际像素高度）：
#   标题 15pt≈20px / 小标题 10pt≈13px / 正文 9pt≈12px / 辅助 8pt≈11px
# 行高由控件 padding 统一控制（见 GAP_* 与按钮/入口 padding）。
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 10, "bold")
FONT_BODY = (FONT_FAMILY, 9)
FONT_SMALL = (FONT_FAMILY, 8)
# 数字专用字体：Segoe UI 的拉丁数字为等宽数字（tabular），用于顶部指标徽标，
# 保证文件数/页数/张数在不同宽度变化时竖直方向仍然对齐。
FONT_NUM = ("Segoe UI", 13, "bold")

# 间距规范（两档：4px 细间距 / 8px 常规间距；卡片间距另用 12px）
GAP_S = 4
GAP = 8
GAP_CARD = 12
PAD_CARD = (12, 10)     # 卡片内边距：左右 12 / 上下 10
ROW_HEIGHT = 28         # 表格行高

PALETTES: Dict[str, Dict[str, str]] = {
    # A 浅色专业风：白卡片 + 冷灰背景 + 蓝色强调
    "A": {
        "bg": "#eef1f5",
        "surface": "#ffffff",
        "surface_alt": "#f5f7fa",
        "border": "#dfe4ea",
        "text": "#1f2933",
        # 辅助文字：在白底 4.96:1、在浅灰底 4.62:1，均达到 WCAG AA（旧值 4.25 不达标）
        "muted": "#66717f",
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "accent_pressed": "#1e40af",
        "accent_soft": "#e8f0fe",
        "accent_text": "#ffffff",
        "danger": "#c01c28",
        # 低饱和危险色（取消/清空等），与主操作的高饱和蓝明确区分
        "danger_text": "#b04a52",
        "danger_soft": "#fdf3f3",
        "danger_border": "#e8cfd2",
        "success": "#1a7f37",
        "warning": "#9a6200",
        "input_bg": "#ffffff",
        "header_bg": "#ffffff",
        "tree_stripe": "#f7f9fc",
        "tree_select": "#dbeafe",
        "shadow": "#d7dde5",
        # 纸张预览的轻投影（比 border 更浅，避免预览区出现厚重阴影）
        "shadow_light": "#e9edf2",
        "active_row": "#eef4ff",
    },
    # B 深色专业风：深灰蓝背景 + 亮蓝强调
    "B": {
        "bg": "#14181f",
        "surface": "#1c222b",
        "surface_alt": "#232b36",
        "border": "#313b47",
        "text": "#e6edf3",
        "muted": "#9aa7b4",
        "accent": "#4c8dff",
        "accent_hover": "#6ba1ff",
        "accent_pressed": "#3b78e0",
        "accent_soft": "#24344f",
        "accent_text": "#0b1220",
        "danger": "#ff6b6b",
        "success": "#5dd39e",
        "warning": "#f2b84b",
        "input_bg": "#232b36",
        "header_bg": "#1c222b",
        "tree_stripe": "#1f2630",
        "tree_select": "#2c3e5c",
        "shadow": "#0f1319",
    },
    # C 极简无边框风：近白背景 + 纯黑白 + 极细分隔
    "C": {
        "bg": "#fafafa",
        "surface": "#ffffff",
        "surface_alt": "#f4f4f5",
        "border": "#e5e5e5",
        "text": "#111111",
        "muted": "#777777",
        "accent": "#111111",
        "accent_hover": "#000000",
        "accent_pressed": "#000000",
        "accent_soft": "#f0f0f0",
        "accent_text": "#ffffff",
        "danger": "#b42318",
        "success": "#2e7d32",
        "warning": "#8a5a00",
        "input_bg": "#ffffff",
        "header_bg": "#ffffff",
        "tree_stripe": "#fafafa",
        "tree_select": "#ebebeb",
        "shadow": "#ededed",
    },
}

DEFAULT_STYLE = "A"
_current_style = DEFAULT_STYLE
_current_palette = PALETTES[DEFAULT_STYLE]


def palette(style: str = "") -> Dict[str, str]:
    """返回指定风格的色板；未知风格回退默认。"""
    return PALETTES.get((style or _current_style).upper(), PALETTES[DEFAULT_STYLE])


def current_style() -> str:
    return _current_style


def resolve_style(explicit: str = "") -> str:
    """固定返回风格 A（浅色专业）。

    B/C 色板定义保留在 PALETTES 中仅作历史资料，界面不再暴露切换入口；
    旧的配置值（ui_style=B/C）也会被收敛回 A，避免出现未适配的配色。
    """
    return DEFAULT_STYLE


def font(size: str = "body") -> tuple:
    """四档字号。"""
    return {
        "title": FONT_TITLE, "subtitle": FONT_SUBTITLE,
        "body": FONT_BODY, "small": FONT_SMALL,
    }.get(size, FONT_BODY)


def _safe(style: ttk.Style, name: str, **kwargs) -> None:
    """容忍个别 ttk 主题不支持的配置项。"""
    try:
        style.configure(name, **kwargs)
    except tk.TclError:
        pass


def _safe_map(style: ttk.Style, name: str, **kwargs) -> None:
    try:
        style.map(name, **kwargs)
    except tk.TclError:
        pass


def apply_theme(root: tk.Tk, style_key: str = "") -> Dict[str, str]:
    """应用主题到整个 Tk 窗口，返回色板。

    使用 clam 作为基础主题：只有 clam 允许完整覆盖背景/边框颜色；
    vista 主题会忽略大部分自定义颜色。
    """
    global _current_style, _current_palette
    _current_style = resolve_style(style_key or _current_style)
    colors = palette(_current_style)
    _current_palette = colors

    st = ttk.Style(root)
    try:
        if "clam" in st.theme_names():
            st.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(background=colors["bg"])
    _safe(st, ".", font=FONT_BODY, background=colors["bg"], foreground=colors["text"])

    # --- 容器 ---------------------------------------------------------------
    _safe(st, "TFrame", background=colors["surface"])
    _safe(st, "Page.TFrame", background=colors["bg"])
    _safe(st, "Card.TFrame", background=colors["surface"])
    _safe(st, "Header.TFrame", background=colors["header_bg"])
    _safe(st, "Bar.TFrame", background=colors["surface"])

    # --- 文字 ---------------------------------------------------------------
    for name, fg in (("TLabel", colors["text"]),
                     ("Muted.TLabel", colors["muted"]),
                     ("Title.TLabel", colors["text"]),
                     ("Accent.TLabel", colors["accent"]),
                     ("Danger.TLabel", colors["danger"])):
        _safe(st, name, background=colors["surface"], foreground=fg, font=FONT_BODY)
    _safe(st, "Title.TLabel", font=FONT_TITLE)
    _safe(st, "Subtitle.TLabel", background=colors["surface"], foreground=colors["text"],
          font=FONT_SUBTITLE)
    _safe(st, "Muted.TLabel", font=FONT_SMALL)
    _safe(st, "Metric.TLabel", background=colors["surface_alt"], foreground=colors["text"],
          font=FONT_SUBTITLE, padding=(10, 2))
    _safe(st, "MetricCaption.TLabel", background=colors["surface_alt"],
          foreground=colors["muted"], font=FONT_SMALL)

    # --- 卡片（替代厚重的 LabelFrame 浮雕）-----------------------------------
    _safe(st, "TLabelframe", background=colors["surface"], bordercolor=colors["border"],
          lightcolor=colors["border"], darkcolor=colors["border"], borderwidth=1,
          relief="flat", padding=(12, 8))
    _safe(st, "Card.TLabelframe", background=colors["surface"], bordercolor=colors["border"],
          lightcolor=colors["border"], darkcolor=colors["border"], borderwidth=1,
          relief="flat", padding=(12, 10))
    _safe(st, "TLabelframe.Label", background=colors["surface"], foreground=colors["muted"],
          font=FONT_SUBTITLE, padding=(2, 0))
    _safe(st, "Card.TLabelframe.Label", background=colors["surface"], foreground=colors["text"],
          font=FONT_SUBTITLE, padding=(2, 0))

    # --- 按钮：flat + 悬停/按下/禁用三态；主按钮 36px、次按钮 30px ----------
    # width=0：按文字自适应宽度（ttk 默认 width=-11 会让每个按钮至少有 11 字符宽，
    # 中文界面下按钮会被撑得过大，窄窗口里整行放不下而截断）
    # focusthickness=1：键盘 Tab 聚焦时显示细描边（focus ring），不改变按钮高度。
    _safe(st, "TButton", background=colors["surface_alt"], foreground=colors["text"],
          bordercolor=colors["border"], lightcolor=colors["surface_alt"],
          darkcolor=colors["surface_alt"], focusthickness=1,
          focuscolor=colors["accent"], relief="flat",
          padding=(12, 4), font=FONT_BODY, width=0)
    _safe_map(st, "TButton",
              background=[("disabled", colors["surface_alt"]), ("pressed", colors["border"]),
                          ("active", colors["accent_soft"])],
              foreground=[("disabled", colors["muted"]), ("active", colors["text"])],
              bordercolor=[("active", colors["accent"])],
              relief=[("pressed", "flat")])
    # 主按钮更高（padding 12,9 ≈ 36px）且加粗，靠高度+颜色形成唯一主操作
    _safe(st, "Accent.TButton", background=colors["accent"], foreground=colors["accent_text"],
          bordercolor=colors["accent"], lightcolor=colors["accent"],
          darkcolor=colors["accent"], focusthickness=1, focuscolor=colors["accent"],
          relief="flat", padding=(22, 9), font=(FONT_FAMILY, 9, "bold"), width=0)
    _safe_map(st, "Accent.TButton",
              background=[("disabled", colors["border"]), ("pressed", colors["accent_pressed"]),
                          ("active", colors["accent_hover"])],
              foreground=[("disabled", colors["muted"]), ("pressed", colors["accent_text"]),
                          ("active", colors["accent_text"])],
              bordercolor=[("disabled", colors["border"]), ("active", colors["accent_hover"]),
                           ("pressed", colors["accent_pressed"])])
    _safe(st, "Ghost.TButton", background=colors["surface"], foreground=colors["muted"],
          bordercolor=colors["surface"], lightcolor=colors["surface"],
          darkcolor=colors["surface"], focusthickness=1, focuscolor=colors["accent"],
          relief="flat", padding=(10, 4), font=FONT_BODY, width=0)
    _safe_map(st, "Ghost.TButton",
              background=[("active", colors["surface_alt"]), ("pressed", colors["border"]),
                          ("disabled", colors["surface"])],
              foreground=[("active", colors["text"]), ("disabled", colors["muted"])],
              bordercolor=[("active", colors["border"])])
    # 次主按钮（如“＋ 添加文件”）：浅蓝底 + 蓝字，比默认灰按钮更醒目，
    # 但仍低于实心 Accent 主操作（开始打印），形成“主 > 次主 > 普通 > 危险”层级。
    _safe(st, "Soft.TButton", background=colors["accent_soft"],
          foreground=colors["accent"], bordercolor=colors["accent_soft"],
          lightcolor=colors["accent_soft"], darkcolor=colors["accent_soft"],
          focusthickness=1, focuscolor=colors["accent"], relief="flat",
          padding=(12, 4), font=(FONT_FAMILY, 9, "bold"), width=0)
    _safe_map(st, "Soft.TButton",
              background=[("disabled", colors["surface_alt"]),
                          ("pressed", colors["accent_soft"]),
                          ("active", colors["accent_soft"])],
              foreground=[("disabled", colors["muted"]),
                          ("active", colors["accent_pressed"])],
              bordercolor=[("active", colors["accent"]), ("disabled", colors["border"])])
    # 危险操作：低饱和红 + 浅红底，只在 hover/按下时出现底色，不抢主操作视觉
    _safe(st, "Danger.TButton", background=colors["surface"],
          foreground=colors["danger_text"], bordercolor=colors["danger_border"],
          lightcolor=colors["surface"], darkcolor=colors["surface"],
          focusthickness=1, focuscolor=colors["danger_text"], relief="flat",
          padding=(12, 4), width=0)
    _safe_map(st, "Danger.TButton",
              background=[("active", colors["danger_soft"]),
                          ("disabled", colors["surface_alt"])],
              foreground=[("disabled", colors["muted"]), ("active", colors["danger_text"])],
              bordercolor=[("active", colors["danger_text"]),
                           ("disabled", colors["border"])])

    # --- 输入类控件 ---------------------------------------------------------
    for name in ("TEntry", "TSpinbox", "TCombobox"):
        _safe(st, name, fieldbackground=colors["input_bg"], background=colors["input_bg"],
              foreground=colors["text"], bordercolor=colors["border"],
              lightcolor=colors["border"], darkcolor=colors["border"],
              insertcolor=colors["text"], relief="flat", padding=(8, 2), arrowcolor=colors["muted"])
    _safe_map(st, "TEntry",
              bordercolor=[("focus", colors["accent"])],
              lightcolor=[("focus", colors["accent"])],
              darkcolor=[("focus", colors["accent"])],
              fieldbackground=[("disabled", colors["surface_alt"])],
              foreground=[("disabled", colors["muted"])])
    # 校验失败态：仅换描边颜色（不增加边框厚度），配合下方预留错误行使用
    _safe(st, "Error.TEntry", fieldbackground=colors["input_bg"], foreground=colors["text"],
          bordercolor=colors["danger"], lightcolor=colors["danger"],
          darkcolor=colors["danger"], insertcolor=colors["text"], relief="flat",
          padding=(8, 2))
    _safe_map(st, "Error.TEntry",
              bordercolor=[("focus", colors["danger"])],
              lightcolor=[("focus", colors["danger"])],
              darkcolor=[("focus", colors["danger"])])
    _safe_map(st, "TCombobox",
              bordercolor=[("focus", colors["accent"]), ("readonly", colors["border"])],
              lightcolor=[("focus", colors["accent"])],
              darkcolor=[("focus", colors["accent"])],
              fieldbackground=[("readonly", colors["surface"]),
                               ("disabled", colors["surface_alt"])],
              foreground=[("readonly", colors["text"]), ("disabled", colors["muted"]),
                          ("active", colors["text"])],
              arrowcolor=[("active", colors["accent"])])
    _safe_map(st, "TSpinbox",
              bordercolor=[("focus", colors["accent"])],
              fieldbackground=[("disabled", colors["surface_alt"])],
              foreground=[("disabled", colors["muted"])])
    # 份数专用：数字与右侧上下箭头之间必须留出空隙。默认 TSpinbox 的 8px 横向
    # padding 在高 DPI 下会被箭头吃掉，出现“9999 贴住箭头/末位被压”；这里加大
    # 横向 padding 并略增纵向 padding，配合 width=5 保证 4 位数字完整显示。
    _safe(st, "Copies.TSpinbox", fieldbackground=colors["input_bg"],
          background=colors["input_bg"], foreground=colors["text"],
          bordercolor=colors["border"], lightcolor=colors["border"],
          darkcolor=colors["border"], insertcolor=colors["text"], relief="flat",
          padding=(10, 5), arrowcolor=colors["muted"])
    _safe_map(st, "Copies.TSpinbox",
              bordercolor=[("focus", colors["accent"])],
              lightcolor=[("focus", colors["accent"])],
              darkcolor=[("focus", colors["accent"])],
              fieldbackground=[("disabled", colors["surface_alt"])],
              foreground=[("disabled", colors["muted"])],
              arrowcolor=[("active", colors["accent"])])

    # --- 选择控件 -----------------------------------------------------------
    for name in ("TCheckbutton", "TRadiobutton"):
        _safe(st, name, background=colors["surface"], foreground=colors["text"],
              focuscolor=colors["accent"], font=FONT_BODY, padding=(2, 3))
        _safe_map(st, name,
                  background=[("active", colors["surface"])],
                  foreground=[("disabled", colors["muted"]), ("active", colors["text"])],
                  indicatorcolor=[("selected", colors["accent"]),
                                  ("disabled", colors["border"])])

    # --- 页签：扁平分段，不做厚边框 -----------------------------------------
    _safe(st, "TNotebook", background=colors["bg"], bordercolor=colors["bg"],
          lightcolor=colors["bg"], darkcolor=colors["bg"], borderwidth=0,
          tabmargins=(0, 6, 0, 0))
    _safe(st, "TNotebook.Tab", background=colors["bg"], foreground=colors["muted"],
          bordercolor=colors["bg"], lightcolor=colors["bg"], darkcolor=colors["bg"],
          padding=(16, 8), font=FONT_SUBTITLE)
    _safe_map(st, "TNotebook.Tab",
              background=[("selected", colors["surface"]), ("active", colors["surface_alt"])],
              foreground=[("selected", colors["accent"]), ("active", colors["text"])],
              bordercolor=[("selected", colors["accent"]), ("active", colors["border"])],
              lightcolor=[("selected", colors["accent"])],
              darkcolor=[("selected", colors["accent"])])

    # --- 表格：行高 28px、斑马纹、无浮雕边框 --------------------------------
    _safe(st, "Treeview", background=colors["surface"], fieldbackground=colors["surface"],
          foreground=colors["text"], bordercolor=colors["border"],
          lightcolor=colors["border"], darkcolor=colors["border"],
          borderwidth=1, relief="flat", rowheight=ROW_HEIGHT, font=FONT_BODY)
    _safe_map(st, "Treeview",
              background=[("selected", colors["tree_select"])],
              foreground=[("selected", colors["text"])])
    _safe(st, "Treeview.Heading", background=colors["surface_alt"],
          foreground=colors["muted"], bordercolor=colors["border"],
          lightcolor=colors["border"], darkcolor=colors["border"],
          relief="flat", font=FONT_SMALL, padding=(6, 6))
    _safe_map(st, "Treeview.Heading", background=[("active", colors["border"])])

    # --- 进度条 / 滚动条（细滚动条：12px 宽，无浮雕箭头底）------------------
    _safe(st, "TProgressbar", background=colors["accent"], troughcolor=colors["surface_alt"],
          bordercolor=colors["surface_alt"], lightcolor=colors["accent"],
          darkcolor=colors["accent"], thickness=8)
    for name, orient in (("Thin.Vertical.TScrollbar", "vertical"),
                         ("Thin.Horizontal.TScrollbar", "horizontal")):
        _safe(st, name, background=colors["surface_alt"], troughcolor=colors["surface_alt"],
              bordercolor=colors["surface_alt"], lightcolor=colors["surface_alt"],
              darkcolor=colors["surface_alt"], arrowcolor=colors["muted"],
              arrowsize=10, gripcount=0, width=12 if orient == "vertical" else 12,
              relief="flat")
        _safe_map(st, name, background=[("active", colors["border"])],
                  arrowcolor=[("active", colors["accent"])])
    _safe(st, "TScrollbar", background=colors["surface_alt"], troughcolor=colors["surface_alt"],
          bordercolor=colors["border"], arrowcolor=colors["muted"],
          lightcolor=colors["surface_alt"], darkcolor=colors["surface_alt"])

    # --- 分隔线：1px 浅灰 ---------------------------------------------------
    _safe(st, "Line.TSeparator", background=colors["border"])
    _safe(st, "VSep.TSeparator", background=colors["border"])
    return colors


def register_tk_widgets(colors: Dict[str, str]) -> None:
    """保留接口：tk 原生控件在创建时直接使用色板，无需二次注册。"""
    return None
