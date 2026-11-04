# -*- coding: utf-8 -*-
"""界面共享常量（从 ui_main.py 机械拆出，供各 mixin 与 ui_main 统一引用）。"""

from window_geometry import DEFAULT_SIZE as WINDOW_DEFAULT_GEOMETRY_SIZE


APP_TITLE = "PDF 批量双面打印工具"
PREVIEW_MAX_THUMBS = 40   # 预览缩略图最多渲染页数，避免超大文档卡顿
PREVIEW_LONG_EDGE = 880   # 成纸预览渲染像素：按最长边渲染，Canvas 再按控件大小缩放
LIVE_REFRESH_DEBOUNCE_MS = 100  # 设置改动后的防抖时间：连续输入只在停顿后刷新一次
# 间距规范：4px 细间距 / 8px 常规间距 / 12px 卡片间距；控件内边距统一 12×10
GAP_S = 4                        # 同一组控件内部（如按钮与文字）的细间距
GAP = 8                          # 8px 基准网格：分区间距/内边距的倍数
GAP_CARD = 12                    # 卡片与卡片之间的间距
CARD_PAD = (10, 6)               # 卡片内边距：左右 10 / 上下 6（720 高度预算）
# 两套布局定义（运行路径只用默认一种；另一套可显式传入 layout_scheme 构建/截图）
LAYOUT_SCHEMES = ("stacked", "split")
DEFAULT_LAYOUT_SCHEME = "stacked"
# 窗口默认/最小尺寸（与 window_geometry 保持一致）
WINDOW_DEFAULT_SIZE = WINDOW_DEFAULT_GEOMETRY_SIZE

# 界面显示用中文标签（变量内部仍保存英文枚举值，便于配置兼容与测试）
SCALE_LABELS = {
    "shrink": "适合可打印区",
    "fit": "拉伸铺满纸张",
    "actual": "实际大小 100%",
    "custom": "自定义百分比",
    "paper": "按纸张缩放",
}
SCALE_LABEL_TO_VALUE = {label: value for value, label in SCALE_LABELS.items()}
RANGE_LABELS = {
    "all": "全部页面",
    "current": "当前页",
    "custom": "指定范围",
    "odd": "仅奇数页",
    "even": "仅偶数页",
    "recent": "上次范围",
}
RANGE_LABEL_TO_VALUE = {label: value for value, label in RANGE_LABELS.items()}
# 高级设置下拉：界面只显示中文，内部仍保存英文枚举值（配置与测试兼容）
PAGE_ORDER_LABELS = {"row": "从左到右", "column": "从上到下"}
COLOR_LABELS = {"color": "彩色", "gray": "灰度", "mono": "黑白"}
QUALITY_LABELS = {"draft": "草稿", "normal": "标准", "high": "高质量"}
WAIT_LABELS = {"queue": "监听打印队列（推荐）", "delay": "固定延时（秒）"}
