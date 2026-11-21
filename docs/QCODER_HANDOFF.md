# Qcoder 项目交接

更新时间：2026-11-22

## 1. 当前状态

- 项目目录：`D:\Qcoder\打印机软件\pdf-duplex-print`
- 当前分支：`fix/p1007-offline-wedge-qcoder`
- 远端：`https://github.com/shenchuhannb/pdf-duplex-print.git`
- 最新代码修复：`fa6a663 fix: 只读下拉框选中后不再显示深色选区（qcoder）`，已推送；修复选择打印机或缩放方式后字段整块变深的问题。
- 2026-11-22 已将工作树回退到 2026-11-21 21:35（`961f851`）。日志分隔条快速拖动的残影属于已知症状；症状记录、已排除方向、根因判断和后续可选方案见 `docs/LOG_DRAG_SYMPTOM_ANALYSIS.md`。相关试验提交保留在历史中，未创建 PR。
- 不创建 Pull Request。
- 交接文档随下一次文档提交推送；最新提交以 `git log -3 --oneline` 为准。
- 当前完整测试基线：`123 passed`。
- 有一个名称显示为 `\340\261\260\355\230\207\307\211` 的未知未跟踪文件。不要读取、删除、修改或提交它。

后续接续先运行 `git status --short`、`git log -3 --oneline`，再针对具体反馈阅读对应模块。不要重新进行全项目审计。

## 2. 启动与排障入口

桌面快捷方式：`C:\Users\Lenovo\Desktop\PDF双面打印Qcoder.lnk`

快捷方式必须从项目源码和现有虚拟环境启动：

```text
目标：D:\Qcoder\打印机软件\pdf-duplex-print\.venv\Scripts\pythonw.exe
参数："D:\Qcoder\打印机软件\pdf-duplex-print\pdf_duplex_print.py"
```

常用命令：

```bat
cd /d D:\Qcoder\打印机软件\pdf-duplex-print
.venv\Scripts\python.exe pdf_duplex_print.py --self-check
.venv\Scripts\python.exe -m pytest -q
```

配置：`%APPDATA%\PDFDuplexPrint\settings.json`

运行日志：`%APPDATA%\PDFDuplexPrint\logs\app.log`

## 3. 本轮日志拖动修复

涉及文件：

- `ui_main.py`：监听工作区自身的 `<Configure>`，不再只依赖根窗口尺寸变化；按实际工作区高度调整左右分栏行高和预览画布最小高度。
- `ui_files.py`：上下布局的工作区过矮时，暂时隐藏“最近使用”和空状态次要说明，把空间留给文件行；拖回后自动恢复。
- `tests/test_ui_files.py`：新增日志分隔条拖动回归。

覆盖 `stacked`、`split`，以及 100%、150% 两档缩放。测试将日志从最小高度拖到窗口约三分之一，再拖回，断言：

- 文件列表至少保留表头和一行内容的高度；
- 预览画布不会越出父容器；
- 分栏不会侵入日志区域；
- 拖回后次要内容恢复；
- 新配置下界面变量为 `duplex`，`manual_duplex=True`。

透明截图实测，1024×720、日志 240px 时：

- `stacked`：文件表格 77px，预览画布 75px，工作区 202px；
- `split`：文件表格 85px，预览画布 68px，工作区 384px。

极端高度下预览会变小，这是用户主动放大日志后的正常降级；不得再通过固定最小高度把预览挤到日志后面。

## 4. 默认手动双面约定

普通默认配置已经是手动双面，无需再次改写：

- `settings_store.default_settings()["manual_duplex"] is True`；
- 默认打印流程为 `batch_all`；
- 主界面 `duplex_mode_var="duplex"`；
- 普通“恢复默认”也恢复手动双面。

两个例外必须保留：

1. 用户已有配置明确选择单面时，加载配置不得被默认值覆盖。
2. “恢复 Acrobat 默认”是显式特殊操作，仍恢复单面。

## 5. 近期已交付修复

| 提交 | 内容 |
|---|---|
| `fa6a663` | 只读下拉框的选区底色改为字段底色，选择后不再出现深色块 |
| `e69a5ae` | 标题栏左侧标志改为读取 `app_icon.ico`，与应用/任务栏图标统一 |
| `93cab1d` | 指标格补齐左右收边、数字 14pt、数字与名称留 3px 呼吸位 |
| `c1ea441` | 顶部四项批次指标改为浅灰底、细分隔的等宽统计格 |
| `c60762e` | 日志栏拖动时工作区、文件列表和预览的动态降级 |
| `9a0b151` | DOCX 字体表中的未使用字体不再造成缺失误报 |
| `21ad77f` | 窗口拉伸时提示文字、按钮间距和换行 |
| `9059999` | 转换期间新增文件自动续转，不再一直待转换 |
| `c41f2f5` | Word 子进程使用 UTF-8 协议，修复中文字体名乱码 |
| `b6f1073` | Tkinter/ttk 界面与响应式布局重构 |

顶部指标栏不要再改回“本批 + 内联数字/文字”：当前为 4 个 `_px(112)` 等宽格，整组左右各有 1px 收边线，格间 1px 分隔；数字使用 14pt `FONT_NUM` + 强调色，名称使用 `FONT_SMALL` + 辅助灰，数字与名称之间留 3px。表头完整高度 76、紧凑高度 58。`test_header_metrics_use_equal_cells_without_overlap` 覆盖 1024 宽与 150% 缩放，断言四格等宽、数字/名称完整、且不与右侧打印机状态重叠。

标题栏左侧标志、窗口图标、任务栏图标和桌面快捷方式统一使用根目录 `app_icon.ico`：`ui_main.py` 的 `_load_app_icon()` 按当前 DPI 选取最接近的 ICO 尺寸并用 LANCZOS 缩放，Pillow 不可用时回退到原来的 Canvas 双页绘制。不要再新增第二套标题栏图标资源。

只读下拉框选完后 Tk 会保持整段文字的全选态；`TCombobox` 的 `selectbackground` / `selectforeground` 必须与字段底色、文字色一致，否则会显示系统深色选区，看起来像控件背景突然变深。`test_combobox_selection_uses_field_background` 锁定 `focus + readonly` 状态。

此前还完成了 P1007 脱机预检、打印队列竞态与取消处理、单页文件跳过翻面、完成态和进度修复、后台切回黑块修复、逐源页缩放、四合一/八合一和横竖版入口。相关逻辑不要因 UI 调整被改写。

字体注意：`MS Mincho` 曾只存在于 DOCX 的 `fontTable.xml`，正文和样式并未使用，不是真实缺失。不得根据字体表条目要求用户安装字体，也不要购买或绕过字体授权。

## 6. 测试与 UI 验证硬约束

用户正在使用电脑进行其他活动，禁止可见前台测试：

- 不弹出可见 Tk 窗口，不抢焦点，不切换窗口；
- 不使用 `ImageGrab` 或前台截图；
- 使用单元测试、静态检查、隐藏或完全透明窗口、几何断言；
- 截图复用 `tools/capture_ui.py` 的 `PrintWindow(PW_RENDERFULLCONTENT)` 路径；
- 测试通过临时 `APPDATA` 隔离真实配置；
- 不触发真实打印、文件对话框或翻面提示。

提交前最小检查：

```bat
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe -m py_compile ui_main.py ui_files.py tests\test_ui_files.py
git diff --check
```

2026-11-21 验证结果：`123 passed`，上述语法和格式检查通过。真实纸张输出不在该结论范围内。

## 7. 真实打印边界与后续事项

HP LaserJet P1007 当前脱机。不要提交测试打印任务，不要让任务堆积，也不要用 Microsoft Print to PDF 的结果冒充真实出纸验证。

只有打印机在线且用户明确参与时，才进行实纸验证。建议按单页、偶数页、奇数页、多文件批量、多份逐份的顺序验证页序、补页、翻面和取消。打印前再次检查队列，测试后确认没有残留任务。

目前没有已知未修复的代码反馈。后续以用户新反馈为准，不主动扩展功能、不引入新的 UI 框架或运行时依赖，并保留快捷键、配置兼容、打印对话框时序、隐藏自动化支持、`stacked` 和 `split`。

## 8. 低余额接续策略

未查询、也不虚构积分余额。为减少消耗：优先使用本文和现有测试定位，只读相关函数；复用 `tools/capture_ui.py` 与现有测试 fixture；避免重复全量测试截图和重复探索已关闭问题。提交时只暂存明确目标文件，推送当前 qcoder 分支，不建 PR。
