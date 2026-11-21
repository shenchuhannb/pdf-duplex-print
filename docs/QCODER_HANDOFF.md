# Qcoder 项目交接

更新时间：2026-11-21

## 1. 当前状态

- 项目目录：`D:\Qcoder\打印机软件\pdf-duplex-print`
- 当前分支：`fix/p1007-offline-wedge-qcoder`
- 远端：`https://github.com/shenchuhannb/pdf-duplex-print.git`
- 最新代码修复：`f09ba7a fix: 点击日志分隔条但不改变高度时跳过排版与整树重绘（qcoder）`，已推送；接续 `d5e1447` 的指示线拖动，纯点击不再触发任何排版或整树重绘。
- 不创建 Pull Request。
- 交接文档随下一次文档提交推送；最新提交以 `git log -3 --oneline` 为准。
- 当前完整测试基线：`125 passed`。
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

问题：单独拖动日志分隔条时，文件列表空状态、日志标题出现旧像素残影，日志框追赶鼠标抖动；窗口边框缩放却正常。原因是 Windows 下内部拖动只产生局部脏区，Tk 逐控件增量重绘来不及，快速拖动时留下未覆盖的旧像素；窗口边框缩放由系统整窗失效，所以干净。首轮“按帧合并 + 拖动中冻结排版”（`3891536`）没有消除，说明问题在绘制层而不是排版频率。

现实现（`d5e1447`）：采用 Tk 经典 PanedWindow 的 `opaqueresize=0` 策略，拖动期间不做任何布局。

- `ui_log.py`：按下分隔条时显示一条全宽强调色指示线（`log_drag_line`）；拖动只 `place` 这条线，日志高度、文件区密度、预览画布全部不动，因此拖动过程中没有会产生残影的重绘。松手时隐藏指示线，一次性落到目标高度，`after_idle` 调 `_finish_log_resize()` 解除冻结并同步一次密度/预览。
- 高度仍夹在 `_px(134)` 与窗口 1/3 之间；指示线位置按夹取后的高度计算，拖到极限时线停住，不会越界。点击分隔条但不改变高度时直接返回，不做排版也不做整树重绘。
- 拖动期间根窗口光标保持 `sb_v_double_arrow`，指针离开 6px 分隔条后仍有反馈。
- `_finish_log_resize()` 末尾调用 `win32_flicker.invalidate_tree()`（`RedrawWindow(RDW_INVALIDATE | RDW_ALLCHILDREN)`），松手后的单次排版整树重绘，与窗口边框缩放一致；拖动中不再有任何全树失效。
- 连续拖动或在落定前再次按下：先补完上一次松手的排版，再开始新拖动。
- `tests/test_ui_files.py`：`test_log_drag_defers_density_and_preview_until_release` 断言拖动中高度不变、密度/预览/整树失效均为 0、指示线可见且随拖动单调上移；松手后高度落定、指示线消失、排版与整树失效各至少一次；纯点击后高度与整树失效计数都不变。

覆盖 `stacked`、`split`，以及 100%、150% 两档缩放。原有拖动回归继续断言：

- 文件列表至少保留表头和一行内容的高度；
- 预览画布不会越出父容器；
- 分栏不会侵入日志区域；
- 拖回后次要内容恢复；
- 新配置下界面变量为 `duplex`，`manual_duplex=True`。

透明窗口离屏验证（1024×720，`stacked`、`split`，100%、150%）：拖动中日志高度与所有排版保持不变，指示线随拖动单调上移，1024 宽窗口实测指示线覆盖 1000px 全宽（PrintWindow 截图可测到强调色像素），松手后指示线消失、高度按夹取后的值一次落定；点击分隔条不拖动不改高度，连续两次拖动可从上限回到下限。指示线跟手程度无法在无可见窗口条件下验证，以用户前台观感为准。

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
| `d5e1447` | 日志拖动改为指示线预览，松手后一次性排版，消除快速拖动残影与日志框抖动 |
| `3891536` | 日志拖动按帧合并，拖动中冻结排版，松开后统一恢复（未彻底消除，已由 `d5e1447` 取代） |
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

2026-11-21 验证结果：`125 passed`，上述语法和格式检查通过。真实纸张输出不在该结论范围内。

## 7. 真实打印边界与后续事项

HP LaserJet P1007 当前脱机。不要提交测试打印任务，不要让任务堆积，也不要用 Microsoft Print to PDF 的结果冒充真实出纸验证。

只有打印机在线且用户明确参与时，才进行实纸验证。建议按单页、偶数页、奇数页、多文件批量、多份逐份的顺序验证页序、补页、翻面和取消。打印前再次检查队列，测试后确认没有残留任务。

目前日志拖动已改为指示线预览、松手落定，拖动时零排版所以不再有绘制残影；指示线跟手观感待用户前台确认。HP LaserJet P1007 仍脱机，真实出纸未验证。后续以用户新反馈为准，不主动扩展功能、不引入新的 UI 框架或运行时依赖，并保留快捷键、配置兼容、打印对话框时序、隐藏自动化支持、`stacked` 和 `split`。

## 8. 低余额接续策略

未查询、也不虚构积分余额。为减少消耗：优先使用本文和现有测试定位，只读相关函数；复用 `tools/capture_ui.py` 与现有测试 fixture；避免重复全量测试截图和重复探索已关闭问题。提交时只暂存明确目标文件，推送当前 qcoder 分支，不建 PR。
