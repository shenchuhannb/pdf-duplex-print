# 第四轮 UI/交互修复报告（任务 19）

日期：2026-11-23

## 实现位置（逐项）

### ① 圆形加号 hover 抖动
- `webapp/static/style.css` `.empty-add`：`transition: background .15s, box-shadow .15s`（去掉 transform）；
  `.empty-add:hover` 仅改背景 + 加深阴影，`.empty-add:active` 仅改背景 + 内收阴影，删除 `transform: scale()`。
- 命中区域固定 30×30，鼠标在边缘进出不再因缩放反复抖动。

### ② 禁止原生提示：应用内确认模态框
- `webapp/static/index.html`：新增 `#modal-confirm`（标题 `#confirm-title`、正文 `#confirm-body`、`#btn-confirm-cancel` / `#btn-confirm-ok`）。
- `webapp/static/app.js`：新增 `confirmAction({title, lines, confirmLabel, danger})`（Promise 式）与 `window.answerConfirm(value)`；
  4 处调用全部替换：
  - `cancelPrint()`（取消打印，async，调用处 `btn-cancel` 不变）
  - `btn-remove` 多选移除
  - `btn-clear` 清空列表
  - `btn-restore-defaults` 恢复默认
- `Esc` 优先关闭确认框并返回 false（`confirmAnswer` 置空，避免二次 resolve）。
- `webapp/static` 已无 `window.confirm` / `alert(` / `prompt(`。

### ③ 选中操作按钮等宽铺开
- `webapp/static/style.css`：`.selection-buttons { width:100%; flex-wrap:nowrap }`，`.selection-buttons .btn.sm { flex:1 1 0; min-width:0 }`。
- 1024 下 4 个按钮等宽（实测宽度差 ≤1px）且铺满条宽，单行、文字不截断。

### ④ 输出单选项顺序
- `webapp/static/index.html`：改为「☑ 手动双面」在前、「☐ 单面」在后，默认仍为手动双面。

### ⑤ 页面范围文案不更新
- `webapp/bridge.py`
  - `_range_label(mode, text, current_page)`：全部/奇数页/偶数页/自定义文本/当前页（第 N 页）。
  - `_record_pages()` 返回 `range_mode/range_text/range_label`。
  - `compute_plan()` 为每个文件写入 `effective_range`、`selected_pages`、`effective_copies`、`plan_text`，并把 `rows` 随响应返回（前端 `state.files` 是副本，必须回传才能刷新）。
- `webapp/static/app.js`
  - `refreshMetrics()` 合并 `res.rows` → `renderFiles()`。
  - 文件行范围取 `f.effective_range || f.range_text || "全部"`；逐文件自定义范围优先。

### ⑥ 打印机脱机文案与中文日志
- `webapp/bridge.py`
  - 新增 `_describe_error()`：去掉 `TypeError:` 之类 Python 异常类名前缀，空消息给可操作兜底。
  - `printer_status()` 失败统一 `status.state = "脱机"`，`detail` 为可操作中文。
  - 转换失败日志改为「文件转换失败 X：…」，堆栈只写日志文件（`get_logger().exception`）。
  - 打印失败汇总不再输出“未知错误”。
- `webapp/static/app.js`
  - 打印机标签 `${name}（脱机）`，并按 `detail` 记录可操作日志（detail 已含打印机名时不重复前缀）。
  - 日志统一：「无法生成打印计划：…」「无法恢复默认设置：…」「无法刷新打印机状态：…」「无法获取打印机列表：…」「无法读取页序详情：…」；`errorText()` 去除异常类名前缀。

## 验证

- `node --check webapp/static/app.js`：**通过**
- `python -m pytest tests -q`：**255 passed**
  - 新增/更新：无原生弹窗静态契约、确认模态框结构、按钮等宽 CSS、加号无 transform、输出顺序、`effective_range`（全部/奇数页/偶数页/自定义/当前页/逐文件优先）、`rows` 回传、打印机脱机 state+detail。
- `WV_SELFTEST`（真实窗口 1024×680）：**PASS**，**55 项全 true，false_keys = []**
  - 新增：`duplex_option_order`、`selection_buttons_equal_width`、`range_label_default_all`、`range_label_follows_odd`、`range_label_follows_custom`、`range_label_back_to_all`、`printer_offline_text`、`confirm_modal_clear`、`confirm_escape_cancels`、`confirm_modal_restore`、`restore_needs_dirty_state`、`no_native_confirm`。
- 正式版进程 PID 42716：检查时**已不在运行**（无需关闭）；测试使用隔离 APPDATA 的临时实例，结束后未重启应用。
- 未执行真实打印。

## 证据

- 确认模态框真实窗口截图：`docs/screenshots/round4_confirm_modal_1024.png`（1024×680 @150% → 1536×1020）
- 日志：`.verify_round4_pytest.log`、`.verify_round4_selftest.log`