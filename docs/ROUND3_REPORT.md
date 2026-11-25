# 第三轮用户反馈修复报告（WebView 版）

日期：2026-11-23　范围：`webapp/static/*`、`webapp/bridge.py`、`pdf_layout.py`、`tests/*`

## 1. 具体改动

### ① 添加文件默认选中
- `webapp/static/app.js`
  - `doAdd()`：成功添加后 `state.selected.add(rec.iid)`，再 `renderFiles()` / `refreshMetrics()` / `refreshRecent()`，并默认预览最后一个新增文件。
  - `files-added`（原生拖放）：同样把新增 iid 加入 `state.selected`，预览最后一个。
  - `file-update`（Office/图片转换完成）：只替换记录对象，**不清空选中**；若原 `previewFile` 为同一 iid，则更新对象并按当前页码重绘（不再跳回第 1 面）。

### ② 左侧选中操作按钮单横排
- `index.html`：文案缩短为「范围…」「份数…」「恢复全局」「应用到全部」。
- `style.css`：`.selection-buttons` 由 2×2 grid 改为 `display:flex; flex-wrap:nowrap;`；按钮 `height/min-height:28px; padding:3px 6px`。
- 1024 逻辑宽度实测：4 个按钮同一 top（307px），横向范围 22.7–248.9px，均在文件面板内且无横向滚动。

### ③ 单面计划误报“面数必须补为偶数”
- `webapp/bridge.py#compute_plan`：`manual_duplex=False` 时**不再调用** `plan_batch_summary`（其底层 `compute_merged_duplex_orders` 要求每个文件面数为偶数），直接构造：
  - `sheets = Σ(output_count × effective_copies)`、`flip_count = 0`、`note = 单面输出…`、`fronts/backs = []`。
- 结果：2 个 3 页文件单面打印 `compute_plan` 返回 ok，`sheets=6`、`flip_count=0`。

### ④ 横版与拼接一致性
- `pdf_layout.py`
  - `paper_size_pt()`：所有方向分支都归一化为「横版宽>高 / 竖版高>宽」；小册子在 `auto` 方向下强制横版纸张（对折后每面为竖页），显式竖版则保留并在 `describe_plan` 给出“建议横版”提示。
  - `ComposeResult.side_scales`：`_place_source()` 返回每页**实测**缩放百分比与溢出标记，`compose_layout_pdf()` 按输出面汇总；预览不再估算。
  - `preview_metrics()` 明确 `page_w/page_h` 必须是源页尺寸（非纸张尺寸）。
- `webapp/bridge.py`
  - `preview_page()`：`paper` 文案直接取**真实合成页面** `page.rect`；`scale_percent` 取排版实测值；新增 `paper_pt`/`landscape`/`overflow` 字段。
  - `plan_description()`：改用源 PDF 首面尺寸计算 metrics（原先把纸张尺寸当源页，横版 N-up 比例算错）。
- 横版实测（A4、源页 595×842）：

| 方向 | N-up | 双面 | plan_paper(pt) | compose 页面(pt) | 槽位(首个) | 面数 | 预览缩放 = 实测缩放 |
|---|---|---|---|---|---|---|---|
| landscape | 2 | true | 841.9×595.3 | 841.9×595.3 | 418.1×595.3（2 槽） | 4 | 70.24% = 70.24% |
| landscape | 2 | false | 841.9×595.3 | 841.9×595.3 | 418.1×595.3（2 槽） | 4 | 70.24% = 70.24% |
| landscape | 4 | true | 841.9×595.3 | 841.9×595.3 | 418.1×294.8（4 槽） | 2 | 49.52% = 49.52% |
| landscape | 4 | false | 841.9×595.3 | 841.9×595.3 | 418.1×294.8（4 槽） | 2 | 49.52% = 49.52% |

即 297×210 mm、宽>高；竖版为 210×297 mm；预览文案与合成页面尺寸逐点一致（误差 < 0.05pt）。

### ⑤ 圆形加号视觉
- `index.html`：30×30 蓝底圆按钮内 SVG 24×24，`M12 5v14M5 12h14`（viewBox 24 中臂长 14），`stroke="#fff"`、`stroke-width="3.4"`、`stroke-linecap="round"`。
- hover/active/focus-visible 样式保留（放大、按下、描边）。

### ⑥ 打印机栏宽度
- `style.css`：`#printer-select` 固定 `width:220px`、`min-width:190px`、`max-width:240px`（≤1100px 时 210px），`text-overflow:ellipsis`；`#printer-name` 同样省略。
- `app.js`：`option` 文本保持完整；`select.title`、`#printer-name` 的 `title` 始终写入全名。

## 2. 验证结果

- `node --check webapp/static/app.js`：**通过**
- 全量 `pytest`：**248 passed**（原 209）
  - 新增：横版 2/4 合 1 合成页与 `paper_pt`、预览 metrics 与实测缩放一致、方向交换、小册子横版约束、单面 3 页多文件 `compute_plan`、横版预览与合成一致、前端契约（单行按钮/加号/打印机宽度/默认选中）以及空状态操作条隐藏的静态回归。
- `WV_SELFTEST`（真实窗口，1024×680）：**PASS**，43 项 **false_keys = []**
  - 检查项：`layout`、`printer_width`、`plus_icon_style`、`empty_selection_bar_hidden`、`empty_state_fits`、`empty_add_button`、`auto_select_added`、`deselect_one_click`、`selection_buttons_single_row`、`preview_landscape` 等。
  - 布局数据：`printerSelectWidth=210`、`selectionBarHidden=true`、`emptyBoxFits=true`、`plusStroke=true`、`pageOverflow=false`。
- 未执行任何真实打印。

## 3. 验收回退修复：空状态圆形加号被裁剪/遮挡

- `app.js#renderFiles`：空文件分支 `selectionBar.hidden = true`，有文件分支 `hidden = false`（原有逻辑不变）。
- `style.css`：`.selection-bar[hidden]{display:none}`；新增 `@media (max-height:760px){ .empty{padding:8px;gap:2px} .empty-note{display:none} }`。
- `bridge.py#_selftest`：`empty_add_button` 改为同时校验 30×30、`BUTTON`、`onclick.name==='addFiles'`、`#selection-bar.hidden===true`、与 `#file-list` 视口垂直可见交集 ≥90%、`elementFromPoint(中心)` 命中按钮或其子节点；并把空状态断言提前到拖放投递之前，消除竞态。新增 `empty_selection_bar_hidden`、`empty_state_fits` 与运行期 `selection_buttons_single_row`。
- 1024×680 实测（客户端 1010×643）：
  - 按钮矩形 `[237.35, 221.67] – [267.35, 251.67]`，列表视口 `154.67 – 301.33`，可见 30px / 30px = **100%**（要求 ≥90%）
  - `elementFromPoint(252.35, 236.67)` → `path`（位于按钮内，`hitIsButton=true`）
  - `#selection-bar.hidden === true`；空状态盒 `170.67–291.33` 完整落在列表内，`scrollHeight <= clientHeight`（图标不截断）
- 截图：`docs/screenshots/round3_empty_1024.png`（1024×680 → 1536×1020 @150%）

## 4. 真实窗口截图（1024 逻辑宽度，150% DPI）

- `docs/screenshots/round3_empty_1024.png`（1024×680 空状态：圆形加号完整可见、图标不截断、操作条隐藏）
- `docs/screenshots/round3_added_1024.png`（添加 2 个文件后默认全选、预览最后一个、按钮单行）
- `docs/screenshots/round3_landscape_nup2_1024.png`（横版 2 合 1：297×210 mm、68%）
- `docs/screenshots/round3_landscape_nup4_1024.png`（横版 4 合 1：297×210 mm、48%）

截图脚本：`tools/capture_webapp_round3.py`（隔离 APPDATA，进程内 evaluate_js，Win32 PrintWindow；必要时前台 BitBlt 回退）。

## 5. 已知遗留（不在本次范围）

- Tk 版 `ui_settings.py` 内嵌小预览仍把纸张尺寸传给 `preview_metrics()`，与本次修复的 WebView 版同类问题，可后续统一。