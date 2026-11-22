# PDF 批量手动双面打印：全面、可验证审计报告

> 本文件保留修复前审计证据。后续整改已完成，当前状态与剩余外部验证条件见 `REMEDIATION.md`。

- 审计日期：2026-11-22
- 审计基线：`fix/log-sash-redraw-CodeArts`，提交 `a470ef7cd26d33017fa71e665394113f23a65f58`
- 项目类型：Windows 单用户桌面应用，无服务端、数据库、账号体系或网络监听
- 审计方式：静态调用链审计 + 两套环境完整测试 + 隐藏 Tk 启动 + 内容级 PDF/图片探针 + 模拟 Win32 故障 + 实际 PyInstaller 构建与冻结产物自检
- 安全边界：未提交真实打印任务，未执行可能误伤 Office 的真实转换，未弹出可见测试窗口，未读取或导出凭据

## 0. 执行结论

### 0.1 直接回答

1. **底层选型不需要更换。** Python、Tkinter、PyMuPDF、Pillow、pywin32/Win32 GDI、Office COM 和 PyInstaller onedir 与当前 Windows 本地、小批量、重视 Word/PDF 保真的业务匹配。当前证据不足以支持 Qt、Electron、WebView、LibreOffice、云打印或数据库迁移。
2. **架构不需要整体重写，但需要局部调整关键边界。** PDF 页序/几何纯函数、主线程快照、后台工作线程和 UI 队列是合理设计。必须调整的是 Office 进程所有权、打印阶段状态机、DEVMODE 失败策略，以及逐文件状态/纸张统计。
3. **当前版本不是“所有打印功能已验证正常”。** 两套环境均为 `126 passed`，源码和冻结产物隐藏自检均成功；但批注、偏移 MediaBox 裁切、损坏配置、队列异常和 DEVMODE 回退存在已复现问题。
4. **最高优先级不是 UI，而是防数据损失和防错误出纸。** 首先停止 Office PID 差集强杀，随后让队列/驱动状态从 fail-open 改为“未知即暂停，不得进入下一阶段”。
5. **真实打印仍无法确认。** HP LaserJet P1007 当前脱机、队列任务数为 0；本次没有以 Microsoft Print to PDF 或 Adobe PDF 冒充实纸验证。

### 0.2 证据等级

- **F（已确认事实）**：有当前提交的测试、日志、探针输出或构建结果。
- **I（合理推测）**：代码机制成立，但需要真实设备、并发或特定外部条件才能观察最终后果。
- **U（未验证）**：当前环境或安全约束下没有取得结论，不作通过或失败推断。

严重程度：S1=可能造成用户数据损失、错误出纸或发布合规阻断；S2=重要功能错误、启动/部署可靠性或明显维护风险；S3=局部正确性、体验或长期维护改进。工作量为熟悉项目的 1 名开发者粗估，含自动化测试，不含等待真实设备/法务的时间。

---

## 1. 项目目标与核心业务流程

### 1.1 目标

项目目标是用普通单面打印机完成 PDF、Word/WPS 和图片的批量手动双面打印，提供范围、份数、N-up、小册子、缩放、裁切、预览、队列监听、暂停/取消和临时文件清理。核心不可变规则见 `README.md:13-22`：

- 页码以用户选择范围为基准转为相对页码。
- 第一面正序打印奇数相对页，第二面逆序打印偶数相对页。
- 奇数范围末尾补同尺寸空白页。
- 多文件按全局页序配对，而不是逐文件各自倒序。

### 1.2 主数据流（F）

```text
PDF / DOC(DOCX/WPS) / 图片
  ↓
ui_files.py：文件记录、状态、页数、逐文件范围与份数
  ↓（Office/图片先转换）
word_convert.py / image_convert.py → 每源文件临时 PDF
  ↓
ui_worker.py:_start_print_impl（ui_worker.py:293-465）
  主线程读取 Tk 变量并固化 jobs/settings/PrinterOptions 快照
  ↓
pdf_layout.py:compose_layout_pdf（985-1046）
  范围过滤 → 空白页过滤 → 补页 → N-up/小册子 → 输出面 PDF
  ↓
单面 / 整批双面 / 逐文件双面（ui_worker.py:674-869）
  ↓
printer_core.py：DEVMODE → CreateDCW → PyMuPDF 栅格化 → GDI StartDoc/Page
  ↓
Windows spooler → wait_for_print_job → 翻面/取纸提示 → 下一阶段
  ↓
ui_queue → 文件行状态、进度、日志、临时文件清理
```

辅助数据流：

- 设置：`%APPDATA%\PDFDuplexPrint\settings.json`，读取清洗、同目录临时文件 + `os.replace`，见 `settings_store.py:192-259`。
- 日志：`%APPDATA%\PDFDuplexPrint\logs\app.log`，1 MiB 滚动、3 个备份，见 `log_setup.py:14-46`。
- 预览：后台生成、按代次更新；主界面状态集中在 `ui_main.py:139-225`。
- 启动清理：`ui_main.py:204-212` 调用 `cleanup_stale_temp_dirs()`。

### 1.3 可靠性设计中的正面证据

- 主线程在 `ui_worker.py:374-437` 完成 Tk 变量读取和任务快照，避免工作线程直接访问 Tk。
- `pdf_layout.py` 将页序、补页和大部分几何计算做成可单测函数。
- 配置写入使用同目录临时文件和原子替换，`settings_store.py:245-256`。
- Office 转换具备子进程、超时、结果 PDF 校验和批量复用机制，`word_convert.py:763-955、1006-1076`。
- 单页位图有 6000 万像素降采样上限，`printer_core.py:338-354`。
- 打印前校验页序，避免 StartDoc 后才越界，`printer_core.py:949-976` 及 `tests/test_printer_core.py:54-68`。

---

## 2. 技术栈、依赖和部署

| 层 | 当前方案 | 证据 | 评价 |
|---|---|---|---|
| 语言 | Python 3.13.9（本次实测）；项目声明 3.10+ | `README.md:35-40`；两套测试环境版本输出 | 当前 3.13 可运行；3.10–3.12 未形成兼容矩阵 |
| GUI | Tkinter/ttk + 原生 Win32 窗口处理 | `ui_main.py`、`ui_log.py`、`win32_flicker.py` | 无额外 GUI 运行时；DPI/绘制细节维护成本较高 |
| PDF | PyMuPDF 1.28.0 | `requirements.txt:4`；`pip show` | 矢量排版和栅格打印均适用；许可证需确认 |
| 图像 | Pillow 12.0.0 | `requirements.txt:6`；`image_convert.py` | 与本地图片转换需求匹配 |
| Windows | pywin32 311 + ctypes | `requirements.txt:5`、`printer_core.py`、`app_identity.py` | 必需的平台集成；函数原型和状态处理需严格测试 |
| 拖拽 | tkinterdnd2 0.7.0 + Win32 WM_DROPFILES 回退 | `requirements.txt:7`、`README.md:27` | 快捷方式环境已安装；系统 Python 未安装时仍可回退 |
| Office | Word/WPS COM 子进程 | `word_convert.py:532-674、1006-1145` | 保真优先的合理选择，但进程所有权是最高风险 |
| 设置 | JSON + 手写 sanitizer | `settings_store.py` | 小规模单用户配置适用，无需数据库 |
| 测试 | pytest；系统 8.4.2，快捷方式环境 9.2.0 | `requirements-dev.txt:3`；实际输出 | 126 项约 10–11 秒，反馈快 |
| 打包 | PyInstaller 6.22.3，onedir | `tools/build_exe.bat:19-36`；实际构建日志 | 当前构建和自检成功，产物 96.07 MiB / 1155 文件 |

**许可证事实（F）**：PyMuPDF 包元数据为 `Dual Licensed - GNU AFFERO GPL 3.0 or Artifex Commercial License`。是否满足项目分发方式属于 U，公开或商业分发前必须确认 AGPL 义务或商业授权，不能仅凭“依赖可安装”认定合规。

### 2.1 部署路径

1. 源码启动：`run.bat:21-87` 发现/创建 venv，安装依赖，通过 `pythonw.exe` 异步启动。
2. 快捷方式：当前 CodeArts 快捷方式直接复用原项目 `.venv\Scripts\pythonw.exe`，启动当前克隆源码。
3. 冻结分发：`tools/build_exe.bat` 生成 `dist\PDFDuplexPrint\` onedir。
4. 隐藏诊断：`pdf_duplex_print.py:65-121` 提供 `--self-check` 和结果文件。

**实际构建（F）**：直接复用构建参数和已有 PyInstaller 6.22.3 完成当前提交构建，日志末尾为 `Building COLLECT ... completed successfully`、`Build complete!`。冻结产物随后生成自检文件：

```text
ok=True
style=A
aumid=PDFTools.PDFDuplexPrint.1.0
icon_exists=True
icon_size=40588
argv=--self-check
```

**部署缺口（F）**：快捷方式环境执行 `python -m pip` 返回 `No module named pip`，而 `run.bat:76-81` 每次启动都强制 pip install。因此该环境能被快捷方式复用，却不能被当前启动器直接复用。`tools/build_exe.bat:19` 还会安装未锁版本的 PyInstaller，且脚本自身不运行产物自检；本次是手动补做自检。

---

## 3. 模块边界与架构评估

### 3.1 合理且应保留的部分

- `core_utils.py`：通用异常、范围和输入解析，无内部依赖。
- `pdf_layout.py`：1046 行、38 个函数，核心页序和排版相对独立，45 项测试通过。
- `printer_core.py`：1151 行，负责 GDI、DEVMODE、队列和设备几何，29 项测试通过。
- `settings_store.py` / `log_setup.py`：规模小、职责明确。
- `ui_worker.py`：任务快照与后台执行分离，是正确方向。
- 测试按核心模块划分，完整执行仅约 10 秒，适合持续回归。

### 3.2 需要局部调整而非重写的部分

**UI 隐式耦合（F）**：`DuplexPrintApp` 在 `ui_main.py:139` 同时继承 5 个 Mixin，所有 Mixin 通过共享 `self` 属性协作。`ui_files.py` 1677 行、`ui_settings.py` 1564 行、`ui_worker.py` 1208 行、`ui_preview.py` 1085 行。多个 UI 文件重复导入同一批核心模块，是机械拆分留下的痕迹。

影响：属性由哪个 Mixin 创建、何时有效主要依赖约定，增加修改成本。但当前模块按文件/设置/预览/日志/工作线程划分仍可理解，**不构成引入大型 MVVM、依赖注入或新 GUI 框架的理由**。建议仅在触碰功能时减少跨 Mixin 属性、为 WorkerSnapshot 和文件记录建立 dataclass，不再继续机械拆文件。

**跨层状态语义（F）**：`wait_for_print_job()` 用 `True` 同时表达“确认完成”和“多次异常后无法区分”，`_print_phase()` 又把超时转为固定延时继续。这里需要明确的 `COMPLETED / UNKNOWN / FAILED / CANCELLED` 状态，而不是更换打印技术。

### 3.3 数据一致性

优点：设置原子替换、打印输入快照、页序不变量测试、多文件全局页序测试均存在。

问题：

- 最终状态只遍历 `prepared`，排版失败的当前文件可能停在“排版中”；其他已准备文件统一覆盖为失败/取消，见 `ui_worker.py:541-558、600-610`。
- 逐文件已完成行会被后续失败覆盖，`ui_worker.py:859` 对比 `607-610`。
- 份数进入页序算法，但纸张统计只使用 `item["sides"]`，见 `ui_worker.py:633-637、820-867`，多份时进度和取纸提示少算。
- 多实例配置/日志没有跨进程协调；单用户场景属于低到中等风险，不是引入数据库的理由。

### 3.4 安全性与权限

已确认边界：

- 无网络服务、远程 API、管理员提权、应用账号或角色系统；权限主要由当前 Windows 用户、文件 ACL 和打印队列 ACL 决定。
- 设置路径不可用时能包装为 `SettingsError`，探针输出 `wrapped=true`。
- 子进程命令使用参数列表，未见 `shell=True`。
- 日志和设置包含文件名、路径、打印机等本地元数据，继承每用户 APPDATA ACL；临时目录回退时更依赖系统配置。

风险：

- Office/PDF/图片解析不是低权限沙箱；恶意或损坏文件主要形成本地可用性风险。
- Word 打开前设置 `Visible=False` 和 `DisplayAlerts=0`，但未见 `AutomationSecurity`；隐藏打开不可信文档的宏行为未验证，应做禁用宏加固。
- 图片没有总帧数、总像素、累计内存或转换超时限制，`image_convert.py:65-116`。
- PyMuPDF AGPL/商业许可路径需发布前确认。

### 3.5 性能与资源

当前提交的简单文本样本实测（F）：

```text
35 页 → 手动双面输出 36 面
compose_layout_pdf 三次：0.0313s / 0.0274s / 0.0299s，中位 0.0299s
300 DPI 单页：0.0107s，2481×3508，RGB 24.9 MiB
```

该结果只代表本机、简单文本 PDF，不包括 Office、复杂图形、GDI、spooler 或机械出纸。默认 300 DPI 的单页内存对桌面应用可接受；6000 万像素上限约对应 172 MiB RGB 原始数据，尚未包含 Pillow 副本和 GDI/spooler 数据。

历史报告 `docs/word_compare/report.md:25-40` 显示 Word 冷启动明显较慢，批量复用后差距缩小。这支持保留批量复用方向；该报告日期为 2026-11-04，且部分默认设置与当前代码不同，不作为当前功能正确性的证明。

---

## 4. 实际执行的验证与结果

### 4.1 完整自动化（F）

| 环境/模块 | 命令或范围 | 结果 |
|---|---|---|
| 系统 Python 3.13.9 | `python -m pytest -q` | **126 passed in 10.19s** |
| 快捷方式 Python 3.13.9 | 原 `.venv\Scripts\python.exe -m pytest -q` | **126 passed in 10.74s** |
| PDF 排版 | `tests/test_pdf_layout.py` | **45 passed** |
| 打印核心 | `tests/test_printer_core.py` | **29 passed** |
| 配置/日志/工具 | 三个测试文件 | **37 passed** |
| UI/Word | `test_ui_files.py` + `test_word_convert.py` | **15 passed** |
| 源码隐藏启动 | `pdf_duplex_print.py --self-check`，隔离 APPDATA | `ok=True` |
| 冻结隐藏启动 | 当前 onedir 产物 `--self-check` | `ok=True`，进程退出 |

测试正常并不等于所有功能正常：现有真实 PDF 测试主要使用空白页；`ui_worker.py`、真实 Office、真实图片转换和真实打印缺少端到端覆盖。

### 4.2 可复现探针（F）

报告附带：

```bat
python docs\audit_evidence\pdf_config_probe.py
python docs\audit_evidence\integration_probe.py
```

探针只使用临时文件和模拟 Win32，不访问真实打印队列、不结束真实进程。

关键输出：

```json
{
  "normal_three_page_duplex": {"output_pages": 4, "selected_count": 3, "all_source_markers_preserved": true},
  "annotations_enabled": {"source_annots": 1, "output_annots": 0, "content_preserved": true},
  "offset_mediabox_crop": {
    "actual": [100.0, 100.0, 566.65, 813.65],
    "expected": [128.35, 128.35, 666.65, 913.65]
  },
  "nan_config": {"accepted_without_error": false, "exception": "ValueError"},
  "persistent_getjob_permission_error": {"returned": true, "getjob_calls": 4},
  "devmode_read_failure": {"create_dc_received_default_devmode": true, "exception_raised": false},
  "office_pid_delta_cleanup": {"new_pids_killed": [101, 102], "ownership_checked": false},
  "rgba_image": {"pages": 1, "leftover_frame_files": []},
  "multipage_tiff": {"pages": 2}
}
```

另将 `{"dpi": NaN}` 放入隔离 APPDATA 后执行真实启动，退出码为 1，调用栈为：

```text
pdf_duplex_print.py:167 DuplexPrintApp
ui_main.py:147 load_settings
settings_store.py:237 _sanitize
settings_store.py:215 int(number)
ValueError: cannot convert float NaN to integer
```

### 4.3 外部集成只读探测（F）

- `available_converters()` 返回 `["Word"]`，仅证明 Word COM 已注册。
- 发现 5 台打印设备；默认 Adobe PDF 状态就绪、任务数 0。
- HP LaserJet P1007：`offline=true`、`text=脱机`、`jobs=0`。
- 未执行真实 Office 转换：当前 PID 差集清理存在误杀风险，在用户可能同时使用电脑时不应冒险。
- 未提交任何打印任务。

---

## 5. 逐项功能状态矩阵

状态含义：正常=当前证据覆盖的功能路径通过；部分正常=主路径有证据但存在明确缺口；异常=已复现错误；无法验证=当前未获得有效证据。

| 功能 | 状态 | 证据与边界 |
|---|---|---|
| 源码入口与默认手动双面 | 正常 | `126 passed`；自检 `ok=True`；`settings_store.py:32-37` |
| PyInstaller onedir 构建与启动 | 正常 | 当前提交实际构建成功；冻结自检 `ok=True`；未做干净机器验证 |
| `run.bat` 环境管理 | 部分正常 | 路径发现逻辑存在；目标 venv 无 pip，与 `run.bat:76-81` 不兼容 |
| PDF 页码范围、奇偶页、越界处理 | 正常 | `resolve_range_pages`，`pdf_layout.py:360-390`；对应测试通过 |
| 单文件双面页序、补页 | 正常 | `pdf_layout.py:153-179、793-837`；配对不变量测试；三页内容标记全保留 |
| 多文件全局页序/按份顺序 | 正常（算法） | `pdf_layout.py:190-308`；全局配对测试通过；实纸另列无法验证 |
| 标准零原点 PDF 排版 | 正常 | 三页带唯一文本探针输出 4 面且所有标记保留 |
| 加密 PDF 与页数边界 | 正常（已有样本） | `tests/test_pdf_layout.py:239-318` 覆盖加密拒绝和无效页数；未穷举所有损坏形式 |
| 批注输出 | **异常（文本批注已复现）** | `annots=True` 时源 1 条、输出 0 条；`pdf_layout.py:928-982` 未传递该选项；表单/戳记未逐一动态验证 |
| 裁切 | **异常（偏移页面）** | 非零 MediaBox 实际矩形与期望不符；`pdf_layout.py:875-913` 未加原点偏移 |
| N-up / 小册子算法 | 部分正常 | 页序、网格和补页测试通过；复杂内容、实际纸张顺序未验证 |
| N-up 成纸预览 | 部分正常 | 预览 50.023% vs 实际一致路径 49.547%；`preview_metrics` 漏 2 mm 间距 |
| PNG 透明图片转 PDF | 正常 | 实际生成 1 页，临时帧清理为空；`image_convert.py` |
| 多页 TIFF | 正常 | 实际生成 2 页 |
| GIF/WebP/EXIF/超大图片 | 无法验证 | 有实现但本次未逐格式运行；无总资源上限 |
| Word COM 当前版本导出 | 部分正常 | Word 已注册；历史报告可参考，但当前未安全执行真实转换，且有清理缺陷 |
| WPS 转换 | 无法验证 | 当前 `available_converters()` 不含 WPS |
| 转换期间新增文件续转 | 正常 | `test_file_added_during_conversion_is_processed` 通过 |
| Office 超时/取消/进程清理 | **异常** | PID 差集可强杀所有新 Office PID；未验证真实文档损失，因风险未做危险复现 |
| GDI 参数映射与页序校验 | 正常（模拟） | DEVMODE 字段和假 GDI 共 29 项测试通过；真实驱动接受情况未验证 |
| DEVMODE 强制单面/单份/单版 | **异常（失败回退）** | 模拟 OpenPrinter 失败后仍以默认 DEVMODE 创建 DC，无异常；`printer_core.py:897-946` |
| 队列正常状态监听 | 部分正常 | 状态解析/模拟测试通过；真实出纸未验证 |
| 队列异常处理 | **异常** | 已观察任务后连续权限异常返回 `True`；`printer_core.py:754-782` |
| 队列超时后的双面流程 | 部分正常/高风险 | 超时仅等待固定延时便继续，`ui_worker.py:1084-1099`；真实交叠未验证 |
| 脱机预检 | 部分正常 | 模拟脱机/超时测试通过；真实 P1007 被只读识别为脱机，未提交任务 |
| 暂停/取消 | 部分正常 | stop/pause 和模拟路径存在；GDI 阻塞、队列 ACL、设备缓冲下效果未验证 |
| 多份打印页序 | 正常（算法） | 配对不变量测试通过 |
| 多份纸张统计和取纸提示 | **异常** | `ui_worker.py:633、826、850、866` 未乘份数 |
| 失败/取消后的逐文件状态 | **异常** | `ui_worker.py:607-610` 统一覆盖 prepared 行，可能抹掉先前完成态 |
| 设置保存/迁移/Unicode | 正常（常规路径） | 原子替换和往返测试通过 |
| 损坏配置恢复 | **异常** | `NaN` 使实际启动退出码 1；`settings_store.py:207-237` |
| 最近文件/范围 | 正常 | 去重、置顶、截断测试通过 |
| 日志 UTF-8 落盘 | 正常（单进程常规） | 日志测试通过；多进程滚动未验证 |
| stacked/split 与 100%/150% DPI | 正常（隐藏几何） | UI 测试通过；不代替真实观感 |
| 日志快速拖动 | 部分正常 | 用户确认首版缓解但仍可复现；子窗口树后续版已测试，尚无用户复测结论 |
| 空状态圆形加号 | 正常（像素回归） | 100%/150% 对称与边缘测试通过；真实视觉复测未反馈 |
| 资源管理器拖拽 | 无法验证 | 快捷方式环境有 tkinterdnd2，另有 Win32 回退；未进行真实 Explorer 拖放 |
| 保留排版 PDF/自动打开目录 | 部分正常 | 代码路径明确；未在完整打印后验证产物和打开行为 |
| 实纸页序、翻面、多份、取消 | **无法验证** | P1007 脱机且未提交任务，不能用虚拟打印机替代 |
| 干净 Windows 机器部署 | 无法验证 | 本机冻结自检成功，不等于无 Python/无开发环境机器验证 |
| 跨用户 ACL/权限隔离 | 无法验证 | 未使用第二 Windows 账号；应用本身无账号/角色体系 |

---

## 6. 分级问题清单与整改建议

### S1：优先处理，建议在下次正式发布前关闭

| 编号 | 问题与证据 | 影响 | 建议 | 工作量/优先级 |
|---|---|---|---|---|
| S1-1 | **Office PID 差集强杀**，`word_convert.py:960-1003、1020-1033`；探针 kills `[101,102]` | 转换期间用户新开的 Word/WPS 也可能被 `/F` 结束，存在未保存数据损失风险 | 移除全局进程名差集强杀；由 worker 记录自建 COM 实例并正常 Quit；需要强杀时使用明确进程所有权/Windows Job Object，补用户并发 Office 测试 | 3–5 人日，P0 |
| S1-2 | **队列异常被当完成**，`printer_core.py:754-782`；持续 Access Denied 实测返回 True | 可能提前翻面、提交反面或显示完成，造成页序错误 | 队列状态改为明确枚举；异常/消失只能进入 UNKNOWN 并暂停，要求用户核查，不能自动进入下一阶段；处理 DELETED 和 pStatus | 2–3 人日，P0 |
| S1-3 | **DEVMODE 失败静默使用默认设置**，`printer_core.py:897-946`；探针确认默认指针 + 无异常 | 驱动默认自动双面、多份或 N-up 时会破坏手动页序 | 强制字段必须验证成功，失败即中止并提示；读取修改后的 DEVMODE 核对 dmDuplex/dmCopies/dmNup | 1–2 人日，P0 |
| S1-4 | **队列超时后自动继续**，`ui_worker.py:1084-1099` | 原正面恢复后可能与反面交叠；完成状态与实际设备不一致 | 超时默认暂停而不是继续；只有用户明确选择且确认队列已清理时才允许固定延时模式 | 1–2 人日，P0 |
| S1-5 | **PyMuPDF 许可证路径未确认**，包元数据 AGPL/商业双许可 | 公开或商业分发可能不合规 | 确认项目分发许可；履行 AGPL 或取得商业授权，保留依赖许可证清单 | 0.5–2 人日 + 法务，P0（发布门槛） |

### S2：重要正确性、可靠性和安全加固

| 编号 | 问题与证据 | 影响 | 建议 | 工作量/优先级 |
|---|---|---|---|---|
| S2-1 | **非零 MediaBox/CropBox 原点裁切错误**，`pdf_layout.py:875-913`；探针实际/期望矩形不同 | 左/上少裁、右/下多裁，输出内容错误 | 在 0/90/180/270 映射中统一加入 MediaBox 原点；增加偏移、旋转、CropBox 组合内容测试 | 1–2 人日，P1 |
| S2-2 | **`annots` 在排版层未接线**，`pdf_layout.py:430、928-982` | 用户开启后仍丢批注；后续栅格 annots 无法找回 | 明确“保留交互批注”还是“扁平化可见批注”；N-up 优先采用可验证的扁平化路径，并测试文本/戳记/表单 | 1–3 人日，P1 |
| S2-3 | **NaN/Infinity 和宽松配置清洗**，`settings_store.py:192-237`；实际启动退出 1 | 单个损坏值可阻断启动；`"false"` 变 True，非法枚举可进入下游 | `math.isfinite`、严格布尔解析、完整枚举校验；load 总兜底并保留损坏文件备份和日志 | 0.5–1 人日，P1 |
| S2-4 | **多份张数少算、逐文件状态被覆盖**，`ui_worker.py:600-637、812-869` | 用户取错纸叠，历史完成状态失真 | 张数统一使用订单长度/2；文件状态使用逐文件状态机，失败文件即使未 prepared 也更新 | 1–2 人日，P1 |
| S2-5 | **启动器每次 pip、当前 venv 无 pip；构建工具未锁版本** | 启动器失败、联网依赖和构建不可复现 | 先探测依赖再安装；无 pip 时给明确修复路径或 ensurepip；锁 PyInstaller；构建后自动运行冻结自检 | 1–2 人日，P1 |
| S2-6 | **批量子进程 PIPE 不实时消费、半截 JSONL 被跳过**，`word_convert.py:889-935` | 大量输出可阻塞；部分写入会误重试 | 持续读取 stdout/stderr 或重定向文件；JSONL 只消费完整行，校验唯一 iid；取消时清理半成品 | 2–3 人日，P1 |
| S2-7 | **输入资源和 Office 安全加固不足**，`image_convert.py:65-116`、`word_convert.py:563-620` | 巨型图片/多帧 TIFF 可耗尽资源；不可信 Office 文档行为不明确 | 总像素/帧数/超时/输出大小限制；禁止 src==dst；设置 AutomationSecurity 禁用宏；保持错误可观测 | 2–4 人日，P1 |
| S2-8 | **关键路径测试缺口**：ui_worker、真实图片/Word、冻结 worker、真实驱动 | 126 passed 对最高风险链路覆盖不足 | 优先补 Worker 状态/份数、图片格式、转换超时/取消、队列异常、冻结 worker 协议；建立真实设备手动验收单 | 3–5 人日，P1 |
| S2-9 | **入口异常可观测性不足**，`pdf_duplex_print.py:124-176` | pythonw 下启动异常可能只表现为无窗口 | 安装 sys.excepthook/Tk reportcallbackexception 统一日志；启动器提供诊断模式和明确退出状态 | 1–2 人日，P1 |

### S3：局部修正与长期维护

| 编号 | 问题 | 建议 | 工作量/优先级 |
|---|---|---|---|
| S3-1 | `preview_metrics` 未使用实际 2 mm N-up 间距，`pdf_layout.py:641` 对比 `1020` | 共用同一个槽位计算入口，锁定预览/输出几何一致 | 0.5 人日，P2 |
| S3-2 | 非法 `nup` 直接调用底层会丢页，`pdf_layout.py:499-501、714-720、1021-1023`；UI 有回退保护 | 排版入口拒绝未支持 N-up，不静默 break；不要误判为当前 UI 可触发 | 0.5 人日，P2 |
| S3-3 | 大 UI Mixin 隐式共享 self、重复导入 | 触碰功能时抽取 dataclass/明确接口，不做全量重写 | 持续 3–8 人日，P2 |
| S3-4 | 多实例配置最后写入者覆盖、日志滚动无跨进程协调 | 单实例锁或配置合并/锁；多实例测试 | 1–2 人日，P2 |
| S3-5 | 部分 ctypes 声明和失败返回值检查不完整，`app_identity.py:35-48、84-104` | 补 argtypes/restype、HRESULT 检查和必要日志 | 0.5–1 人日，P2 |
| S3-6 | 冻结目录 96.07 MiB / 1155 文件，构建未锁 PyInstaller、无签名/校验和 | 先保留 onedir；稳定后再分析 collect-all 裁剪、签名和哈希，不为体积贸然换打包器 | 1–3 人日，P3 |
| S3-7 | 无 `.github` CI，Python 仅实测 3.13 | 增加 Windows CI、3.10–3.13 矩阵和构建自检；真实打印机保留独立验收 | 1–2 人日，P2 |

---

## 7. 合理推测与未验证事项

### 7.1 合理推测，未作为已发生事故事实

- **I：实际双面页序可能受 DEVMODE 回退影响。** 已证明失败时使用默认值，但具体驱动是否默认自动双面/多份未在 P1007 验证。
- **I：Office 清理可能造成未保存内容损失。** 已证明所有新 PID 都可被强杀，但为保护用户数据，没有真实打开文档复现损失。
- **I：队列消失可能是完成，也可能是删除/权限故障。** 当前实现无法区分，真实设备上的发生频率未知。
- **I：批量 PIPE 可能在输出足够大时阻塞。** 代码未持续读取，但未构造大规模 Office 输出压测。
- **I：巨型图片可能耗尽内存。** 未为了不中断用户工作而执行破坏性压力测试。
- **I：多实例设置可能最后写入者覆盖。** 未见锁，但未执行双实例并发写入。
- **I：隐藏 Office 打开不可信文档的宏风险需要加固。** 未验证具体 Office 安全策略，不声称已发生宏执行。

### 7.2 明确未验证

- HP P1007 或其他物理设备的实际出纸、页序、翻纸方向、多份、暂停、取消、卡纸和缺纸。
- 当前提交的真实 DOC/DOCX 导出、WPS、字体替换和 Office 首次运行/受保护视图。
- 真实资源管理器拖拽和前台 UI 观感；日志子窗口树修复尚待用户复测。
- 无 Python、无开发工具、无 Office 的干净 Windows 机器部署。
- 第二 Windows 用户的 ACL、共享机器、多实例并发。
- Python 3.10/3.11/3.12、32 位环境和不同 Office 版本。
- 团队能力、发布许可策略和目标分发规模；因此维护成本评价以“小型 Windows 桌面项目、团队能维护 Python/Win32”为前提。

---

## 8. 建议整改顺序

1. **立即停止危险清理和错误推进（P0）**：S1-1～S1-4。先让故障停在安全状态，再追求自动恢复。
2. **确认发布合规（P0）**：S1-5，与代码修复并行。
3. **补最高价值回归（P1）**：Office 所有权、队列 UNKNOWN、DEVMODE 失败、多份统计、配置 NaN 必须先有失败测试再修。
4. **修 PDF 内容正确性（P1）**：S2-1、S2-2，增加真实内容、批注和偏移/旋转坐标样本。
5. **修 Worker 状态与转换可靠性（P1）**：S2-4、S2-6、S2-9。
6. **加固输入与部署（P1）**：S2-3、S2-5、S2-7，构建后强制冻结自检。
7. **真实设备验收**：打印机在线且用户明确参与时，按单页、偶数、奇数、多文件、多份、取消顺序验证，并确认队列无残留。
8. **最后处理体验和长期维护（P2/P3）**：预览一致性、底层参数校验、CI、UI 接口收敛、体积和签名。

**不建议的整改方式**：整体换 GUI、重写打印栈、引入数据库/服务端、以 LibreOffice 替代 Word、用虚拟打印机结果宣布实纸通过，或为了提高测试数而复制大量低价值 mock。

---

## 9. 最终评估

当前项目具备可运行、可构建、核心算法测试反馈快的基础，技术路线与业务匹配，**保留底层选型和总体架构是维护成本最低的决策**。但“126 passed”主要证明核心纯逻辑、配置和部分 UI/Win32 模拟路径，不能抵消已复现的跨层故障。

发布成熟度判断：**源码/冻结隐藏启动正常；标准 PDF 和图片基本路径可用；真实生产打印尚不具备充分验收证据，S1 问题关闭前不建议作为稳定正式发布。** 最优路线是定点修复外部进程、打印状态机和 PDF 坐标/批注问题，再用真实 Office 和真实打印机完成分层验收，而不是更换技术栈。
