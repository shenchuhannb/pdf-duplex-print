# CodeArts 全面整改报告（修复后）

日期：2026-11-22
对应修复前报告：`REPORT.md`

## 1. 验证总览

- 完整测试：**151 passed**（原 126 项 + 24 项审计回归 + 1 项批量弹窗作用域回归）。
- 系统 Python 与快捷方式虚拟环境均执行完整测试。
- PDF/配置与集成故障探针全部重新执行。
- 真实 Word 单文件隐藏转换：1 页，文本标记保留。
- 真实 Word 批量转换：2/2 成功，2 次回调。
- 冻结 EXE Word worker：真实转换成功，文本标记保留。
- 双实例日志：两个进程各 100 条消息全部保留。
- 正式 `tools/build_exe.bat`：构建、版本信息嵌入、冻结自检和 `SHA256SUMS.txt` 全部成功。
- 未提交真实打印任务，未进行可见前台测试。

## 2. S1 问题处理

| 项目 | 处理结果 | 主要改动 |
|---|---|---|
| Office PID 差集误杀 | 已修复 | 删除全局新增 PID 强杀；worker 仅登记自己 DispatchEx 创建且可取得 HWND 的实例，异常时才按登记 PID 清理；正常路径仍 Quit。历史性能脚本也不再强杀。 |
| 队列异常误判完成 | 已修复 | 持续 GetJob 异常进入 `PrintJobUnknownError`，不再返回 True；pStatus 与数值位并集；DELETED 纳入错误；ClosePrinter 不再掩盖原始结果。 |
| DEVMODE 静默回退 | 已修复 | OpenPrinter/DocumentProperties/驱动校验失败立即中止；回读检查 dmDuplex、dmCopies、dmNup 等关键字段，不再用默认 DEVMODE 创建 DC。 |
| 队列超时自动继续 | 已修复 | 不再自动固定延时继续；UNKNOWN/超时通过现有 UI 工作线程询问用户，默认停止，只有明确确认完整出纸才继续。 |
| PyMuPDF 许可证 | **需要维护者决定** | 已新增 `THIRD_PARTY_NOTICES.md`，但 AGPL 或商业许可必须根据实际发布方式选择，代码无法代替决定。 |

## 3. S2 问题处理

| 项目 | 处理结果 | 主要改动 |
|---|---|---|
| 非零 MediaBox/CropBox 裁切 | 已修复 | 以当前 CropBox 原点和尺寸映射 0/90/180/270 度显示坐标，裁切只缩小当前 CropBox。新增原点与旋转测试。 |
| 批注丢失 | 已修复 | `annots=True` 时用 PyMuPDF `bake()` 将批注/表单扁平化后再排版；False 不复制。像素测试确认可见内容保留。 |
| NaN/Infinity、布尔和枚举配置 | 已修复 | 非有限值回退默认；严格布尔解析；完整枚举/N-up 校验；损坏 JSON 隔离备份；保存 fsync 且 allow_nan=False。NaN 实际启动已恢复 `ok=True`。 |
| 多份张数、逐文件份数和状态 | 已修复 | 张数取实际页序长度；逐文件使用本行份数；不同份数双面自动逐文件、单面逐文件提交；完成标记不再被后续失败覆盖；排版失败文件也更新状态。 |
| 启动器与构建不可复现 | 已修复 | 缺依赖时才 pip；无 pip 先 ensurepip；检查 Python>=3.10；PyInstaller 锁 6.22.3；构建后强制冻结自检并生成 SHA-256。 |
| 批量 PIPE、半截 JSONL、取消和旧输出 | 已修复 | stdout/stderr 落盘；只消费完整 JSONL 行；iid 唯一；取消只清理本次改动输出；旧输出签名检查；单文件失败自动恢复旧 PDF。 |
| 输入资源和 Office 安全 | 已修复 | Word 强制禁用宏，WPS 尽力禁用；src==dst 拒绝；图片独立工作目录、原子替换、单帧/总像素/帧数/时间限制。 |
| 关键测试缺口 | 已补充 | 新增 `tests/test_audit_regressions.py`，覆盖 PDF、Office、队列、配置、Worker、图片和冻结命令协议。 |
| 入口异常可观测性 | 已修复 | 增加 sys.excepthook 与 Tk report_callback_exception 日志；自检结果原子写入，hold 上限 300 秒。 |

## 4. S3 问题处理

| 项目 | 处理结果 | 主要改动 |
|---|---|---|
| N-up 预览间距 | 已修复 | 预览与实际拼版统一使用 2 mm 间距，缩放测试一致。 |
| 非法 N-up 和页序防御 | 已修复 | 非法 N-up 直接报错；槽位不足不再 break 丢页；异常双面输入和逐文件奇数面数被拒绝，单内容面多份保持兼容。 |
| UI Mixin 隐式耦合 | 保持架构、局部降低风险 | 未做无关重写；保留现有 Mixin 与接口，仅明确任务快照、实际份数和完成状态，避免引入新的 UI 框架。 |
| 多实例配置/日志 | 已修复 | 配置写锁 + 基线差异合并；日志延迟打开 + 锁文件串行化写/滚动；锁文件不追加增长。双实例日志实测 200/200。 |
| ctypes 与图标资源 | 已修复 | 补齐 Win32 argtypes/restype、HRESULT 检查和日志；进程退出释放 HICON。 |
| 冻结目录与发布完整性 | 已处理可代码化部分 | 保留已验证 onedir，不盲删组件；加入版本信息、SHA-256 和第三方说明。Authenticode 需要外部证书。 |
| CI 与 Python 矩阵 | 已修复 | 新增 Windows CI，覆盖 Python 3.10–3.13、测试、探针、构建和冻结自检。实际 GitHub 运行结果需推送后确认。 |

## 5. 额外发现并修复

- 最大化黑块：`win32_flicker.py` 在 WM_SIZE 时先填充主题底色，Tk 完成布局后整树同步重绘；禁用 DWM 过渡动画。离屏最大化连续 8 帧黑像素比例为 0。
- 日志快速拖动残余：恢复窗口树时改为先累计整树失效、子 HWND 逆序恢复、顶层最后提交，最后调用 `DwmFlush()` 等待完整帧被合成。
- 冻结 worker 选择现在同时考虑目标 EXE，不再只检查父进程是否冻结；真实冻结 Word 转换已通过。
- 图片成功路径不再重复关闭 PyMuPDF 文档。
- 批量反面弹窗恢复读取快照份数，避免变量作用域错误。
- 队列/固定等待计时改用 monotonic。
- 暂停/恢复/取消队列失败会回滚或明确提示，不再假装状态成功。
- 打印任务 ID 在 StartDoc 成功后立即回调 UI，取消不再等整个 print_pages 返回。
- 字体提示按实际 BitmapMissingFonts 结果生成，不再误报位图化导出。
- 自检文件原子写入，配置/日志写锁文件不无限增长。

## 6. 仍需用户或外部环境补充

这些不是可以用代码凭空完成的事项：

1. **许可证选择**：项目实际分发方式、是否接受 AGPL，或是否购买 Artifex 商业许可。
2. **代码签名**：合法 Authenticode 证书及签名策略；SHA-256 不能替代签名。
3. **真实打印机**：在线设备、纸张和用户配合，验收页序、翻面、多份、暂停、取消、卡纸和队列残留。
4. **WPS**：当前环境未安装，无法验证 WPS COM。
5. **干净机器和第二用户**：用于无 Python 部署、ACL 和多用户验证。
6. **GitHub CI 运行**：工作流已创建，但需推送后查看 3.10–3.13 的实际矩阵结果。
7. **前台视觉复测**：日志拖动和空状态图标仍需用户最终确认。

## 7. 交付文件

业务修复涉及 `word_convert.py`、`image_convert.py`、`pdf_layout.py`、`printer_core.py`、`core_utils.py`、`settings_store.py`、`log_setup.py`、`ui_worker.py`、`ui_main.py`、`ui_settings.py`、`app_identity.py` 和 `pdf_duplex_print.py`。

部署与文档涉及 `run.bat`、`tools/build_exe.bat`、`requirements-build.txt`、`tools/version_info.txt`、`tools/write_checksums.py`、`.github/workflows/windows.yml` 和 `THIRD_PARTY_NOTICES.md`。

回归与证据位于 `tests/test_audit_regressions.py`、`tests/test_printer_core.py` 和 `docs/audit_evidence/`。`docs/word_compare/scripts/profile_convert.py` 仅移除危险的历史 Office 强杀调用。
