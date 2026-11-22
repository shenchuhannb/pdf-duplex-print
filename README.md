# PDF 批量手动双面打印

Windows 平台的 PDF 批量打印工具：用**普通单面打印机**实现**手动双面打印**，批量导入、自动补空白页、逐面提示翻纸，全程可选、可预览、可取消。

当前版本：**1.1.0**（2026-11-22） · [更新日志](CHANGELOG.md)

![主界面](docs/screenshots/ui_screenshot.png)

## 为什么需要它

很多办公打印机不支持自动双面，手动双面又容易出错：正面打完忘记翻面、纸张进纸方向不一致、总页数是奇数导致最后一张背面空白却错位。本工具把这一整套容易出错的流程固化成一条固定规则，用户只需要在两批之间把整叠纸翻一次面。

## 核心规则（不可更改的约定）

1. **相对页码**：以用户选定的打印范围为基准重新编号。例如打印 `3-7`，则第 3 页是"相对第 1 页"，第 1、3、5 相对页走第一面。
2. **手动双面顺序**：
   - 第一面：按相对页码**正序**打印**奇数页**（1、3、5…）；
   - 弹出提示，用户把整叠纸取出、**整体翻面**后放回；
   - 第二面：按**逆序**打印**偶数页**（…、6、4、2），使 p 与 p+1 落在同一张纸的正反面。
3. **奇数补页**：范围总页数为奇数时，末尾自动补 1 张**同尺寸、同方向**的空白页，保证每张纸都有正反两面，避免最后一张单面输出。

批量场景下，正反面顺序按**全局页序**计算（而不是逐文件各自倒序），这样多份文件一次打完时纸叠顺序不会乱。

## 功能

- **批量导入**：PDF、Word（`.doc`/`.docx`）、WPS、常见图片格式（自动转成临时 PDF）
- **拖拽添加**：支持从资源管理器拖入文件，自动去重、非支持类型提示；无 tkinterdnd2 时回退到 Win32 `WM_DROPFILES`；也可用"添加文件夹"整目录导入
- **打印控制**：打印机选择、页码范围（`1-5,8,10-12`）、份数、逐份/逐页、手动双面、单面
- **版式**：纸张与进纸盒、方向、缩放（适合可打印区/实际大小/百分比）、N-up 拼版、小册子、页边距、裁切
- **可视化预览**：合成"成纸"效果实时预览，正反面方块示意图（第一面正序、第二面逆序），补页单独标注，批量列表显示页数与补页
- **完成判定**：打印队列监听（推荐，精确）或固定延时（兼容虚拟打印机/PDF 输出）
- **健壮性**：加密/损坏 PDF 识别、页数越界、取消与暂停、队列超时兜底、临时文件清理（含崩溃残留）
- **其他**：设置持久化、最近文件、日志落盘（`%APPDATA%\PDFDuplexPrint\logs\`）、高 DPI 适配、拖拽与快捷键

## 运行环境

- Windows 10 / 11（x64）
- Python 3.10+
- 打印通过 Win32 GDI（依赖 pywin32），**不依赖** Adobe Reader 等外部阅读器
- Word/WPS 转换需要本机安装对应软件（无对应软件时仍可正常处理 PDF 与图片）

## 安装与运行

```bat
run.bat
```

启动器会自动检测并复用环境，顺序为：已保存的环境、项目目录 `.venv`、
`%LOCALAPPDATA%\PDFDuplexPrint\venv`、`%USERPROFILE%\.venvs\PDFDuplexPrint`
和 `%USERPROFILE%\Envs\PDFDuplexPrint`。找到后不会重复创建或安装新的 Python。
如果都不存在，才会在项目目录创建 `.venv`，并把路径保存到本地
`.pdf-duplex-env`（不会提交到 Git）。

也可以显式指定环境，或重置后让启动器重新自动检测：

```bat
run.bat "D:\PythonEnvs\PDFDuplexPrint"
run.bat --reset-env
```

如果需要指定创建 venv 所用的基础 Python，可设置 `PDF_DUPLEX_PYTHON` 环境变量。
Python 3.10+ 必须已安装；程序不会自动在其他磁盘安装 Python。

也可以从 [Releases](https://github.com/shenchuhannb/pdf-duplex-print/releases) 下载打包版，解压后双击 `PDFDuplexPrint.exe`，无需安装 Python；如果发布页暂时没有资产，可按下一节自行打包。请保留同目录的 `_internal` 文件夹。

## 打包为免安装 EXE

```bat
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
tools\build_exe.bat ".venv"
```

产物位于 `dist\PDFDuplexPrint\`，整个目录可拷贝分发。构建脚本会嵌入版本信息、自动执行冻结自检并生成 `SHA256SUMS.txt`。第三方许可证说明见 `THIRD_PARTY_NOTICES.md`。

## 测试

```bat
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest tests -q
```

覆盖页序推导、补页规则、范围解析、DEVMODE 强制单面、设置迁移与日志初始化等纯逻辑；其中"正反面配对"会遍历所有份数与逐份组合，断言每张纸的反面页码恒等于正面页码 + 1。当前基线为 151 项全部通过，另附审计探针与真实 Word 隐藏转换验证。

## 启动自检

```bat
.venv\Scripts\python.exe pdf_duplex_print.py --self-check
```

输出 `ok=True` 表示依赖、图标与 AppUserModelID 均正常。

## 项目结构

| 文件 | 职责 |
|---|---|
| `pdf_duplex_print.py` | 程序入口：参数解析、依赖检查、DPI 与图标初始化、隐藏自检 |
| `pdf_layout.py` | 页序推导、奇数补页、N-up/小册子拼版与预览几何 |
| `printer_core.py` | GDI 打印后端、DEVMODE、打印机状态与打印队列监听 |
| `settings_store.py` | 设置读写与旧配置迁移、最近文件、Acrobat 默认值 |
| `core_utils.py` | 自定义异常、页码范围解析、拖拽数据解析与单位换算 |
| `word_convert.py` | Word/WPS → 临时 PDF（COM，子进程超时保护） |
| `image_convert.py` | 图片 → 临时 PDF |
| `log_setup.py` | 应用日志与滚动落盘（`%APPDATA%\PDFDuplexPrint\logs`） |
| `app_identity.py` | AppUserModelID、窗口/任务栏图标与进程分组 |
| `theme.py` | 颜色、字体、ttk 样式与尺寸常量 |
| `window_geometry.py` | 窗口居中、位置记忆与多屏可见性修正 |
| `win32_flicker.py` | 顶层窗口背景擦除，消除从后台还原时的黑色方块 |
| `ui_main.py` | 主窗口骨架（mixin 组合）、响应式布局与快捷键 |
| `ui_files.py` | 文件列表、导入/拖拽、最近使用与批次统计 |
| `ui_settings.py` | 打印机、纸张、版式、缩放与高级设置面板 |
| `ui_preview.py` | 成纸预览、页序示意、缩放与翻页 |
| `ui_log.py` | 运行日志面板、分隔条拖动与日志导出 |
| `ui_worker.py` | 打印线程、进度与取消、翻面提示与队列等待 |
| `ui_helpers.py` | 界面共享常量：间距、标签映射、预览限制与布局枚举 |
| `run.bat` | 启动器：检测/复用 Python 环境并启动程序 |
| `requirements.txt` / `requirements-dev.txt` | 运行依赖 / 测试与开发依赖 |
| `app_icon.ico` | 应用、窗口、任务栏与快捷方式的统一图标 |
| `CHANGELOG.md` | 版本更新、测试基线与已知问题 |
| `README.md` | 功能、运行、打包与使用说明 |
| `.gitignore` | 虚拟环境、缓存、打包产物与个人文档的忽略规则 |
| `tests/` | 页序、补页、范围解析、DEVMODE、设置迁移、日志与 UI 几何回归 |
| `tools/` | 图标生成、离屏截图、快捷方式、打包与身份校验脚本 |
| `docs/word_compare/` | Word 打印链路对比实测（乱码率、速度）与可复现脚本 |
| `docs/` | 项目交接、日志拖动分析与界面截图 |

## 文档

- `CHANGELOG.md`：版本更新记录、已知问题与升级说明
- `docs/word_compare/report.md`：本工具链路与"Word 手动转 PDF 再打印"的实测对比（10 份文档 / 59 页）
- `docs/LOG_DRAG_SYMPTOM_ANALYSIS.md`：日志分隔条拖动残影的现象、根因分析与可选方案
- `tools/README.md`：图标生成、截图、快捷方式、身份校验脚本说明

## 已知限制

- 仅支持 Windows（依赖 GDI 与 spooler）
- 手动双面需要人工翻面一次；打印过程中请勿改变纸张方向或取出中间纸叠
- 虚拟打印机（如 Microsoft Print to PDF）无法监听真实出纸，建议改用固定延时判定
- 日志分隔条拖动已采用整帧同步重绘，避免子窗口新旧像素混杂；真实前台快速拖动的观感仍以本机验证为准，详见 `docs/LOG_DRAG_SYMPTOM_ANALYSIS.md`
