# 输入法跳变问题 —— 交接文档（Handoff）

- 项目：`pdf-duplex-print`（PDF 批量双面打印工具）
- 交接日期：2026-11-24
- 当前状态：**已修复并实测验证**（2026-11-24 审计）；应用侧单次 Shift 补偿已落地，InputTip 保持原始配置
- 目标读者：后续维护者
- 权威结论：见 `docs/IME_SWITCH_ANALYSIS.md`（本次审计重写版；原分析文档关于"已修复"和 IMM32 状态的结论此前有误）

---

## 1. 任务目标

修好下面这个现象，且不破坏现有功能：

> 双击桌面 `PDF双面打印CodeArts.lnk` 打开应用后，微信输入法从中文跳成英文，切到其他应用后仍保持英文，必须手动按 `Shift` 才能切回中文。

验收标准见第 9 节。

---

## 2. 当前代码状态

| 项 | 状态 |
|---|---|
| `webapp/bridge.py` | 已包含 `_restore_ime()` + `_is_own_window_visible()` + `_send_single_shift()`；延迟 0.4s，窗口可见且未最小化时发送一次；TSF 预热/聚焦逻辑不存在 |
| `tests/test_webapp_bridge.py` | `attach` 契约断言为 `["_drop_pump", "_apply_window_icon", "_restore_ime"]`，另有 2 项补偿行为单测 |
| `pytest tests -q` | **297 passed**（在测试进程不注入 Agent 安全 shim 时；带 shim 时 296 passed + 1 项环境冲突失败） |
| `WV_SELFTEST` | **PASS**（修复后最近一次） |
| InputTip | `D:\InputTip\src\data\config.ini` 已恢复原始内容（仅 `[Settings]`），无规则、无自定义模式 |
| 运行入口 | `pythonw.exe webapp\main.py`（pywebview 6.2.1 + WebView2） |

下面的"已排除方案"是历史实验记录，相关实验代码已回退；当前生效的是第 7 节方向 A 的落地实现（单次 Shift 补偿）。

---

## 3. 环境基线（实测）

| 项 | 值 |
|---|---|
| Windows | `10.0.26300` |
| WebView2 Runtime | `157.0.4341.60` |
| 输入法 | **WeType 3.0.2.50**（微信输入法，纯 TSF TIP：`wetype_tip_x64.dll`） |
| 输入法进程 | `wetype_server` / `wetype_service` / `wetype_renderer` |
| TIP 注册 | `HKLM\SOFTWARE\Microsoft\CTF\TIP\{86598FB9-66A2-463E-B9C2-AEB906D477AD}`，LanguageProfile `0x00000804\{607FDF85-FCC8-4DBD-A365-41296F980C9C}`，`Enable=1` |
| 键盘布局 | `HKCU\Keyboard Layout\Preload` 仅 `00000804`，**无 en-US** |
| 每应用输入法 | `HKCU\...\Explorer\Advanced\EnablePerAppSettings` **未设置** |
| 应用是否管理员 | 否（`.lnk` `LinkFlags=0x000000F7`，`RunAsUser=False`） |
| 环境干扰 | `lxwp.exe`（LxWallpaper）的 `BkShadowWndClass` 会长期占据 `GetForegroundWindow` |

---

## 4. 复现步骤

1. 确认微信输入法当前是中文状态（状态条显示“中”）
2. 双击 `C:\Users\Lenovo\Desktop\PDF双面打印CodeArts.lnk`
3. 不点任何地方，直接敲 `nihao`
4. 现象：输入英文 `nihao`，无候选窗；状态条为“英”
5. 切到记事本敲 `nihao`：仍是英文，按一次 `Shift` 才恢复中文

---

## 5. 已确认的事实（有数据）

### 5.1 应用窗口链没有 IME 上下文

采样脚本：`.resultverify/ime_20261123/probe_ime_ab.py`（每 250ms 读 `ImmGetContext`）

| 组 | 应用窗口采样 | 窗口链 | 有 HIMC | HKL |
|---|---|---|---|---|
| 基线 | 31 | 3 / 5 / 6 | **0** | `0x08040804` |
| TSF 预热 | 31 | 3 / 5 / 6 | **0** | `0x08040804` |
| 聚焦 `#copies` | 34 | 3 / 5 / 6 | **0** | `0x08040804` |

### 5.2 应用前台下 IMM32 状态

应用前台时读 `ImmGetDefaultIMEWnd` + `WM_IME_CONTROL(0x283)`：

| 操作 | open（0x5） | conv（0x1） |
|---|---|---|
| 初始 | **0** | 1 |
| 模拟 LShift | **0** | 1 |
| 模拟 Ctrl+Space | **0** | 1 |
| 模拟 RShift | **0** | 1 |

注意：`open=0` 不能用来判定英文——本次审计在记事本中实测，任务栏显示"中"和"英"时 `open` 均为 0、`conv` 均为 1。微信输入法不更新这个值。

### 5.3 GitHub Desktop 对照

脚本：`.resultverify/ime_20261123/probe_timeline2.py`（14 秒，200ms）

| 时间 | 窗口链 | HIMC | 焦点 |
|---|---|---|---|
| 1.22s 窗口出现 | `Chrome_WidgetWin_1` → `Chrome_RenderWidgetHostHWND` → `Intermediate D3D Window` | 无 | `Chrome_WidgetWin_1` |
| 1.22s~14s | 不变 | 无 | 不变 |

GitHub Desktop 的 IME 恢复发生在 TSF 层，不在 IMM32 层。

### 5.4 手动/程序按键对比

- 用户手动按 `Shift`：能切回中文
- 程序 `keybd_event(0x10)`：IMM32 `open` 仍 0
- 说明微信输入法实际可能已经切换，但 **IMM32 不反馈状态**

---

## 6. 根因机制

1. **Chromium 的既定策略**：窗口焦点在非可编辑元素时执行
   `ImmAssociateContextEx(hwnd, NULL, 0)`（`IMM32Manager::DisableIME`）。
2. **pywebview 的焦点时机**：`.venv/.../webview/platforms/winforms.py:367-370`
   在 `Form.Shown` 的同步栈里调用 `self.browser.webview.Focus()`，早于 TSF 焦点文档建立。
3. **微信输入法是纯 TSF**：不依赖 IMM32，中英文状态由 `wetype_server` 全局维护，
   只跟随“窗口焦点变化”，不跟随页面 `document.activeElement`。
4. **页面没有可编辑焦点**：`webapp/static/app.js` 启动期没有 `focus()`；
   `#range-text` 默认 `disabled`（`index.html:145`）。
5. **Windows 输入法上下文线程共享**：同一线程内窗口共享默认上下文（MS Learn《Input Context》）。

结果：窗口激活 → Chromium 关 IME → 微信切英文 → 页面聚焦不产生它认的 TSF 事件 → 停在英文；
切到其他应用时微信的全局状态没有重置。

---

## 7. 已排除方案（不要重复踩坑）

| 方案 | 实测结果 | 结论 |
|---|---|---|
| TSF 预热：`focus()` → `blur()` 空 textarea | 日志 `True`，HIMC 仍 0 | 无效 |
| 控件级延迟 `webview.Focus()` | 日志已提交，无变化 | 无效 |
| 聚焦 `#copies`（number input） | 日志 `ok`，用户验证仍英文 | 无效 |
| 聚焦 `#range-text`（text，临时启用） | 日志 `ok`，用户验证仍英文 | 无效 |
| 外部 `ImmAssociateContextEx` / `ImmSetConversionStatus` | Tauri #15436 实测被 Chromium 下一帧覆盖 | 无效 |
| InputTip 窗口规则 `switchStateCN-LShift` | 仍英文 | 无效 |
| InputTip `switchStateCN/EN-LShift` | 仍英文 | 无效 |
| InputTip `idleTimer` 1200ms 触发 | 仍英文 | 无效 |
| InputTip 自定义模式（默认 EN，状态码 `1`→CN） | 仍英文 | 无效 |
| `focus=False` | 未实施 | `WS_EX_NOACTIVATE` 可能导致窗口无法点击激活 |
| API hook 注入 `msedgewebview2.exe` | 未实施 | 需注入 + hook，随 Edge 更新失效，风险高 |

### InputTip 失败的具体机制

InputTip 切换后 200ms 执行 `_checkAgain`，再次读 IMM32 `open`；微信输入法不更新该值（恒 0），
InputTip 判定“没切成功”，**再发一次切换键**，把刚切好的中文切回英文。

所以任何依赖 IMM32 状态反馈的工具（InputTip、KBLAutoSwitch 等）都无法管理微信输入法。

---

## 8. 修复方向

按可行性排序。

### 方向 A（**已实现并实测有效**，属于补偿）

`webapp/bridge.py::_restore_ime()`：`shown` + 0.4 秒后，确认本窗口可见且未最小化，
补发**一次**真实 Shift；`attach()` 已投递该任务。

```python
if not window.events.shown.wait(15):
    return
time.sleep(IME_RESTORE_DELAY_SECONDS)   # 0.4
if not _is_own_window_visible(window):
    return
_send_single_shift()                     # 只发一次，不回读状态
```

- 关键：**只发一次，不读状态校验**，否则会像 InputTip 一样二次切换
- 实测：修复后 1.67s 跳英 → 3.87s 恢复中 → 切记事本仍中
- 开关：`PDF_DUPLEX_RESTORE_IME=0` 关闭；自检模式跳过
- 风险：盲发；若某次启动没变英文，会把中文切成英文

### 方向 B（未验证）

- `create_window(..., focus=False)` 让窗口不抢前台，再在 shown 后移除 `WS_EX_NOACTIVATE`
  （`winforms.py:294-299`），用 `SetWindowPos(..., SWP_NOACTIVATE)` 保持不激活
- 目标：不触发微信输入法的全局切换；用户点击窗口后再激活
- 需要实测：窗口是否可见、能否从任务栏/点击激活

### 方向 C（未验证，风险较高）

- Win32 `AttachThreadInput` + `SetFocus` 到 `Chrome_RenderWidgetHostHWND`
- 或给 WebView2 控件发 `WM_SETFOCUS`
- 目标：让 Chromium/TSF 真正建立焦点文档，触发 `EnableIME`
- 风险：跨进程/跨线程，可能被 Chromium 覆盖

### 方向 D（不推荐）

- 外部 `ImmAssociateContextEx` / `ImmSetConversionStatus`：已被证明会被 Chromium 覆盖
- API hook 注入：高风险、随 Edge 更新失效
- 让应用改用微软拼音：改变用户使用习惯，不在本任务范围

---

## 9. 验收标准

| 项 | 期望 |
|---|---|
| 启动后微信输入法 | 中文（状态条“中”，可出候选窗） |
| 应用内外切换 | 切到记事本/浏览器仍中文 |
| 窗口交互 | 可正常点击、拖动、最小化、任务栏激活 |
| `pytest tests -q` | 全绿（当前 297；需让测试进程不注入 Agent 安全 shim，否则 1 项环境冲突失败） |
| `WV_SELFTEST` | PASS（修复后实测） |
| 不引入 | 全局键盘钩子、进程注入、管理员权限要求 |

---

## 10. 关键文件与行号

| 文件 | 作用 |
|---|---|
| `webapp/main.py:92-106` | `create_window` / `webview.start(gui="edgechromium")` |
| `webapp/bridge.py:222-247` | `attach()`，窗口事件与启动任务入口 |
| `.venv/Lib/site-packages/webview/platforms/winforms.py:294-299` | `focus=False` 时加 `WS_EX_NOACTIVATE` |
| `.venv/Lib/site-packages/webview/platforms/winforms.py:367-370` | `on_shown` 里 `browser.webview.Focus()` |
| `webapp/static/index.html:145` | `#range-text` 默认 `disabled` |
| `webapp/static/app.js:48,469-470` | 现有 `focus()` 只在弹窗内调用 |
| `app_identity.py:23,57-67` | 显式 AUMID |
| `webapp/bridge.py` 顶部 | `IME_RESTORE_DELAY_SECONDS`、`_is_own_window_visible()`、`_send_single_shift()` |
| `webapp/bridge.py::_restore_ime` | 一次性 Shift 补偿，含开关与自检跳过 |
| `tests/test_webapp_bridge.py` | `attach` 任务顺序契约 + 2 项补偿行为测试 |
| `docs/IME_SWITCH_ANALYSIS.md` | 本次审计的完整、可信分析（取代旧结论） |

---

## 11. 验证脚本

| 脚本 | 用途 |
|---|---|
| `.resultverify/ime_20261123/probe_ime_ab.py` | 启动应用，采样窗口链 HIMC/HKL；支持 `--prime 0/1` 环境变量切换 |
| `.resultverify/ime_20261123/probe_timeline2.py` | GitHub Desktop 启动 IME 时间线 |
| `.resultverify/ime_20261123/ime_timeline_probe_v2.py` | 修复前后真实启动的"中/英"状态时间线（UIA+模板识别） |
| `.resultverify/ime_20261123/fix_test_controller.py` / `fix_test_app.py` | 候选方案 A/B（none/focus/shift/shift_if_en） |
| `.resultverify/ime_20261123/imm_status_probe.py` | 证明 IMM32 `open` 在中/英两态恒为 0 |
| `%APPDATA%\PDFDuplexPrint\logs\app.log` | 应用运行日志（预热/聚焦/Shift 记录位置） |

注意：`probe_ime_ab.py` 的 HIMC/HKL 采样只能证明"无 IMM32 上下文"，不能判定"中/英"；判定请用 `ime_vision.py`。

---

## 12. 禁止事项

- 不要手动修改 `D:\InputTip\src\data\config.ini`（InputTip 文档要求走托盘菜单）
- 不要依赖 `ImmGetContext` 判断微信输入法状态（恒 0）
- 不要用“切换后读状态再重试”的逻辑（会导致二次切换）
- 不要引入全局键盘钩子或向 `msedgewebview2.exe` 注入 DLL
- 修改 `bridge.attach()` 时同步更新 `tests/test_webapp_bridge.py:196-227` 的契约断言