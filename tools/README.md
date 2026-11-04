# tools 目录说明

开发与维护辅助脚本（不属于运行时程序）：

- `create_shortcut.py` — 创建/更新桌面快捷方式（指向 .venv 的 pythonw，写入 AUMID）
- `verify_identity.py` — 校验任务栏图标/AUMID/快捷方式是否生效
- `make_icons.py` — 重新生成 app_icon.ico（多尺寸）
- `capture_ui.py` — UI 截图工具（供文档与回归对比）

文档资源在 `docs/`：word_compare 为 Word 转换质量对比实验（含脚本与数据），
screenshots 为历次 UI 改动的过程截图存档。
