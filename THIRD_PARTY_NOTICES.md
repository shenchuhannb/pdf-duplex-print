# 第三方组件与许可证说明

本项目分发时需同时保留各组件许可证。以下信息来自 `requirements.txt` 和当前包元数据，不替代完整的第三方 LICENSE 文本。

| 组件 | 版本 | 许可证/注意事项 |
|---|---|---|
| PyMuPDF | 1.28.0 | GNU AGPL v3 或 Artifex 商业许可，双许可 |
| pywin32 | 311 | PSF |
| Pillow | 12.0.0 | MIT-CMU |
| tkinterdnd2 | 0.7.0 | 以随包许可证为准 |
| PyInstaller | 6.22.3 | GPL（含 bootloader 例外；以其官方许可证为准） |

## 发布前必须确认

PyMuPDF 的许可证选择取决于项目本身的发布方式。若按 AGPL 分发，需要履行对应源码和许可证义务；若项目不按 AGPL 分发，需要取得 Artifex 商业许可。当前仓库没有足够信息替维护者作出选择。

## 完整性

`tools/build_exe.bat` 会在冻结自检通过后生成 `dist/PDFDuplexPrint/SHA256SUMS.txt`。代码签名证书不包含在仓库中，发布者如需 Authenticode 签名，必须在自己的受控环境中使用合法证书完成，不能把 SHA-256 清单误当作发布者身份签名。
