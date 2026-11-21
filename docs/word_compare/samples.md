# 样本清单与构建方式

## 构建方式

- 生成脚本：`scripts/make_samples.py`（Word COM，`DispatchEx("Word.Application")`，以 `Selection.TypeText/TypeParagraph` 逐段写入）。
- 生成时间：2026-11-04；全部样本在正式测量**之前**一次性生成，测量阶段未替换、未删除任何样本。
- 原始文件与正文导出：`samples/`（`.docx` / `.doc` / `*.source.txt` / `samples.json`）。
- 页面尺寸 A4，默认正文宋体 10.5pt，按样本需要覆盖不同字号与字体。
- 09 的“缺失字体”是在生成后直接改写 OOXML 的 `<w:rFonts>` 注入 `NoSuchFontABC`（Word 在输入时会静默替换不存在的字体，若不改写 XML 则样本里根本不存在该字体名）。脚本：`scripts/fix_missing_font.py`。
- 07 的页脚域顺序在首次生成时写错（`第 页 / 共 页51`），正式测量前用 `scripts/fix_footer07.py` 修正为 `第 {PAGE} 页 / 共 {NUMPAGES} 页`，并重测该样本；其余 9 份未做任何修改。

## 覆盖矩阵

| # | 文件 | 页数 | 内容与字号 | 声明字体（OOXML `rFonts` / .doc 采样） | 覆盖点 |
|---|---|---|---|---|---|
| 01 | 01_mixed_short.docx | 1 | 中英混排短文档，10.5/11/12/16pt，含中英文标题与正文 | Times New Roman、宋体、微软雅黑 | 中英混排、短文档、字号跨度 |
| 02 | 02_chinese_pure.docx | 2 | 纯中文，宋体/黑体/楷体 12pt | Times New Roman、宋体、楷体、黑体 | 纯中文、多中文字体 |
| 03 | 03_english_pure.docx | 6 | 纯英文，Calibri / Times New Roman / Arial 10–12pt | Arial、Calibri、Times New Roman、宋体 | 纯英文、拉丁字体、中篇 |
| 04 | 04_table.docx | 4 | 12×5、8×4 两张表格，9pt，含合并单元格与中英数字 | Calibri、Times New Roman、宋体、微软雅黑 | 表格、边框、小字号 |
| 05 | 05_image.docx | 2 | 插入 12cm×7cm PNG 图表，含图注 | Times New Roman、宋体、微软雅黑 | 图片、图文混排 |
| 06 | 06_formula.docx | 1 | 真正的 OMML 公式（`OMaths.Add` + `BuildUp`）与特殊符号，Cambria Math 13–14pt | Cambria Math、Times New Roman、宋体、微软雅黑 | 公式、特殊符号 |
| 07 | 07_header_footer.docx | 5 | 页眉、页脚、页码域，11pt | Times New Roman、宋体、微软雅黑 | 页眉/页脚/页码 |
| 08 | 08_legacy.doc | 2 | 旧二进制 `.doc`（`SaveAs2(FileFormat=0)`），中英混排 | Times New Roman、宋体、黑体 | 旧格式 `.doc` |
| 09 | 09_missing_font.docx | 1 | 正文与标题强制声明 `NoSuchFontABC`，12/14pt | NoSuchFontABC、Times New Roman、宋体 | 本机缺失字体 |
| 10 | 10_long_35p.docx | 35 | 长文档，10.5–12pt，覆盖多章节 | Times New Roman、宋体、微软雅黑、黑体 | 30 页以上长文档 |

字号与字体清单来自 `samples/samples.json`：`.docx` 解析 `word/document.xml`、`styles.xml`、`header/footer*.xml` 的 `rFonts`；`.doc` 用 Word 逐字符抽样（步长 `Count//200`）。

## 筛选标准

1. 先按需求列出必须覆盖的类别（中英混排 / 纯中文 / 纯英文 / 表格 / 图片 / 公式或特殊符号 / 页眉页脚页码 / `.doc` / 缺失字体 / 30 页以上），每类至少 1 份。
2. 页数覆盖 1、2、4、5、6、35 页；短文档与长文档都在同一轮测量中，不因结果好坏增删。
3. 全部 10 份都进入 A/B 两条路径、3 次重复与全部分析；没有“效果不好就换样本”的操作。
4. 已知失败只记录不替换：B 冷启动导出 08_legacy 第 3 次出现 COM `接口未知`，该次记为失败，统计用其余 2 次成功值。
