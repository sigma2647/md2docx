# md2docx — 项目约定与踩坑

Markdown → 公文 docx 转换器。这里只写**读代码看不出来**的坑。

## 字体规范

- 标题 (Heading 1–5)：**黑体 14pt 不加粗**，所有级别外观一致，仅 `outlineLvl` 区分大纲层级
- 正文 (BodyText)：**仿宋 14pt**，首行缩进 2 字符
- 题注 (Caption)：黑体小五 (10pt)，居中

## 头号大坑：`<w:docGrid>` 是中文 docx 的命门

**症状**：标题段前段后视觉不对称——段前异常低、段后正常；改 paragraph spacing 怎么改都不对。

**根因**：中文 Word/WPS 文档的 `<w:sectPr>` 必须含 `<w:docGrid w:type="lines" w:linePitch="312"/>`：

- **作用**：把每一行锁定到 312 twips（约 15.6pt）的固定网格上
- **没有它时**：行高完全由字号 + line 计算，中文字体（黑体、仿宋）的 ascent/descent **不对称**——段前留白小于段后，看起来"塌顶"
- **不是 paragraph spacing 的问题**：怎么调 `Before/After` 都治标不治本

**修复**：在 `Run()` 构造 sectPr 时加：
```csharp
new Columns { Space = "425" },
new DocGrid { Type = DocGridValues.Lines, LinePitch = 312 }
```

**反面教材**（已修）：之前花了 5 轮调 `Before/After/Line` 都没用，因为根本不在那一层。

## 调试 docx 渲染问题的方法论

**LibreOffice / pandoc / Linux 不能当验证标准** ——它们对 OOXML 的解释和 WPS/Word 不完全一致：

- LibreOffice 不严格依赖 docGrid，所以本地测试看不出 docGrid 缺失
- WPS 严格按 OOXML 处理，没 docGrid 就用字体度量 fallback

**正确流程**：
1. 拿到一个"正常"的样本 docx（Word/WPS 创建的）
2. 用 `python scripts/office/unpack.py` 解包两份
3. diff 两份的 `word/document.xml` 和 `word/styles.xml`
4. 真正的差异往往不在 `<w:spacing>`，而是在 `sectPr / docDefaults / docGrid / 命名空间`
5. **在 WPS 实际验证**，不要相信 LibreOffice 渲染

参见 `不正常.docx` vs `正常.docx` 的对比 — 真正的 root cause 隐藏在 sectPr 的两行里。

## 双保险是反模式

**反例**：在 style 里设了 黑体14pt 不加粗，然后又在每个 run 的 `<w:rPr>` 里重复声明一遍 + 写 `<w:b w:val="false"/>`。

**问题**：
- Run 里的 `<w:b w:val="false"/>` 会**强制覆盖** style 的 `<w:bCs/>`（CJK 加粗触发器），导致 WPS 字重渲染异常
- 多余声明只会引入 bug，不会增加稳定性

**正确做法**：style 单点负责字体/字号/加粗。Run 里**只**在两种情况写属性：
1. `<w:rFonts w:hint="eastAsia"/>` —— 告诉 WPS 走东亚渲染路径（必需）
2. 真正的内联格式（`**bold**` `*italic*`）

## 题注：必须用 SEQ 域

图题/表题统一 `图 N caption` / `表 N caption`，编号**必须**用 Word `SEQ` 域：
```
图 { SEQ 图 \* ARABIC } caption
```
参见 `AppendCaptionField`。这样后续手工插入图/表时，编号和我们生成的共用同一 Word 计数器，不会断裂。

## 题注防跨页

OOXML 没有"keepWithPrevious"——题注要粘住前面的图/表，必须给**前面的元素**加 `<w:keepNext/>`。
我们的做法：

- **图片段落**不加 `KeepNext()` → 允许图片与图题被分到不同页（避免图片被强推到下一页造成大块空白）
- **表格每行**加 `CantSplit()` → 单行不允许拦腰切开
- **表格单元格内段落**加 `KeepNext()` → 整表 + 表题作为不可分组合

**副作用**：表格本身大于一页时，Word 会把整表强推到下一页（可能造成空白）。本项目场景表格通常较小，可接受这个代价。如果遇到必然超长的表格，需要单独豁免 KeepNext。

## Word/WPS 二次保存的副作用（不是 bug）

打开生成的 docx 又保存后会看到：

- styleId `"Heading1"` → `"1"`、`"Normal"` → `"a"`、`"BodyText"` → `"a3"`
- 多出 `theme/`、`webSettings.xml`、`fontTable.xml`
- 段落多出 `w14:paraId` `w:rsidR` 等 Word 跟踪 ID
- styles.xml 顶部加 `<w:docDefaults>` 和 376 条 `<w:latentStyles>`

这些是 Word 在补齐它认为完整文档应有的元数据，**不影响样式渲染**。排查 C# 输出 bug 时，**用未经 Word 打开的原始 docx**，不要被这层污染迷惑。

## 图片三种语法

`ParseImage` 同时支持：

- Obsidian：`![[path|171]]` — 171 = 绝对像素宽
- Typora：`<img src="..." style="zoom:80%">` — 相对原图百分比
- 标准 Markdown：`![alt](path)` — 撑满页宽

像素 → EMU 换算用图片自带 DPI（PNG `pHYs` chunk / JPEG `JFIF` APP0），不硬编码 96。

## 页面：A4

`A4W=11906 A4H=16838`（DXA），左右边距 1800（≈3.17cm）。公文标准，不要改成 US Letter。

## Build

```bash
just install   # 单文件自包含 → ~/.local/bin/md2docx
just build     # 仅快速编译
```

依赖 `DocumentFormat.OpenXml` + `System.CommandLine`。
