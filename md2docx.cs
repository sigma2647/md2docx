#:package DocumentFormat.OpenXml@3.1.0
#:property TargetFramework=net10.0
// MD to DOCX Converter - 精确匹配模板样式
// 标题1-5: 黑体四号(14pt), 正文: 仿宋四号, 所有字体黑色
// 图片: 黑色边框0.75磅


using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace md2docx;

class Program
{
    private const int A4W = 11906;
    private const int A4H = 16838;
    private const string BLACK = "000000";  // 所有字体统一黑色

    private static uint _docPrId = 1;
    private static int _figCounter = 0;
    private static int _bookmarkId = 0;

    static void Main(string[] args)
    {
        string mdPath = args.Length > 0 ? args[0] : "input.md";
        string outPath = args.Length > 1 ? args[1] : "output.docx";

        if (!File.Exists(mdPath))
        {
            Console.WriteLine($"Error: Markdown file not found: {mdPath}");
            Console.WriteLine("Usage: dotnet run -- <input.md> [output.docx]");
            Environment.Exit(1);
        }

        string mdText = File.ReadAllText(mdPath);
        var mdDir = Path.GetDirectoryName(Path.GetFullPath(mdPath)) ?? ".";

        using var doc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        // 1. 添加模板样式
        AddStyles(mainPart);

        // 2. 添加文档设置
        AddDocumentSettings(mainPart);

        // 3. 解析 Markdown
        ParseMarkdown(body, mainPart, mdText, mdDir);

        // 4. 页面设置
        body.Append(new SectionProperties(
            new PageSize { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 720, Footer = 720 }
        ));

        doc.Save();
        Console.WriteLine($"Generated: {Path.GetFullPath(outPath)}");
        Console.WriteLine($"Size: {new FileInfo(outPath).Length} bytes");
    }

    // =========================================================================
    // 样式定义 - 精确匹配模板
    // =========================================================================
    static void AddStyles(MainDocumentPart mainPart)
    {
        var sp = mainPart.AddNewPart<StyleDefinitionsPart>();
        sp.Styles = new Styles();

        // ---- Normal (默认段落字体) ----
        // 模板: widowControl="0", 两端对齐, 段前段后0磅
        sp.Styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new StyleParagraphProperties(
                new WidowControl { Val = false },
                new Justification { Val = JustificationValues.Both },
                new SpacingBetweenLines { Before = "0", After = "0" }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        // ---- Heading1 ----
        // 模板: 黑体, sz=28, szCs=28, line=360 auto, keepNext/keepLines, outline=0, bCs, kern=44
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 1" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 0 }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new BoldComplexScript(),
                new Kern { Val = 44 },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        // ---- Heading2 ----
        // 模板: 黑体, sz=28, szCs=28, line=360 auto, keepNext/keepLines, outline=1, bCs
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 2" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 1 }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new BoldComplexScript(),
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading2" });

        // ---- Heading3 ----
        // 模板: 黑体, sz=28, szCs=32, line=360 auto, keepNext/keepLines, wordWrap=0, outline=2, bCs
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 3" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new WordWrap { Val = false },
                new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 2 }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new BoldComplexScript(),
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "32" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading3" });

        // ---- Heading4 ----
        // 模板: 黑体, sz=28, szCs=28, line=360 auto, keepNext/keepLines, wordWrap=0, outline=3
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 4" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new WordWrap { Val = false },
                new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 3 }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading4" });

        // ---- Heading5 ----
        // 模板: 黑体, sz=28, szCs=28, line=360 auto, keepNext/keepLines, wordWrap=0, outline=4
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 5" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new WordWrap { Val = false },
                new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new OutlineLevel { Val = 4 }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading5" });

        // ---- 正文样式 (No Spacing / BodyText) ----
        // 模板: 仿宋, sz=28, szCs=28, widowControl=0, wordWrap=0
        // firstLineChars=200 firstLine=200, 两端对齐, 段前段后0磅
        sp.Styles.Append(new Style(
            new StyleName { Val = "Body Text" },
            new Aliases { Val = "正文" },
            new StyleParagraphProperties(
                new WidowControl { Val = false },
                new WordWrap { Val = false },
                new Indentation { FirstLineChars = 200, FirstLine = "200" },
                new Justification { Val = JustificationValues.Both },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "BodyText" });

        // ---- 题注样式 ----
        sp.Styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new FontSize { Val = "24" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Caption" });

        // ---- 表格样式 ----
        sp.Styles.Append(new Style(
            new StyleName { Val = "Table Grid" },
            new StyleTableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
                ),
                new TableCellMarginDefault(
                    new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new StartMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new EndMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                )
            )
        ) { Type = StyleValues.Table, StyleId = "TableGrid" });
    }

    static void AddDocumentSettings(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new UpdateFieldsOnOpen { Val = true }
        );
    }

    // =========================================================================
    // Markdown 解析器
    // =========================================================================
    static void ParseMarkdown(Body body, MainDocumentPart mainPart, string md, string baseDir)
    {
        var lines = md.Split('\n');
        var i = 0;

        while (i < lines.Length)
        {
            var line = lines[i];
            var trimmed = line.TrimStart();

            if (string.IsNullOrWhiteSpace(line))
            {
                i++;
                continue;
            }

            // 标题
            if (trimmed.StartsWith("#"))
            {
                int level = 0;
                while (level < trimmed.Length && trimmed[level] == '#') level++;
                if (level <= 5 && trimmed.Length > level && trimmed[level] == ' ')
                {
                    var text = trimmed.Substring(level + 1).Trim();
                    body.Append(CreateHeading(text, level));
                    i++;
                    continue;
                }
            }

            // 图片
            if (trimmed.StartsWith("![") && trimmed.Contains("]("))
            {
                var imgInfo = ParseImageTag(trimmed);
                if (imgInfo != null)
                {
                    AddImageWithCaption(body, mainPart, Path.Combine(baseDir, imgInfo.Path), imgInfo.Alt);
                }
                i++;
                continue;
            }

            // 表格
            if (trimmed.StartsWith("|") && i + 1 < lines.Length && lines[i + 1].Trim().Contains("|-"))
            {
                i = ParseTable(body, lines, i);
                continue;
            }

            // 普通段落
            var paraText = line.Trim();
            i++;
            while (i < lines.Length && !string.IsNullOrWhiteSpace(lines[i]) &&
                   !lines[i].TrimStart().StartsWith("#") &&
                   !lines[i].TrimStart().StartsWith("|") &&
                   !(lines[i].TrimStart().StartsWith("![") && lines[i].Contains("](")))
            {
                paraText += lines[i].Trim();
                i++;
            }
            body.Append(CreateRichParagraph(paraText));
        }
    }

    static Paragraph CreateHeading(string text, int level)
    {
        var styleId = $"Heading{level}";
        var bmkId = (++_bookmarkId).ToString();
        var bmkName = $"_Toc{_bookmarkId:D3}";

        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
            new BookmarkStart { Id = bmkId, Name = bmkName },
            new Run(new Text(text)),
            new BookmarkEnd { Id = bmkId }
        );
    }

    static Paragraph CreateRichParagraph(string text)
    {
        var para = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" })
        );

        var parts = ParseInlineStyles(text);
        foreach (var part in parts)
        {
            var rpr = new RunProperties(
                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            );
            if (part.IsBold) rpr.Append(new Bold());
            if (part.IsItalic) rpr.Append(new Italic());
            para.Append(new Run(rpr, new Text(part.Text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return para;
    }

    static List<TextPart> ParseInlineStyles(string text)
    {
        var parts = new List<TextPart>();
        int i = 0;

        while (i < text.Length)
        {
            if (i + 1 < text.Length && text[i] == '*' && text[i + 1] == '*')
            {
                int end = text.IndexOf("**", i + 2);
                if (end > 0)
                {
                    parts.Add(new TextPart { Text = text.Substring(i + 2, end - i - 2), IsBold = true });
                    i = end + 2;
                    continue;
                }
            }
            if (text[i] == '*')
            {
                int end = text.IndexOf('*', i + 1);
                if (end > 0)
                {
                    parts.Add(new TextPart { Text = text.Substring(i + 1, end - i - 1), IsItalic = true });
                    i = end + 1;
                    continue;
                }
            }

            int nextSpecial = int.MaxValue;
            int boldPos = text.IndexOf("**", i);
            int italicPos = text.IndexOf('*', i);
            if (boldPos >= 0) nextSpecial = Math.Min(nextSpecial, boldPos);
            if (italicPos >= 0) nextSpecial = Math.Min(nextSpecial, italicPos);

            if (nextSpecial == int.MaxValue)
            {
                parts.Add(new TextPart { Text = text.Substring(i) });
                break;
            }
            else
            {
                parts.Add(new TextPart { Text = text.Substring(i, nextSpecial - i) });
                i = nextSpecial;
            }
        }

        if (parts.Count == 0) parts.Add(new TextPart { Text = text });
        return parts;
    }

    // =========================================================================
    // 图片 + 题注（SEQ 域）+ 黑色边框 0.75磅
    // =========================================================================
    static void AddImageWithCaption(Body body, MainDocumentPart mainPart, string imagePath, string caption)
    {
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Warning: Image not found: {imagePath}");
            body.Append(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(new Color { Val = "E74C3C" }, new FontSize { Val = "28" }),
                    new Text($"[图片未找到: {Path.GetFileName(imagePath)}]")
                )
            ));
            return;
        }

        _figCounter++;

        // 添加 ImagePart
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        using (var ms = new MemoryStream(imageBytes)) imagePart.FeedData(ms);
        var imageId = mainPart.GetIdOfPart(imagePart);

        // 读取 PNG 尺寸
        int imgWidth, imgHeight;
        using (var ms = new MemoryStream(imageBytes))
        {
            ms.Seek(16, SeekOrigin.Begin);
            byte[] wb = new byte[4], hb = new byte[4];
            ms.Read(wb, 0, 4); ms.Read(hb, 0, 4);
            if (BitConverter.IsLittleEndian) { Array.Reverse(wb); Array.Reverse(hb); }
            imgWidth = BitConverter.ToInt32(wb, 0);
            imgHeight = BitConverter.ToInt32(hb, 0);
        }

        // 计算尺寸（最大宽度 15cm）
        long maxWidthEmu = 15 * 360000L;
        long cx = maxWidthEmu;
        long cy = (long)(cx * ((double)imgHeight / imgWidth));
        uint prId = _docPrId++;

        // 0.75磅 = 9525 EMU (1磅 = 12700 EMU)
        long borderWidth = 9525;

        // 创建带黑色边框的图片 (分步构造避免括号嵌套)
        var shapeProps = new PIC.ShapeProperties(
            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle },
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = BLACK })) { Width = (int)borderWidth }
        );

        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0, Name = $"fig{_figCounter}.png" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(new A.Blip { Embed = imageId }, new A.Stretch(new A.FillRectangle())),
            shapeProps
        );

        var graphicData = new A.GraphicData(picture)
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

        var graphic = new A.Graphic(graphicData);

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = borderWidth, TopEdge = borderWidth, RightEdge = borderWidth, BottomEdge = borderWidth },
            new DW.DocProperties { Id = prId, Name = $"Fig{_figCounter}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            graphic
        )
        { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 };

        var drawing = new Drawing(inline);

        // 图片段落（居中，段前段后0磅）
        body.Append(new Paragraph(
            new ParagraphProperties(
                new KeepNext(),
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new Run(drawing)));

        // 题注（SEQ 域）
        var capPara = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" })
        );

        capPara.Append(new Run(
            new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
            new Text("图 ") { Space = SpaceProcessingModeValues.Preserve }
        ));

        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        capPara.Append(new Run(new FieldCode(" SEQ Figure \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        capPara.Append(new Run(new Text(_figCounter.ToString())));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        capPara.Append(new Run(
            new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
            new Text($" - {caption}") { Space = SpaceProcessingModeValues.Preserve }
        ));
        body.Append(capPara);
    }

    // =========================================================================
    // 表格解析
    // =========================================================================
    static int ParseTable(Body body, string[] lines, int startIdx)
    {
        var headerLine = lines[startIdx].Trim();
        var headers = ParseTableRow(headerLine);

        int dataStart = startIdx + 2;
        var rows = new List<string[]>();
        int i = dataStart;
        while (i < lines.Length && lines[i].Trim().StartsWith("|"))
        {
            rows.Add(ParseTableRow(lines[i].Trim()));
            i++;
        }

        var table = new Table();
        table.Append(new TableProperties(
            new TableStyle { Val = "TableGrid" },
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableJustification { Val = TableRowAlignmentValues.Center }
        ));

        var grid = new TableGrid();
        int colCount = headers.Length;
        int colWidth = 9000 / colCount;
        for (int c = 0; c < colCount; c++)
            grid.Append(new GridColumn { Width = colWidth.ToString() });
        table.Append(grid);

        // 表头行
        var headerRow = new TableRow();
        headerRow.Append(new TableRowProperties(new TableHeader()));
        foreach (var h in headers)
        {
            headerRow.Append(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = colWidth.ToString(), Type = TableWidthUnitValues.Dxa },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "000000" },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new Indentation { FirstLine = "0" }
                    ),
                    new Run(
                        new RunProperties(
                            new Bold(),
                            new Color { Val = "FFFFFF" },
                            new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                            new FontSize { Val = "28" }
                        ),
                        new Text(h.Trim())
                    )
                )
            ));
        }
        table.Append(headerRow);

        // 数据行
        bool altBg = false;
        foreach (var rowData in rows)
        {
            var dataRow = new TableRow();
            for (int c = 0; c < colCount; c++)
            {
                var cellProps = new TableCellProperties(
                    new TableCellWidth { Width = colWidth.ToString(), Type = TableWidthUnitValues.Dxa }
                );
                if (altBg)
                    cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" });

                dataRow.Append(new TableCell(
                    cellProps,
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center },
                            new Indentation { FirstLine = "0" }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                                new FontSize { Val = "28" },
                                new Color { Val = BLACK }
                            ),
                            new Text(c < rowData.Length ? rowData[c].Trim() : "")
                        )
                    )
                ));
            }
            table.Append(dataRow);
            altBg = !altBg;
        }

        body.Append(table);
        return i;
    }

    static string[] ParseTableRow(string line)
    {
        var inner = line.Trim();
        if (inner.StartsWith("|")) inner = inner.Substring(1);
        if (inner.EndsWith("|")) inner = inner.Substring(0, inner.Length - 1);
        return inner.Split('|');
    }

    // =========================================================================
    // 辅助结构
    // =========================================================================
    class TextPart
    {
        public string Text = "";
        public bool IsBold;
        public bool IsItalic;
    }

    class ImageInfo
    {
        public string Alt = "";
        public string Path = "";
    }

    static ImageInfo? ParseImageTag(string line)
    {
        var match = Regex.Match(line, @"!\[(.*?)\]\((.*?)\)");
        if (!match.Success) return null;
        return new ImageInfo { Alt = match.Groups[1].Value, Path = match.Groups[2].Value };
    }
}
