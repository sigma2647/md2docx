#:package DocumentFormat.OpenXml@3.1.0
#:package System.CommandLine@2.0.0-beta4.22272.1
#:property TargetFramework=net10.0
// MD to DOCX Converter - 精确匹配模板样式
// 标题1-5: 黑体四号(14pt), 正文: 仿宋四号, 所有字体黑色
// 图片: 黑色边框0.75磅

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.CommandLine;
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
    private const string BLACK = "000000";
    private const string VERSION = "1.1.0";

    private static uint _docPrId = 1;
    private static int _figCounter = 0;
    private static int _tableCounter = 0;
    private static int _bookmarkId = 0;

    static void Die(string msg, int code = 1)
    {
        Console.Error.WriteLine($"Error: {msg}");
        Environment.Exit(code);
    }

    /// <summary>
    /// 解析最终的输出路径，处理冲突（递增后缀）或强制覆盖。
    /// </summary>
    static string ResolveOutputPath(string desired, bool force)
    {
        if (force || !File.Exists(desired))
            return desired;

        // 冲突：追加 _1、_2 …
        var dir  = Path.GetDirectoryName(desired) ?? ".";
        var stem = Path.GetFileNameWithoutExtension(desired);
        var ext  = Path.GetExtension(desired);   // ".docx"

        for (int n = 1; ; n++)
        {
            var candidate = Path.Combine(dir, $"{stem}_{n}{ext}");
            if (!File.Exists(candidate))
                return candidate;
        }
    }

    // =========================================================================
    // 入口 (使用 System.CommandLine 重构)
    // =========================================================================
    static int Main(string[] args)
    {
        // 1. 定义参数 (Argument: 位置参数，无需前缀)
        var inputArgument = new Argument<string?>(
            name: "input.md",
            description: "Input Markdown file. Omit to read from stdin.",
            getDefaultValue: () => null);

        // 2. 定义选项 (Option: 命名参数，带前缀)
        var outputOption = new Option<string?>(
            aliases: new[] { "-o", "--output" },
            description: "Output .docx path.\nDefault: same name as input (.docx), or 'output.docx' when reading from stdin.");

        var baseDirOption = new Option<string?>(
            name: "--base-dir",
            description: "Base directory for resolving image paths.\nDefault: directory of input file, or cwd for stdin.");

        var forceOption = new Option<bool>(
            aliases: new[] { "-f", "--force" },
            description: "Overwrite output file if it already exists.");

        // 3. 配置根命令
        var rootCommand = new RootCommand($"md2docx v{VERSION} - Markdown to DOCX Converter")
        {
            inputArgument,
            outputOption,
            baseDirOption,
            forceOption
        };

        // 4. 绑定执行逻辑 (将解析后的强类型参数传递给业务方法)
        rootCommand.SetHandler((string? inputFile, string? outputFile, string? baseDir, bool force) =>
        {
            RunConverter(inputFile, outputFile, baseDir, force);
        }, inputArgument, outputOption, baseDirOption, forceOption);

        // 5. 调用并返回退出码 (自动接管 -h, --help, --version)
        return rootCommand.Invoke(args);
    }

    /// <summary>
    /// 原来的 Main 方法中的核心业务逻辑提取到这里
    /// </summary>
    static void RunConverter(string? inputFile, string? outputFileOpt, string? baseDirOpt, bool force)
    {
        bool fromStdin = inputFile == null;

        string mdText;
        string baseDir;
        string desiredOutput;

        if (fromStdin)
        {
            if (Console.IsInputRedirected)
            {
                mdText = Console.In.ReadToEnd();
            }
            else
            {
                Console.Error.WriteLine("Reading from stdin (end with Ctrl+D / Ctrl+Z):");
                mdText = Console.In.ReadToEnd();
            }

            baseDir       = baseDirOpt ?? Directory.GetCurrentDirectory();
            desiredOutput = outputFileOpt ?? Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        }
        else
        {
            var inputPath = Path.GetFullPath(inputFile!);
            if (!File.Exists(inputPath))
                Die($"Input file not found: {inputPath}");

            mdText        = File.ReadAllText(inputPath);
            baseDir       = baseDirOpt != null
                                ? Path.GetFullPath(baseDirOpt)
                                : Path.GetDirectoryName(inputPath)!;
            desiredOutput = outputFileOpt
                            ?? Path.ChangeExtension(inputPath, ".docx");
        }

        if (!Directory.Exists(baseDir))
            Die($"Base directory not found: {baseDir}");

        var outPath = ResolveOutputPath(Path.GetFullPath(desiredOutput), force);

        // --- 生成 DOCX ---
        using var doc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        AddStyles(mainPart);
        AddDocumentSettings(mainPart);
        ParseMarkdown(body, mainPart, mdText, baseDir);

        body.Append(new SectionProperties(
            new PageSize  { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 720, Footer = 720 }
        ));

        doc.Save();

        var info = new FileInfo(outPath);
        Console.WriteLine($"Generated : {outPath}");
        Console.WriteLine($"Size      : {info.Length:N0} bytes");

        if (force && desiredOutput != outPath)
            Console.WriteLine("(overwritten)");
        else if (desiredOutput != outPath)
            Console.WriteLine($"Note      : '{Path.GetFileName(desiredOutput)}' already existed → saved as '{Path.GetFileName(outPath)}'");
    }

    // =========================================================================
    // 样式定义
    // =========================================================================
    static void AddStyles(MainDocumentPart mainPart)
    {
        var sp = mainPart.AddNewPart<StyleDefinitionsPart>();
        sp.Styles = new Styles();

        sp.Styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new StyleParagraphProperties(
                new WidowControl { Val = false },
                new Justification { Val = JustificationValues.Both },
                new SpacingBetweenLines { Before = "0", After = "0" }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

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
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Heading3" });

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

        sp.Styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new FontSize { Val = "20" },
                new Color { Val = BLACK }
            )
        ) { Type = StyleValues.Paragraph, StyleId = "Caption" });

        sp.Styles.Append(new Style(
            new StyleName { Val = "Table Grid" },
            new StyleTableProperties(
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder   { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder  { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4, Color = "000000" }
                ),
                new TableCellMarginDefault(
                    new TopMargin    { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new StartMargin  { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new EndMargin    { Width = "80", Type = TableWidthUnitValues.Dxa }
                )
            )
        ) { Type = StyleValues.Table, StyleId = "TableGrid" });
    }

    static void AddDocumentSettings(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings();
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
            var line    = lines[i];
            var trimmed = line.TrimStart();

            if (string.IsNullOrWhiteSpace(line)) { i++; continue; }

            // 标题
            if (trimmed.StartsWith("#"))
            {
                int level = 0;
                while (level < trimmed.Length && trimmed[level] == '#') level++;
                if (level <= 5 && trimmed.Length > level && trimmed[level] == ' ')
                {
                    body.Append(CreateHeading(trimmed.Substring(level + 1).Trim(), level));
                    i++;
                    continue;
                }
            }

            // 图片
            if (trimmed.StartsWith("![") && trimmed.Contains("]("))
            {
                var imgInfo = ParseImageTag(trimmed);
                if (imgInfo != null)
                    AddImageWithCaption(body, mainPart, Path.Combine(baseDir, imgInfo.Path), imgInfo.Alt);
                i++;
                continue;
            }

            // 表格
            if (trimmed.StartsWith("|") && i + 1 < lines.Length && lines[i + 1].Trim().Contains("|-"))
            {
                ParseTable(body, lines, i, out int endIdx);
                i = endIdx;

                string tableCaption = "";
                if (i < lines.Length)
                {
                    var capMatch = Regex.Match(lines[i].Trim(), @"^\[(.+)\]$");
                    if (capMatch.Success) { tableCaption = capMatch.Groups[1].Value; i++; }
                }
                AppendTableCaption(body, tableCaption);
                continue;
            }

            // 普通段落（合并连续非空行）
            var paraText = line.Trim();
            i++;
            while (i < lines.Length
                   && !string.IsNullOrWhiteSpace(lines[i])
                   && !lines[i].TrimStart().StartsWith("#")
                   && !lines[i].TrimStart().StartsWith("|")
                   && !(lines[i].TrimStart().StartsWith("![") && lines[i].Contains("](")))
            {
                paraText += lines[i].Trim();
                i++;
            }
            body.Append(CreateRichParagraph(paraText));
        }
    }

    static Paragraph CreateHeading(string text, int level)
    {
        var bmkId = (++_bookmarkId).ToString();
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{level}" }),
            new BookmarkStart { Id = bmkId, Name = $"_Toc{_bookmarkId:D3}" },
            new Run(new Text(text)),
            new BookmarkEnd { Id = bmkId }
        );
    }

    static Paragraph CreateRichParagraph(string text)
    {
        var para = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" })
        );
        foreach (var part in ParseInlineStyles(text))
        {
            var rpr = new RunProperties(
                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                new FontSize { Val = "28" },
                new FontSizeComplexScript { Val = "28" },
                new Color { Val = BLACK }
            );
            if (part.IsBold)   rpr.Append(new Bold());
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
                if (end > 0) { parts.Add(new TextPart { Text = text.Substring(i + 2, end - i - 2), IsBold = true }); i = end + 2; continue; }
            }
            if (text[i] == '*')
            {
                int end = text.IndexOf('*', i + 1);
                if (end > 0) { parts.Add(new TextPart { Text = text.Substring(i + 1, end - i - 1), IsItalic = true }); i = end + 1; continue; }
            }
            int nextSpecial = int.MaxValue;
            int boldPos   = text.IndexOf("**", i);
            int italicPos = text.IndexOf('*',  i);
            if (boldPos   >= 0) nextSpecial = Math.Min(nextSpecial, boldPos);
            if (italicPos >= 0) nextSpecial = Math.Min(nextSpecial, italicPos);
            if (nextSpecial == int.MaxValue) { parts.Add(new TextPart { Text = text.Substring(i) }); break; }
            parts.Add(new TextPart { Text = text.Substring(i, nextSpecial - i) });
            i = nextSpecial;
        }
        if (parts.Count == 0) parts.Add(new TextPart { Text = text });
        return parts;
    }

    // =========================================================================
    // 图片 + 题注
    // =========================================================================
    static void AddImageWithCaption(Body body, MainDocumentPart mainPart, string imagePath, string caption)
    {
        if (!File.Exists(imagePath))
        {
            Console.Error.WriteLine($"Warning: Image not found: {imagePath}");
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
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        using (var ms = new MemoryStream(imageBytes)) imagePart.FeedData(ms);
        var imageId = mainPart.GetIdOfPart(imagePart);

        int imgWidth, imgHeight;
        using (var ms = new MemoryStream(imageBytes))
        {
            ms.Seek(16, SeekOrigin.Begin);
            byte[] wb = new byte[4], hb = new byte[4];
            ms.Read(wb, 0, 4); ms.Read(hb, 0, 4);
            if (BitConverter.IsLittleEndian) { Array.Reverse(wb); Array.Reverse(hb); }
            imgWidth  = BitConverter.ToInt32(wb, 0);
            imgHeight = BitConverter.ToInt32(hb, 0);
        }

        long maxWidthEmu = 15 * 360000L;
        long cx = maxWidthEmu;
        long cy = (long)(cx * ((double)imgHeight / imgWidth));
        uint prId = _docPrId++;
        long borderWidth = 9525; // 0.75pt

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

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = borderWidth, TopEdge = borderWidth, RightEdge = borderWidth, BottomEdge = borderWidth },
            new DW.DocProperties { Id = prId, Name = $"Fig{_figCounter}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        ) { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 };

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new Run(new Drawing(inline))));

        var capPara = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }));
        capPara.Append(new Run(
            new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
            new Text("图 ") { Space = SpaceProcessingModeValues.Preserve }
        ));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        capPara.Append(new Run(new FieldCode(" SEQ 图 \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        capPara.Append(new Run(new Text(_figCounter.ToString())));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        capPara.Append(new Run(
            new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
            new Text($" {caption}") { Space = SpaceProcessingModeValues.Preserve }
        ));
        var goBackId = (++_bookmarkId).ToString();
        capPara.Append(new BookmarkStart { Id = goBackId, Name = "_GoBack" });
        capPara.Append(new BookmarkEnd { Id = goBackId });
        body.Append(capPara);
    }

    // =========================================================================
    // 表格题注
    // =========================================================================
    static void AppendTableCaption(Body body, string captionText)
    {
        _tableCounter++;
        var capPara = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }));
        capPara.Append(new Run(
            new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
            new Text("表 ") { Space = SpaceProcessingModeValues.Preserve }
        ));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        capPara.Append(new Run(new FieldCode(" SEQ 表 \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        capPara.Append(new Run(new Text(_tableCounter.ToString())));
        capPara.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        if (!string.IsNullOrWhiteSpace(captionText))
            capPara.Append(new Run(
                new RunProperties(new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" }),
                new Text($" {captionText}") { Space = SpaceProcessingModeValues.Preserve }
            ));
        var goBackId = (++_bookmarkId).ToString();
        capPara.Append(new BookmarkStart { Id = goBackId, Name = "_GoBack" });
        capPara.Append(new BookmarkEnd { Id = goBackId });
        body.Append(capPara);
    }

    // =========================================================================
    // 表格解析
    // =========================================================================
    static void ParseTable(Body body, string[] lines, int startIdx, out int endIdx)
    {
        var headers  = ParseTableRow(lines[startIdx].Trim());
        int dataStart = startIdx + 2;
        var rows = new List<string[]>();
        int i = dataStart;
        while (i < lines.Length && lines[i].Trim().StartsWith("|"))
            rows.Add(ParseTableRow(lines[i++].Trim()));

        var table = new Table();
        table.Append(new TableProperties(
            new TableStyle { Val = "TableGrid" },
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableJustification { Val = TableRowAlignmentValues.Center }
        ));

        int colCount = headers.Length;
        int colWidth = 9000 / colCount;
        var grid = new TableGrid();
        for (int c = 0; c < colCount; c++) grid.Append(new GridColumn { Width = colWidth.ToString() });
        table.Append(grid);

        // 表头
        var headerRow = new TableRow();
        headerRow.Append(new TableRowProperties(new TableHeader()));
        foreach (var h in headers)
            headerRow.Append(MakeCell(h.Trim(), colWidth, isHeader: true));
        table.Append(headerRow);

        // 数据行
        foreach (var rowData in rows)
        {
            var dataRow = new TableRow();
            for (int c = 0; c < colCount; c++)
                dataRow.Append(MakeCell(c < rowData.Length ? rowData[c].Trim() : "", colWidth));
            table.Append(dataRow);
        }

        body.Append(table);
        endIdx = i;
    }

    static TableCell MakeCell(string text, int colWidth, bool isHeader = false)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Width = colWidth.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
            ),
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
                    new Text(text)
                )
            )
        );
    }

    static string[] ParseTableRow(string line)
    {
        var inner = line.Trim();
        if (inner.StartsWith("|")) inner = inner.Substring(1);
        if (inner.EndsWith("|"))   inner = inner.Substring(0, inner.Length - 1);
        return inner.Split('|');
    }

    // =========================================================================
    // 辅助类型
    // =========================================================================
    class TextPart { public string Text = ""; public bool IsBold; public bool IsItalic; }
    class ImageInfo  { public string Alt = ""; public string Path = ""; }

    static ImageInfo? ParseImageTag(string line)
    {
        var m = Regex.Match(line, @"!\[(.*?)\]\((.*?)\)");
        return m.Success ? new ImageInfo { Alt = m.Groups[1].Value, Path = m.Groups[2].Value } : null;
    }
}
