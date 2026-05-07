// MD to DOCX Converter v1.2.0
// 图片语法支持：
//   ![[path|171]]          Obsidian Wikilink，171 = 绝对像素宽
//   <img style="zoom:80%"> Typora HTML，80% = 相对原图百分比
//   ![alt](path)           标准 Markdown，自动撑满页宽
// 依赖 NuGet：DocumentFormat.OpenXml, System.CommandLine

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.CommandLine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW  = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A   = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace md2docx;

// =============================================================================
// 单位换算（集中管理，避免魔法数字散落各处）
// =============================================================================
static class OpenXmlUnits
{
    public const double EmusPerInch  = 914_400;
    public const double EmusPerCm    = 360_000;
    public const double EmusPerPt    =  12_700;
    public const double DefaultDpi   =      96;

    /// <summary>按默认 96dpi 将像素转为 EMU</summary>
    public static long PxToEmu(double px) => (long)(px * EmusPerInch / DefaultDpi);

    /// <summary>按实际 DPI 将像素转为 EMU（图片文件内嵌 DPI 时使用）</summary>
    public static long PxToEmu(double px, double dpi) => (long)(px * EmusPerInch / dpi);

    /// <summary>厘米转 EMU</summary>
    public static long CmToEmu(double cm) => (long)(cm * EmusPerCm);

    /// <summary>毫米转 EMU</summary>
    public static long MmToEmu(double mm) => (long)(mm * EmusPerCm / 10);
}

// =============================================================================
// 图片缩放策略
// =============================================================================
enum ScaleMode
{
    FitPage,        // 默认：撑满可用页宽
    FixedPixel,     // Obsidian |171  → 固定像素宽
    Percentage,     // Typora zoom:80% → 相对原图百分比
    FixedMm,        // <img width="30mm"> → 毫米（预留扩展）
}

record ImageScale(ScaleMode Mode, double Value = 0)
{
    public static readonly ImageScale FitPage = new(ScaleMode.FitPage);

    /// <summary>
    /// 根据原图尺寸和策略计算最终 EMU 宽高，并限制在页面内。
    /// </summary>
    public (long cx, long cy) Calculate(
        int origPxW, int origPxH,
        double dpi,
        long maxWidthEmu)
    {
        double ratio = (double)origPxH / origPxW;

        long cx = Mode switch
        {
            ScaleMode.FixedPixel  => OpenXmlUnits.PxToEmu(Value, dpi),
            ScaleMode.Percentage  => (long)(OpenXmlUnits.PxToEmu(origPxW, dpi) * Value),
            ScaleMode.FixedMm     => OpenXmlUnits.MmToEmu(Value),
            _                     => maxWidthEmu,   // FitPage
        };

        cx = Math.Min(cx, maxWidthEmu);  // 不超出页面
        long cy = (long)(cx * ratio);
        return (cx, cy);
    }
}

// =============================================================================
// 数据类
// =============================================================================
class ImageInfo
{
    public string     Alt   = "";
    public string     Path  = "";
    public ImageScale Scale = ImageScale.FitPage;
}

class TextPart
{
    public string Text     = "";
    public bool   IsBold   = false;
    public bool   IsItalic = false;
}

// =============================================================================
// 图片元信息读取（内置 PNG / JPEG，无需额外 NuGet）
// =============================================================================
static class ImageReader
{
    /// <summary>
    /// 读取 PNG 或 JPEG 文件的宽、高、DPI。
    /// 未能识别格式时返回 (0, 0, 96)，调用方须做保护。
    /// </summary>
    public static (int width, int height, double dpi) GetInfo(string path)
    {
        using var fs = File.OpenRead(path);
        Span<byte> hdr = stackalloc byte[12];
        fs.Read(hdr);

        // PNG: 签名 89 50 4E 47 0D 0A 1A 0A
        if (hdr[0] == 0x89 && hdr[1] == 0x50 && hdr[2] == 0x4E && hdr[3] == 0x47)
            return ReadPng(fs);

        // JPEG: 签名 FF D8
        if (hdr[0] == 0xFF && hdr[1] == 0xD8)
            return ReadJpeg(fs);

        return (0, 0, 96);
    }

    // ── PNG ──────────────────────────────────────────────────────────────────
    // IHDR chunk 在偏移 8 处（4字节长度 + 4字节"IHDR" + 4字节宽 + 4字节高）
    // pHYs chunk 含 DPI（可选）
    private static (int, int, double) ReadPng(Stream fs)
    {
        // 已读 12 字节（签名8 + 长度4），还差"IHDR" tag 和数据
        Span<byte> buf = stackalloc byte[8];
        fs.Read(buf);                        // "IHDR" + width(4)
        int w = ReadBigEndianInt32(buf[4..]);
        fs.Read(buf[..4]);                   // height
        int h = ReadBigEndianInt32(buf[..4]);

        // 尝试找 pHYs chunk 获取实际 DPI（单位为像素/单位）
        // pHYs: pxPerUnitX(4) + pxPerUnitY(4) + unit(1), unit=1 表示 pixels/metre
        double dpi = 96;
        try
        {
            // 跳过 IHDR 剩余部分（5字节色彩信息 + 4字节CRC）
            fs.Seek(5 + 4, SeekOrigin.Current);
            Span<byte> chunk = stackalloc byte[12];
            while (fs.Read(chunk) == 12)
            {
                int len  = ReadBigEndianInt32(chunk[..4]);
                string t = System.Text.Encoding.ASCII.GetString(chunk[4..8]);
                if (t == "pHYs" && len == 9)
                {
                    Span<byte> phys = stackalloc byte[9];
                    fs.Read(phys);
                    int ppmX = ReadBigEndianInt32(phys[..4]);
                    byte unit = phys[8];
                    if (unit == 1 && ppmX > 0)
                        dpi = ppmX / 39.3701;   // pixels/metre → DPI
                    break;
                }
                fs.Seek(len + 4, SeekOrigin.Current); // 跳到下一 chunk
            }
        }
        catch { /* pHYs 读取失败不影响主流程 */ }

        return (w, h, dpi > 0 ? dpi : 96);
    }

    // ── JPEG ─────────────────────────────────────────────────────────────────
    // 扫描 APP0(JFIF/JFXX) 或 SOF0/SOF1/SOF2 段获取尺寸和 DPI
    private static (int, int, double) ReadJpeg(Stream fs)
    {
        int w = 0, h = 0;
        double dpi = 96;
        Span<byte> marker = stackalloc byte[2];

        while (fs.Read(marker) == 2)
        {
            if (marker[0] != 0xFF) break;
            byte tag = marker[1];

            // APP0 (JFIF) → 可获取 DPI
            if (tag == 0xE0)
            {
                Span<byte> seg = stackalloc byte[16];
                fs.Read(seg);
                // seg[0..1]=length, [2..5]="JFIF", [6]=unit, [7..8]=Xdensity, [9..10]=Ydensity
                if (seg[2] == 'J' && seg[3] == 'F' && seg[4] == 'I' && seg[5] == 'F')
                {
                    byte unit = seg[6];
                    int xd = (seg[7] << 8) | seg[8];
                    if (unit == 1 && xd > 0) dpi = xd;  // dots/inch
                    if (unit == 2 && xd > 0) dpi = xd * 2.54;  // dots/cm
                }
                int len = ((seg[0] << 8) | seg[1]) - 16;
                if (len > 0) fs.Seek(len, SeekOrigin.Current);
                continue;
            }

            // SOF 段（包含宽高）: SOF0=C0, SOF1=C1, SOF2=C2
            if (tag is 0xC0 or 0xC1 or 0xC2)
            {
                Span<byte> sof = stackalloc byte[8];
                fs.Read(sof);
                h = (sof[3] << 8) | sof[4];
                w = (sof[5] << 8) | sof[6];
                break;
            }

            // 其他段：跳过
            Span<byte> lenBuf = stackalloc byte[2];
            if (fs.Read(lenBuf) < 2) break;
            int skip = ((lenBuf[0] << 8) | lenBuf[1]) - 2;
            if (skip > 0) fs.Seek(skip, SeekOrigin.Current);
        }

        return (w, h, dpi);
    }

    private static int ReadBigEndianInt32(ReadOnlySpan<byte> b)
        => (b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3];
}

// =============================================================================
// 主程序
// =============================================================================
class Program
{
    private const int    A4W     = 11906;
    private const int    A4H     = 16838;
    private const string BLACK   = "000000";
    private const string VERSION = "1.2.0";

    // 页面可用宽度 = A4 宽 - 左右边距(各 1800 twip = 3.17cm)
    // = 11906 - 3600 = 8306 twip → EMU = 8306 * 635 ≈ 5_274_310
    // 保守取 15cm = 5_400_000 EMU（与原版一致）
    private static readonly long PageWidthEmu = OpenXmlUnits.CmToEmu(15);

    private static uint _docPrId    = 1;
    private static int  _figCounter = 0;
    private static int  _tblCounter = 0;
    private static int  _bmkId      = 0;

    // =========================================================================
    // 入口
    // =========================================================================
    static int Main(string[] args)
    {
        var inputArg = new Argument<string?>(
            "input.md",
            () => null,
            "Input Markdown file. Omit to read from stdin.");

        var outputOpt = new Option<string?>(
            new[] { "-o", "--output" },
            "Output .docx path.");

        var baseDirOpt = new Option<string?>(
            "--base-dir",
            "Base directory for resolving image paths.");

        var forceOpt = new Option<bool>(
            new[] { "-f", "--force" },
            "Overwrite output file if it already exists.");

        var root = new RootCommand($"md2docx v{VERSION} - Markdown to DOCX Converter")
            { inputArg, outputOpt, baseDirOpt, forceOpt };

        root.SetHandler(Run, inputArg, outputOpt, baseDirOpt, forceOpt);
        return root.Invoke(args);
    }

    static void Run(string? inputFile, string? outputOpt, string? baseDirOpt, bool force)
    {
        string mdText, baseDir, desiredOut;

        if (inputFile == null)
        {
            if (!Console.IsInputRedirected)
                Console.Error.WriteLine("Reading from stdin (end with Ctrl+D / Ctrl+Z):");
            mdText     = Console.In.ReadToEnd();
            baseDir    = baseDirOpt ?? Directory.GetCurrentDirectory();
            desiredOut = outputOpt  ?? Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        }
        else
        {
            var fullIn = Path.GetFullPath(inputFile);
            if (!File.Exists(fullIn)) Die($"Input file not found: {fullIn}");
            mdText     = File.ReadAllText(fullIn);
            baseDir    = baseDirOpt != null ? Path.GetFullPath(baseDirOpt) : Path.GetDirectoryName(fullIn)!;
            desiredOut = outputOpt  ?? Path.ChangeExtension(fullIn, ".docx");
        }

        if (!Directory.Exists(baseDir)) Die($"Base directory not found: {baseDir}");

        var outPath = ResolveOutput(Path.GetFullPath(desiredOut), force);

        using var doc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        AddStyles(mainPart);
        mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

        ParseMarkdown(body, mainPart, mdText, baseDir);

        body.Append(new SectionProperties(
            new PageSize   { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 720, Footer = 720 }
        ));

        doc.Save();
        var info = new FileInfo(outPath);
        Console.WriteLine($"Generated : {outPath}");
        Console.WriteLine($"Size      : {info.Length:N0} bytes");
        if (desiredOut != outPath)
            Console.WriteLine($"Note      : saved as '{Path.GetFileName(outPath)}' (name conflict)");
    }

    static string ResolveOutput(string desired, bool force)
    {
        if (force || !File.Exists(desired)) return desired;
        var dir = Path.GetDirectoryName(desired) ?? ".";
        var stem = Path.GetFileNameWithoutExtension(desired);
        var ext  = Path.GetExtension(desired);
        for (int n = 1; ; n++)
        {
            var c = Path.Combine(dir, $"{stem}_{n}{ext}");
            if (!File.Exists(c)) return c;
        }
    }

    static void Die(string msg, int code = 1)
    {
        Console.Error.WriteLine($"Error: {msg}");
        Environment.Exit(code);
    }

    // =========================================================================
    // Markdown 解析器
    // =========================================================================
    static void ParseMarkdown(Body body, MainDocumentPart mainPart, string md, string baseDir)
    {
        var lines = md.Split('\n');
        int i = 0;

        while (i < lines.Length)
        {
            var line    = lines[i];
            var trimmed = line.TrimStart();

            if (string.IsNullOrWhiteSpace(line)) { i++; continue; }

            // ── 标题 ──────────────────────────────────────────────────────
            if (trimmed.StartsWith("#"))
            {
                int level = 0;
                while (level < trimmed.Length && trimmed[level] == '#') level++;
                if (level <= 5 && trimmed.Length > level && trimmed[level] == ' ')
                {
                    body.Append(CreateHeading(trimmed[(level + 1)..].Trim(), level));
                    i++;
                    continue;
                }
            }

            // ── 图片（Obsidian / Typora / 标准 Markdown）─────────────────
            var imgInfo = ParseImage(trimmed);
            if (imgInfo != null)
            {
                // 将相对路径 resolve 到 baseDir；绝对路径保持原样
                string resolvedPath = Path.IsPathRooted(imgInfo.Path)
                    ? imgInfo.Path
                    : Path.Combine(baseDir, imgInfo.Path);
                AddImageWithCaption(body, mainPart, resolvedPath, imgInfo.Alt, imgInfo.Scale);
                i++;
                continue;
            }

            // ── 表格 ──────────────────────────────────────────────────────
            if (trimmed.StartsWith("|") && i + 1 < lines.Length && lines[i + 1].Trim().Contains("|-"))
            {
                ParseTable(body, lines, i, out int endIdx);
                i = endIdx;

                string caption = "";
                if (i < lines.Length)
                {
                    var capM = Regex.Match(lines[i].Trim(), @"^\[(.+)\]$");
                    if (capM.Success) { caption = capM.Groups[1].Value; i++; }
                }
                AppendTableCaption(body, caption);
                continue;
            }

            // ── 普通段落（合并连续非空行）────────────────────────────────
            var paraText = line.Trim();
            i++;
            while (i < lines.Length
                   && !string.IsNullOrWhiteSpace(lines[i])
                   && !lines[i].TrimStart().StartsWith("#")
                   && !lines[i].TrimStart().StartsWith("|")
                   && ParseImage(lines[i].TrimStart()) == null)
            {
                paraText += lines[i].Trim();
                i++;
            }
            body.Append(CreateRichParagraph(paraText));
        }
    }

    // =========================================================================
    // 图片语法统一解析
    // =========================================================================
    /// <summary>
    /// 支持三种语法，返回 null 表示当前行不是图片行。
    /// 优先级：Obsidian Wikilink > Typora HTML img > 标准 Markdown
    /// </summary>
    static ImageInfo? ParseImage(string line)
    {
        // ── Obsidian: ![[path]] 或 ![[path|width_px]] ────────────────────────
        if (line.StartsWith("![["))
        {
            var m = Regex.Match(line, @"!\[\[([^\|\]\n]+?)(?:\|(\d+))?\]\]");
            if (m.Success)
            {
                var scale = m.Groups[2].Success
                    ? new ImageScale(ScaleMode.FixedPixel, double.Parse(m.Groups[2].Value))
                    : ImageScale.FitPage;
                return new ImageInfo { Path = m.Groups[1].Value.Trim(), Scale = scale };
            }
        }

        // ── Typora HTML: <img src="..." style="zoom:N%;" /> ──────────────────
        if (line.StartsWith("<img", StringComparison.OrdinalIgnoreCase))
        {
            var srcM = Regex.Match(line, @"src=""([^""]+)""", RegexOptions.IgnoreCase);
            if (srcM.Success)
            {
                var altM  = Regex.Match(line, @"alt=""([^""]*)""",      RegexOptions.IgnoreCase);
                var zoomM = Regex.Match(line, @"zoom:\s*([\d.]+)%",     RegexOptions.IgnoreCase);
                var widM  = Regex.Match(line, @"width=""(\d+)""",       RegexOptions.IgnoreCase);
                var widMm = Regex.Match(line, @"width=""(\d+)mm""",    RegexOptions.IgnoreCase);

                ImageScale scale;
                if (zoomM.Success)
                    scale = new ImageScale(ScaleMode.Percentage, double.Parse(zoomM.Groups[1].Value) / 100.0);
                else if (widMm.Success)
                    scale = new ImageScale(ScaleMode.FixedMm, double.Parse(widMm.Groups[1].Value));
                else if (widM.Success)
                    scale = new ImageScale(ScaleMode.FixedPixel, double.Parse(widM.Groups[1].Value));
                else
                    scale = ImageScale.FitPage;

                return new ImageInfo
                {
                    Path  = srcM.Groups[1].Value,
                    Alt   = altM.Success ? altM.Groups[1].Value : "",
                    Scale = scale,
                };
            }
        }

        // ── 标准 Markdown: ![alt](path) ──────────────────────────────────────
        if (line.StartsWith("!["))
        {
            var m = Regex.Match(line, @"!\[(.*?)\]\((.*?)\)");
            if (m.Success)
                return new ImageInfo { Alt = m.Groups[1].Value, Path = m.Groups[2].Value };
        }

        return null;
    }

    // =========================================================================
    // 添加图片 + 题注
    // =========================================================================
    static void AddImageWithCaption(
        Body body, MainDocumentPart mainPart,
        string imagePath, string caption, ImageScale scale)
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

        // 读取图片字节和元信息
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        var (imgW, imgH, dpi) = ImageReader.GetInfo(imagePath);

        // 容错：若读取失败则用保守默认值
        if (imgW <= 0 || imgH <= 0) { imgW = 800; imgH = 600; }
        if (dpi   <= 0)               dpi = 96;

        // 计算 EMU 尺寸
        var (cx, cy) = scale.Calculate(imgW, imgH, dpi, PageWidthEmu);

        // 上传图片到文档
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(imageBytes)) imagePart.FeedData(ms);
        var imageId = mainPart.GetIdOfPart(imagePart);

        uint prId       = _docPrId++;
        int  borderEmu  = 9525; // 0.75pt 边框

        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0, Name = $"fig{_figCounter}{Path.GetExtension(imagePath)}" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(new A.Blip { Embed = imageId }, new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle },
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = BLACK })) { Width = borderEmu }
            )
        );

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = borderEmu, TopEdge = borderEmu, RightEdge = borderEmu, BottomEdge = borderEmu },
            new DW.DocProperties { Id = prId, Name = $"Fig{_figCounter}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(picture)
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        ) { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 };

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }
            ),
            new Run(new Drawing(inline))));

        // 图题（SEQ 域自动编号）
        var capPara = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }));
        AppendCaptionField(capPara, "图", _figCounter.ToString(), caption);
        body.Append(capPara);
    }

    // =========================================================================
    // 表格题注
    // =========================================================================
    static void AppendTableCaption(Body body, string captionText)
    {
        _tblCounter++;
        var capPara = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }));
        AppendCaptionField(capPara, "表", _tblCounter.ToString(), captionText);
        body.Append(capPara);
    }

    /// <summary>向题注段落追加「图/表 N captionText」，含 SEQ 域</summary>
    static void AppendCaptionField(Paragraph p, string seqLabel, string seqNum, string captionText)
    {
        var hei = new RunProperties(
            new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" });

        p.Append(new Run(hei.CloneNode(true) as RunProperties ?? hei,
            new Text($"{seqLabel} ") { Space = SpaceProcessingModeValues.Preserve }));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        p.Append(new Run(new FieldCode($" SEQ {seqLabel} \\* ARABIC ")
            { Space = SpaceProcessingModeValues.Preserve }));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        p.Append(new Run(new Text(seqNum)));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        if (!string.IsNullOrWhiteSpace(captionText))
            p.Append(new Run(hei.CloneNode(true) as RunProperties ?? hei,
                new Text($" {captionText}") { Space = SpaceProcessingModeValues.Preserve }));

        var bmkId = (++_bmkId).ToString();
        p.Append(new BookmarkStart { Id = bmkId, Name = "_GoBack" });
        p.Append(new BookmarkEnd   { Id = bmkId });
    }

    // =========================================================================
    // 段落 / 标题
    // =========================================================================
    static Paragraph CreateHeading(string text, int level)
    {
        var bmkId = (++_bmkId).ToString();
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{level}" }),
            new BookmarkStart { Id = bmkId, Name = $"_Toc{_bmkId:D3}" },
            new Run(new Text(text)),
            new BookmarkEnd { Id = bmkId }
        );
    }

    static Paragraph CreateRichParagraph(string text)
    {
        var para = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }));
        foreach (var part in ParseInline(text))
        {
            var rpr = new RunProperties(
                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" },
                new Color    { Val = BLACK });
            if (part.IsBold)   rpr.Append(new Bold());
            if (part.IsItalic) rpr.Append(new Italic());
            para.Append(new Run(rpr,
                new Text(part.Text) { Space = SpaceProcessingModeValues.Preserve }));
        }
        return para;
    }

    static List<TextPart> ParseInline(string text)
    {
        var parts = new List<TextPart>();
        int i = 0;
        while (i < text.Length)
        {
            // **bold**
            if (i + 1 < text.Length && text[i] == '*' && text[i + 1] == '*')
            {
                int end = text.IndexOf("**", i + 2);
                if (end > 0)
                {
                    parts.Add(new TextPart { Text = text[(i + 2)..end], IsBold = true });
                    i = end + 2; continue;
                }
            }
            // *italic*
            if (text[i] == '*')
            {
                int end = text.IndexOf('*', i + 1);
                if (end > 0)
                {
                    parts.Add(new TextPart { Text = text[(i + 1)..end], IsItalic = true });
                    i = end + 1; continue;
                }
            }
            // plain text until next special
            int next = int.MaxValue;
            int bp   = text.IndexOf("**", i);
            int ip   = text.IndexOf('*',  i);
            if (bp >= 0) next = Math.Min(next, bp);
            if (ip >= 0) next = Math.Min(next, ip);
            if (next == int.MaxValue) { parts.Add(new TextPart { Text = text[i..] }); break; }
            parts.Add(new TextPart { Text = text[i..next] });
            i = next;
        }
        if (parts.Count == 0) parts.Add(new TextPart { Text = text });
        return parts;
    }

    // =========================================================================
    // 表格
    // =========================================================================
    static void ParseTable(Body body, string[] lines, int startIdx, out int endIdx)
    {
        var headers = SplitRow(lines[startIdx].Trim());
        var rows    = new List<string[]>();
        int i       = startIdx + 2;
        while (i < lines.Length && lines[i].Trim().StartsWith("|"))
            rows.Add(SplitRow(lines[i++].Trim()));

        int colCount = headers.Length;
        int colWidth = 9000 / colCount;

        var table = new Table();
        table.Append(new TableProperties(
            new TableStyle { Val = "TableGrid" },
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableJustification { Val = TableRowAlignmentValues.Center }
        ));

        var grid = new TableGrid();
        for (int c = 0; c < colCount; c++) grid.Append(new GridColumn { Width = colWidth.ToString() });
        table.Append(grid);

        var hRow = new TableRow();
        hRow.Append(new TableRowProperties(new TableHeader()));
        foreach (var h in headers) hRow.Append(MakeCell(h.Trim(), colWidth));
        table.Append(hRow);

        foreach (var row in rows)
        {
            var dRow = new TableRow();
            for (int c = 0; c < colCount; c++)
                dRow.Append(MakeCell(c < row.Length ? row[c].Trim() : "", colWidth));
            table.Append(dRow);
        }

        body.Append(table);
        endIdx = i;
    }

    static TableCell MakeCell(string text, int colWidth)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Width = colWidth.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new Indentation   { FirstLine = "0" }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                        new FontSize { Val = "28" },
                        new Color    { Val = BLACK }),
                    new Text(text))));
    }

    static string[] SplitRow(string line)
    {
        var s = line.Trim();
        if (s.StartsWith("|")) s = s[1..];
        if (s.EndsWith("|"))   s = s[..^1];
        return s.Split('|');
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

        foreach (var (id, name, level) in new[]
        {
            ("Heading1","heading 1",0),("Heading2","heading 2",1),
            ("Heading3","heading 3",2),("Heading4","heading 4",3),
            ("Heading5","heading 5",4),
        })
        {
            sp.Styles.Append(new Style(
                new StyleName { Val = name },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new StyleParagraphProperties(
                    new KeepNext(), new KeepLines(),
                    new SpacingBetweenLines { Before = "0", After = "0", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                    new OutlineLevel { Val = level }),
                new StyleRunProperties(
                    new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                    new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" },
                    new Color    { Val = BLACK })
            ) { Type = StyleValues.Paragraph, StyleId = id });
        }

        sp.Styles.Append(new Style(
            new StyleName { Val = "Body Text" },
            new Aliases { Val = "正文" },
            new StyleParagraphProperties(
                new WidowControl { Val = false },
                new WordWrap { Val = false },
                new Indentation { FirstLineChars = 200, FirstLine = "200" },
                new Justification { Val = JustificationValues.Both },
                new SpacingBetweenLines { Before = "0", After = "0" }),
            new StyleRunProperties(
                new RunFonts { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" },
                new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" },
                new Color    { Val = BLACK })
        ) { Type = StyleValues.Paragraph, StyleId = "BodyText" });

        sp.Styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" }),
            new StyleRunProperties(
                new RunFonts { Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体", ComplexScript = "黑体" },
                new FontSize { Val = "20" },
                new Color    { Val = BLACK })
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
                    new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4, Color = "000000" }),
                new TableCellMarginDefault(
                    new TopMargin    { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new StartMargin  { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new EndMargin    { Width = "80", Type = TableWidthUnitValues.Dxa }))
        ) { Type = StyleValues.Table, StyleId = "TableGrid" });
    }
}
