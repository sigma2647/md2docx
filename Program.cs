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
// 修复 CA2022：所有 Read() 改为 ReadExactly()，保证读满缓冲区
// 修复 CA2014：stackalloc 全部移到循环体外，消除栈溢出风险
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

        // CA2014 fix：stackalloc 在方法顶部声明，不在任何循环内
        Span<byte> hdr = stackalloc byte[12];
        fs.ReadExactly(hdr);   // CA2022 fix：ReadExactly 保证读满

        // PNG: 签名 89 50 4E 47 0D 0A 1A 0A
        if (hdr[0] == 0x89 && hdr[1] == 0x50 && hdr[2] == 0x4E && hdr[3] == 0x47)
            return ReadPng(fs);

        // JPEG: 签名 FF D8
        if (hdr[0] == 0xFF && hdr[1] == 0xD8)
            return ReadJpeg(fs);

        return (0, 0, 96);
    }

    // ── PNG ──────────────────────────────────────────────────────────────────
    // 文件布局：[8字节签名][4字节IHDR长度][4字节"IHDR"][4字节宽][4字节高][5字节色彩][4字节CRC]...
    // GetInfo 已读了签名(8) + IHDR长度(4) = 12字节，fs 当前位置在"IHDR"标签处
    private static (int, int, double) ReadPng(Stream fs)
    {
        // CA2014 fix：所有 stackalloc 在方法入口一次性声明
        Span<byte> buf8  = stackalloc byte[8];   // 通用 8 字节缓冲
        Span<byte> chunk = stackalloc byte[12];  // chunk 头：length(4)+type(4)+前4字节数据
        Span<byte> phys  = stackalloc byte[9];   // pHYs 数据：ppmX(4)+ppmY(4)+unit(1)

        // 读 IHDR 标签(4) + 宽(4)
        fs.ReadExactly(buf8);
        int w = ReadBE32(buf8[4..]);

        // 读 高(4)
        fs.ReadExactly(buf8[..4]);
        int h = ReadBE32(buf8[..4]);

        double dpi = 96;
        try
        {
            // 跳过 IHDR 剩余：色彩信息(5字节) + CRC(4字节)
            fs.Seek(5 + 4, SeekOrigin.Current);

            // 逐 chunk 扫描，找 pHYs（含实际 DPI）
            // pHYs 在图片中出现顺序不定，但一般紧跟 IHDR
            while (true)
            {
                // CA2022 fix：改用 Read() 的返回值判断 EOF，因为 ReadExactly 在 EOF 会抛异常
                int got = fs.Read(chunk);
                if (got < 12) break;

                int    len  = ReadBE32(chunk[..4]);
                string type = System.Text.Encoding.ASCII.GetString(chunk[4..8]);

                if (type == "IEND") break;  // 文件结束

                if (type == "pHYs" && len == 9)
                {
                    // CA2022 fix：ReadExactly 保证读满 9 字节
                    fs.ReadExactly(phys);
                    int  ppmX = ReadBE32(phys[..4]);
                    byte unit = phys[8];
                    // unit=1: pixels/metre → inch = metre/0.0254 → DPI = ppm/39.3701
                    if (unit == 1 && ppmX > 0)
                        dpi = ppmX / 39.3701;
                    break;
                }

                // 跳过本 chunk 的数据(len) + CRC(4)
                fs.Seek(len + 4, SeekOrigin.Current);
            }
        }
        catch { /* pHYs 读取失败不影响主流程，保持 dpi=96 */ }

        return (w, h, dpi > 0 ? dpi : 96);
    }

    // ── JPEG ─────────────────────────────────────────────────────────────────
    // 扫描 APP0(JFIF) 获取 DPI，扫描 SOF0/SOF1/SOF2 段获取宽高
    // GetInfo 已读了 12 字节签名，fs 当前位置在第 13 字节（SOI 之后）
    private static (int, int, double) ReadJpeg(Stream fs)
    {
        int    w = 0, h = 0;
        double dpi = 96;

        // CA2014 fix：所有 stackalloc 在循环外声明
        Span<byte> marker = stackalloc byte[2];   // 段标记：FF xx
        Span<byte> lenBuf = stackalloc byte[2];   // 段长度（不含标记本身）
        Span<byte> seg16  = stackalloc byte[16];  // APP0 前 16 字节
        Span<byte> sof8   = stackalloc byte[8];   // SOF 前 8 字节（含宽高）

        // 回到文件起点重新扫描（GetInfo 已读了 12 字节，JPEG 段在 SOI 之后）
        fs.Seek(2, SeekOrigin.Begin);   // 跳过 FF D8 (SOI)

        while (true)
        {
            // CA2022 fix：用 Read 检查 EOF，不用 ReadExactly（可能正好到文件尾）
            if (fs.Read(marker) < 2) break;
            if (marker[0] != 0xFF) break;   // 非法格式
            byte tag = marker[1];

            // FF FF / FF 00 是填充字节，跳过
            if (tag == 0xFF || tag == 0x00) continue;

            // SOI / EOI 没有数据段
            if (tag == 0xD8 || tag == 0xD9) continue;

            // ── APP0 (JFIF)：含 DPI ──────────────────────────────────────
            if (tag == 0xE0)
            {
                if (fs.Read(seg16) < 16) break;
                // seg16[0..1]=段总长, [2..5]="JFIF\0", [6]=unit, [7..8]=Xdensity
                if (seg16[2] == 'J' && seg16[3] == 'F' && seg16[4] == 'I' && seg16[5] == 'F')
                {
                    byte unit = seg16[6];
                    int  xd   = (seg16[7] << 8) | seg16[8];
                    if (unit == 1 && xd > 0) dpi = xd;           // dots/inch
                    if (unit == 2 && xd > 0) dpi = xd * 2.54;    // dots/cm → dpi
                }
                // 跳过段剩余部分（段总长包含自身 2 字节，已读 16 字节）
                int remain = ((seg16[0] << 8) | seg16[1]) - 16;
                if (remain > 0) fs.Seek(remain, SeekOrigin.Current);
                continue;
            }

            // ── SOF0 / SOF1 / SOF2：含宽高 ──────────────────────────────
            if (tag is 0xC0 or 0xC1 or 0xC2)
            {
                if (fs.Read(sof8) < 8) break;
                // sof8[0..1]=段长, [2]=精度, [3..4]=高, [5..6]=宽
                h = (sof8[3] << 8) | sof8[4];
                w = (sof8[5] << 8) | sof8[6];
                break;  // 宽高找到，不需要继续扫描
            }

            // ── 其他段：读段长跳过 ───────────────────────────────────────
            if (fs.Read(lenBuf) < 2) break;
            int skip = ((lenBuf[0] << 8) | lenBuf[1]) - 2;  // 段长含自身 2 字节
            if (skip > 0) fs.Seek(skip, SeekOrigin.Current);
        }

        return (w, h, dpi);
    }

    // 大端序 int32 读取
    private static int ReadBE32(ReadOnlySpan<byte> b)
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
        var inputArg = new Argument<string?>("input.md")
        {
            Description = "Input Markdown file. Omit to read from stdin.",
            Arity = ArgumentArity.ZeroOrOne
        };

        var outputOpt = new Option<string?>("--output", "-o")
        {
            Description = "Output .docx path."
        };

        var baseDirOpt = new Option<string?>("--base-dir")
        {
            Description = "Base directory for resolving image paths."
        };

        var forceOpt = new Option<bool>("--force", "-f")
        {
            Description = "Overwrite output file if it already exists."
        };

        var root = new RootCommand($"md2docx v{VERSION} - Markdown to DOCX Converter")
            { inputArg, outputOpt, baseDirOpt, forceOpt };

        root.SetAction(parseResult =>
        {
            var inputFile  = parseResult.GetValue(inputArg);
            var outOpt     = parseResult.GetValue(outputOpt);
            var baseDir    = parseResult.GetValue(baseDirOpt);
            var force      = parseResult.GetValue(forceOpt);
            Run(inputFile, outOpt, baseDir, force);
            return 0;
        });

        return root.Parse(args).Invoke();
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
            new PageMargin { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 720, Footer = 720 },
            new Columns    { Space = "425" },
            // 关键：中文文档必须有 docGrid，否则 WPS 渲染时段前段后不对称
            // （依赖字体 ascent/descent，黑体不对称会"塌顶"）。linePitch=312 对应单倍行距网格。
            new DocGrid    { Type = DocGridValues.Lines, LinePitch = 312 }
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
                var widMm = Regex.Match(line, @"width=""(\d+)mm""",     RegexOptions.IgnoreCase);

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
