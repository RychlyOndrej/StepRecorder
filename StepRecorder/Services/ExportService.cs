using System.IO;
using System.Net;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using StepRecorder.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WpDocument = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace StepRecorder.Services;

/// <summary>Exports a completed recording session to PDF, MHT and/or Word.</summary>
public sealed class ExportService
{
    // ── Main entry ────────────────────────────────────────────────────────

    public async Task<List<string>> ExportAsync(
        RecordingSession session, AppSettings cfg,
        IProgress<(int current, int total, string label)>? progress = null)
    {
        var exported = new List<string>();
        int total    = CountFormats(cfg.ExportFormats);
        int done     = 0;

        string dir = !string.IsNullOrWhiteSpace(session.SessionFolder)
            ? session.SessionFolder
            : GetUniqueDirectoryPath(Path.Combine(cfg.OutputFolder, SanitizeFileName(session.Name)));
        Directory.CreateDirectory(dir);

        if (cfg.ExportFormats.HasFlag(ExportFormatFlags.PDF))
        {
            progress?.Report((done, total, "Generuji PDF…"));
            var p = await ExportPdfAsync(session, cfg, dir);
            if (p != null) exported.Add(p);
            done++;
        }

        if (cfg.ExportFormats.HasFlag(ExportFormatFlags.MHT))
        {
            progress?.Report((done, total, "Generuji MHT…"));
            var p = await ExportMhtAsync(session, cfg, dir);
            if (p != null) exported.Add(p);
            done++;
        }

        if (cfg.ExportFormats.HasFlag(ExportFormatFlags.Word))
        {
            progress?.Report((done, total, "Generuji Word…"));
            var p = await ExportWordAsync(session, cfg, dir);
            if (p != null) exported.Add(p);
            done++;
        }

        progress?.Report((done, total, "Hotovo"));
        return exported;
    }

    // ── PDF ───────────────────────────────────────────────────────────────

    private static Task<string?> ExportPdfAsync(
        RecordingSession session, AppSettings cfg, string outDir)
    {
        return Task.Run<string?>(() =>
        {
            try
            {
                var path = Path.Combine(outDir, SanitizeFileName(session.Name) + ".pdf");

                QuestPDF.Fluent.Document.Create(c =>
                {
                    c.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(35);
                        page.DefaultTextStyle(t =>
                            t.FontFamily("Segoe UI").FontSize(10).FontColor(Colors.Grey.Darken3));

                        // ─ Header ─
                        page.Header().Column(col =>
                        {
                            col.Item()
                               .Text(session.Name)
                               .FontSize(20).Bold().FontColor("#0078D4");

                            col.Item()
                               .Text($"Vytvořeno: {session.StartTime:dd.MM.yyyy HH:mm}" +
                                     $"  •  Kroků: {session.StepCount}" +
                                     $"  •  Trvání: {session.Duration:mm\\:ss}" +
                                     $"  •  Autor: {cfg.DocumentAuthor}")
                               .FontSize(9).FontColor(Colors.Grey.Medium);

                            col.Item().LineHorizontal(1.5f).LineColor("#0078D4");
                            col.Item().Height(8);
                        });

                        // ─ Content ─
                        page.Content().Column(col =>
                        {
                            foreach (var step in session.Steps)
                            {
                                col.Item()
                                   .PaddingBottom(12)
                                   .Border(1).BorderColor(Colors.Grey.Lighten2)
                                   .CornerRadius(4)
                                   .Column(card =>
                                   {
                                       // Step header bar
                                       card.Item()
                                           .Background("#0078D4")
                                           .Padding(8)
                                           .Row(row =>
                                           {
                                               row.ConstantItem(28).Height(28)
                                                  .Background(Colors.White)
                                                  .CornerRadius(14)
                                                  .AlignCenter().AlignMiddle()
                                                  .Text(step.StepNumber.ToString())
                                                  .FontSize(12).Bold().FontColor("#0078D4");

                                               row.RelativeItem().PaddingLeft(8).Column(inner =>
                                               {
                                                   inner.Item().Text(step.DisplayName)
                                                        .FontSize(11).Bold().FontColor(Colors.White);

                                                   if (cfg.IncludeTimestamps)
                                                       inner.Item()
                                                            .Text($"{step.TimeDisplay}  {step.SourceIcon}")
                                                            .FontSize(8).FontColor("#B3D9FF");
                                               });
                                           });

                                       // Screenshot
                                       if (step.ImagePath != null && File.Exists(step.ImagePath))
                                       {
                                           card.Item()
                                               .Padding(8)
                                               .Image(step.ImagePath)
                                               .FitWidth();
                                       }

                                       // Metadata footer
                                       card.Item()
                                           .Background(Colors.Grey.Lighten4)
                                           .Padding(8)
                                           .Column(meta =>
                                           {
                                               if (cfg.IncludeWindowTitles && !string.IsNullOrEmpty(step.WindowTitle))
                                                   meta.Item().Text($"🪟  {step.WindowTitle}  [{step.ProcessName}]")
                                                       .FontSize(9).Italic().FontColor(Colors.Grey.Darken1);

                                               if (!string.IsNullOrEmpty(step.KeysPressed))
                                                   meta.Item().Text($"⌨  {step.KeysPressed}")
                                                       .FontSize(9).FontColor(Colors.Grey.Darken2);

                                               if (!string.IsNullOrEmpty(step.Description))
                                               {
                                                   meta.Item().Height(4);
                                                   meta.Item()
                                                       .Background("#EFF6FC")
                                                       .Padding(6)
                                                       .Text(step.Description)
                                                       .FontSize(10).FontColor(Colors.Grey.Darken3);
                                               }
                                           });
                                   });
                            }
                        });

                        // ─ Footer ─
                        page.Footer()
                            .AlignCenter()
                            .Text(x =>
                            {
                                x.Span("StepRecorder  •  Strana ").FontSize(8).FontColor(Colors.Grey.Medium);
                                x.CurrentPageNumber().FontSize(8).FontColor(Colors.Grey.Medium);
                                x.Span(" / ").FontSize(8).FontColor(Colors.Grey.Medium);
                                x.TotalPages().FontSize(8).FontColor(Colors.Grey.Medium);
                            });
                    });
                }).GeneratePdf(path);

                return path;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PDF] {ex.Message}");
                return null;
            }
        });
    }

    // ── MHT ───────────────────────────────────────────────────────────────

    private static async Task<string?> ExportMhtAsync(
        RecordingSession session, AppSettings cfg, string outDir)
    {
        try
        {
            var path     = Path.Combine(outDir, SanitizeFileName(session.Name) + ".mht");
            var boundary = "----=_NextPart_" + Guid.NewGuid().ToString("N");
            var sb       = new StringBuilder();

            sb.AppendLine("MIME-Version: 1.0");
            sb.AppendLine($"Content-Type: multipart/related; boundary=\"{boundary}\"; type=\"text/html\"");
            sb.AppendLine($"Subject: {session.Name}");
            sb.AppendLine($"Date: {session.StartTime:R}");
            sb.AppendLine();

            // HTML part
            sb.AppendLine($"--{boundary}");
            sb.AppendLine("Content-Type: text/html; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: base64");
            sb.AppendLine("Content-ID: <main@sr>");
            sb.AppendLine();

            string html = BuildHtml(session, cfg, embedCid: true);
            string htmlB64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(html));
            for (int i = 0; i < htmlB64.Length; i += 76)
                sb.AppendLine(htmlB64.Substring(i, Math.Min(76, htmlB64.Length - i)));
            sb.AppendLine();

            // Image parts
            foreach (var step in session.Steps)
            {
                if (step.ImagePath == null || !File.Exists(step.ImagePath)) continue;

                var bytes   = await File.ReadAllBytesAsync(step.ImagePath);
                var b64     = Convert.ToBase64String(bytes);
                var mime    = step.ImagePath.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                              step.ImagePath.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                            ? "image/jpeg"
                            : "image/png";
                var cid     = $"img{step.StepNumber}@sr";

                sb.AppendLine($"--{boundary}");
                sb.AppendLine($"Content-Type: {mime}");
                sb.AppendLine("Content-Transfer-Encoding: base64");
                sb.AppendLine($"Content-ID: <{cid}>");
                sb.AppendLine($"Content-Location: images/step{step.StepNumber:000}");
                sb.AppendLine();

                for (int i = 0; i < b64.Length; i += 76)
                    sb.AppendLine(b64.Substring(i, Math.Min(76, b64.Length - i)));
                sb.AppendLine();
            }

            sb.AppendLine($"--{boundary}--");

            await File.WriteAllTextAsync(path, sb.ToString(), Encoding.UTF8);
            return path;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MHT] {ex.Message}");
            return null;
        }
    }

    // ── Word ──────────────────────────────────────────────────────────────

    private static Task<string?> ExportWordAsync(
        RecordingSession session, AppSettings cfg, string outDir)
    {
        return Task.Run<string?>(() =>
        {
            try
            {
                var path = Path.Combine(outDir, SanitizeFileName(session.Name) + ".docx");

                using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new WpDocument(new Body());
                var body = mainPart.Document.Body!;

                AddParagraph(body, session.Name, bold: true, fontSizeHalfPoints: 32);
                AddParagraph(body,
                    $"Datum: {session.StartTime:dd.MM.yyyy HH:mm}   Kroků: {session.StepCount}   Trvání: {session.Duration:mm\\:ss}   Autor: {cfg.DocumentAuthor}");
                body.Append(new Paragraph(new Run(new Break())));

                foreach (var step in session.Steps)
                {
                    AddParagraph(body, $"Krok {step.StepNumber}: {step.DisplayName}", bold: true, fontSizeHalfPoints: 26);

                    if (cfg.IncludeTimestamps)
                        AddParagraph(body, $"Čas: {step.TimeDisplay}  {step.SourceIcon}");

                    if (cfg.IncludeWindowTitles && !string.IsNullOrWhiteSpace(step.WindowTitle))
                        AddParagraph(body, $"Okno: {step.WindowTitle} [{step.ProcessName}]");

                    if (!string.IsNullOrWhiteSpace(step.KeysPressed))
                        AddParagraph(body, $"Klávesy: {step.KeysPressed}");

                    if (!string.IsNullOrWhiteSpace(step.Description))
                        AddParagraph(body, $"Popis: {step.Description}");

                    if (!string.IsNullOrWhiteSpace(step.ImagePath) && File.Exists(step.ImagePath))
                        AddImageParagraph(mainPart, body, step.ImagePath);

                    body.Append(new Paragraph(new Run(new Break())));
                }

                mainPart.Document.Save();
                return path;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DOCX] {ex.Message}");
                return null;
            }
        });
    }

    private static void AddParagraph(Body body, string text, bool bold = false, int? fontSizeHalfPoints = null)
    {
        var runProperties = new RunProperties();
        if (bold)
            runProperties.Append(new Bold());
        if (fontSizeHalfPoints.HasValue)
            runProperties.Append(new FontSize { Val = fontSizeHalfPoints.Value.ToString() });

        var run = new Run();
        if (runProperties.ChildElements.Count > 0)
            run.Append(runProperties);

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        body.Append(new Paragraph(run));
    }

    private static void AddImageParagraph(MainDocumentPart mainPart, Body body, string imagePath)
    {
        var imagePartType = imagePath.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
            ? ImagePartType.Png
            : ImagePartType.Jpeg;

        var imagePart = mainPart.AddImagePart(imagePartType);
        using (var stream = File.OpenRead(imagePath))
            imagePart.FeedData(stream);

        var relationshipId = mainPart.GetIdOfPart(imagePart);

        using var image = System.Drawing.Image.FromFile(imagePath);
        double horizontalResolution = image.HorizontalResolution > 0 ? image.HorizontalResolution : 96d;
        double verticalResolution = image.VerticalResolution > 0 ? image.VerticalResolution : 96d;

        const long emusPerInch = 914400;
        long widthEmus = (long)(image.Width / horizontalResolution * emusPerInch);
        long heightEmus = (long)(image.Height / verticalResolution * emusPerInch);

        const long maxWidthEmus = (long)(6.5 * emusPerInch);
        if (widthEmus > maxWidthEmus)
        {
            double scale = (double)maxWidthEmus / widthEmus;
            widthEmus = maxWidthEmus;
            heightEmus = (long)(heightEmus * scale);
        }

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmus, Cy = heightEmus },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties { Id = 1U, Name = "Step screenshot" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = Path.GetFileName(imagePath) },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

        body.Append(new Paragraph(new Run(drawing)));
    }

    // ── HTML builder (shared by MHT + optional standalone) ────────────────

    private static string BuildHtml(RecordingSession session,
        AppSettings cfg, bool embedCid)
    {
        var sb = new StringBuilder();
        sb.Append("""
            <!DOCTYPE html>
            <html lang="cs">
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <style>
            *{box-sizing:border-box}
            body{font-family:"Segoe UI",Arial,sans-serif;background:#f3f3f3;
                 max-width:960px;margin:0 auto;padding:24px;color:#222}
            h1{color:#0078d4;margin:0 0 4px}
            .meta{font-size:12px;color:#666;margin-bottom:20px}
            .step{background:#fff;border-radius:6px;margin:14px 0;
                  overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.12)}
            .step-head{background:#0078d4;color:#fff;padding:10px 14px;
                       display:flex;align-items:center;gap:10px}
            .badge{background:#fff;color:#0078d4;border-radius:50%;
                   width:28px;height:28px;display:flex;align-items:center;
                   justify-content:center;font-weight:700;font-size:13px;flex-shrink:0}
            .step-head h3{margin:0;font-size:13px;flex:1}
            .step-head .ts{font-size:10px;opacity:.8}
            .step-img{padding:10px}
            .step-img img{max-width:100%;border:1px solid #ddd;border-radius:4px;display:block}
            .step-foot{background:#f7f7f7;border-top:1px solid #eee;
                       padding:8px 14px;font-size:11px;color:#555}
            .step-foot .desc{background:#eff6fc;border-left:3px solid #0078d4;
                             padding:6px 10px;margin-top:6px;border-radius:0 4px 4px 0;
                             font-size:12px;color:#222}
            code{background:#f4f4f4;padding:1px 5px;border-radius:3px;font-size:11px}
            hr{border:none;border-top:1px solid #ddd;margin:8px 0}
            </style>
            """);

        sb.Append($"<title>{WebUtility.HtmlEncode(session.Name)}</title></head><body>");
        sb.Append($"<h1>{WebUtility.HtmlEncode(session.Name)}</h1>");
        sb.Append($"<p class='meta'>Datum: {session.StartTime:dd.MM.yyyy HH:mm} &nbsp;|&nbsp; " +
                  $"Kroků: {session.StepCount} &nbsp;|&nbsp; " +
                  $"Trvání: {session.Duration:mm\\:ss}</p>");

        foreach (var step in session.Steps)
        {
            sb.Append("<div class='step'>");
            sb.Append("<div class='step-head'>");
            sb.Append($"<div class='badge'>{step.StepNumber}</div>");
            sb.Append($"<h3>{WebUtility.HtmlEncode(step.DisplayName)}</h3>");
            if (cfg.IncludeTimestamps)
                sb.Append($"<span class='ts'>{step.TimeDisplay} {WebUtility.HtmlEncode(step.SourceIcon)}</span>");
            sb.Append("</div>");

            if (step.ImagePath != null && File.Exists(step.ImagePath))
            {
                var imgSrc = embedCid
                    ? $"cid:img{step.StepNumber}@sr"
                    : step.ImagePath.Replace('\\', '/');
                sb.Append($"<div class='step-img'><img src='{imgSrc}' alt='Krok {step.StepNumber}'></div>");
            }

            sb.Append("<div class='step-foot'>");
            if (cfg.IncludeWindowTitles && !string.IsNullOrEmpty(step.WindowTitle))
                sb.Append($"🪟 {WebUtility.HtmlEncode(step.WindowTitle)} " +
                          $"[{WebUtility.HtmlEncode(step.ProcessName)}]<hr>");
            if (!string.IsNullOrEmpty(step.KeysPressed))
                sb.Append($"⌨ <code>{WebUtility.HtmlEncode(step.KeysPressed)}</code><hr>");
            if (!string.IsNullOrEmpty(step.Description))
                sb.Append($"<div class='desc'>{WebUtility.HtmlEncode(step.Description)}</div>");
            sb.Append("</div></div>");
        }

        sb.Append("</body></html>");
        return sb.ToString();
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static int CountFormats(ExportFormatFlags f)
    {
        int n = 0;
        if (f.HasFlag(ExportFormatFlags.PDF))  n++;
        if (f.HasFlag(ExportFormatFlags.MHT))  n++;
        if (f.HasFlag(ExportFormatFlags.Word)) n++;
        return n;
    }

    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name.Trim();
    }

    private static string GetUniqueDirectoryPath(string basePath)
    {
        if (!Directory.Exists(basePath))
            return basePath;

        string parent = Path.GetDirectoryName(basePath) ?? string.Empty;
        string name = Path.GetFileName(basePath);

        for (int i = 2; ; i++)
        {
            string candidate = Path.Combine(parent, $"{name} ({i})");
            if (!Directory.Exists(candidate))
                return candidate;
        }
    }
}
