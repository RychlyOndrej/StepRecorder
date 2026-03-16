using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace StepRecorder.Services;

/// <summary>Draws click highlights and step badges onto captured bitmaps.</summary>
public sealed class ImageAnnotationService
{
    // ── Click highlight ───────────────────────────────────────────────────

    public void AnnotateClick(Bitmap bitmap, int clickX, int clickY,
        string colorHex, int radius)
    {
        if (clickX < 0 || clickY < 0) return;
        if (clickX >= bitmap.Width || clickY >= bitmap.Height) return;

        using var g = Graphics.FromImage(bitmap);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        Color c = ColorTranslator.FromHtml(colorHex);

        // Semi-transparent fill
        using var fill = new SolidBrush(Color.FromArgb(50, c));
        var outer = new RectangleF(clickX - radius, clickY - radius,
                                   radius * 2, radius * 2);
        g.FillEllipse(fill, outer);

        // Solid outline
        using var pen = new Pen(c, 3f);
        g.DrawEllipse(pen, outer);

        // Inner dot
        using var dot = new SolidBrush(c);
        g.FillEllipse(dot, clickX - 5, clickY - 5, 10, 10);

        // White border around dot for contrast
        using var dotBorder = new Pen(Color.White, 1.5f);
        g.DrawEllipse(dotBorder, clickX - 5, clickY - 5, 10, 10);
    }

    // ── Step number badge ─────────────────────────────────────────────────

    public void AddStepBadge(Bitmap bitmap, int stepNumber)
    {
        using var g = Graphics.FromImage(bitmap);
        g.SmoothingMode   = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

        const int size   = 38;
        const int margin = 10;

        // Shadow
        using var shadow = new SolidBrush(Color.FromArgb(80, 0, 0, 0));
        g.FillEllipse(shadow, margin + 2, margin + 2, size, size);

        // Badge background
        using var bg = new SolidBrush(Color.FromArgb(230, 0, 102, 204));
        g.FillEllipse(bg, margin, margin, size, size);

        // Border
        using var border = new Pen(Color.White, 1.5f);
        g.DrawEllipse(border, margin, margin, size, size);

        // Number text
        string text = stepNumber.ToString();
        using var font = new Font("Segoe UI", stepNumber > 99 ? 9f : 13f, FontStyle.Bold,
                                  GraphicsUnit.Point);
        using var brush = new SolidBrush(Color.White);
        var tf     = g.MeasureString(text, font);
        float tx   = margin + (size - tf.Width)  / 2f;
        float ty   = margin + (size - tf.Height) / 2f;
        g.DrawString(text, font, brush, tx, ty);
    }
}
