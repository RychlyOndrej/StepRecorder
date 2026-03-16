using System.Drawing;
using System.Windows.Forms;

namespace StepRecorder.Services;

/// <summary>Manages the NotifyIcon in the Windows system tray.</summary>
public sealed class TrayService : IDisposable
{
    private NotifyIcon? _tray;

    public event Action? ShowRequested;
    public event Action? StopAndExportRequested;
    public event Action? ExitRequested;

    // ── Initialise ────────────────────────────────────────────────────────

    public void Initialize()
    {
        _tray = new NotifyIcon
        {
            Text    = "StepRecorder",
            Icon    = BuildIcon(recording: false),
            Visible = false
        };

        var menu = new ContextMenuStrip();
        menu.Font = new Font("Segoe UI", 9.5f);

        // Show
        var show = (ToolStripMenuItem)menu.Items.Add("📋  Zobrazit");
        show.Font = new Font(show.Font, System.Drawing.FontStyle.Bold);
        show.Click += (_, _) => ShowRequested?.Invoke();

        menu.Items.Add(new ToolStripSeparator());

        // Stop & export
        var stop = (ToolStripMenuItem)menu.Items.Add("⏹  Zastavit a exportovat");
        stop.Enabled = false;
        stop.Name    = "stopItem";
        stop.Click  += (_, _) => StopAndExportRequested?.Invoke();

        menu.Items.Add(new ToolStripSeparator());

        // Exit
        var exit = (ToolStripMenuItem)menu.Items.Add("✕  Ukončit");
        exit.Click += (_, _) => ExitRequested?.Invoke();

        _tray.ContextMenuStrip = menu;
        _tray.DoubleClick     += (_, _) => ShowRequested?.Invoke();
    }

    // ── Visibility ────────────────────────────────────────────────────────

    public void Show() { if (_tray != null) _tray.Visible = true; }
    public void Hide() { if (_tray != null) _tray.Visible = false; }

    // ── State ─────────────────────────────────────────────────────────────

    public void SetRecordingState(bool recording, int stepCount = 0)
    {
        if (_tray == null) return;

        _tray.Icon = BuildIcon(recording);
        _tray.Text = recording
            ? $"StepRecorder  •  Nahrávám ({stepCount} kroků)"
            : "StepRecorder";

        if (_tray.ContextMenuStrip?.Items["stopItem"] is ToolStripMenuItem s)
            s.Enabled = recording;
    }

    public void ShowBalloon(string title, string text,
        ToolTipIcon icon = ToolTipIcon.Info)
    {
        _tray?.ShowBalloonTip(3000, title, text, icon);
    }

    // ── Icon builder ──────────────────────────────────────────────────────

    private static Icon BuildIcon(bool recording)
    {
        using var bmp = new Bitmap(16, 16);
        using var g   = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);

        if (recording)
        {
            // Pulsing red record dot
            g.FillEllipse(new SolidBrush(Color.FromArgb(220, 30, 30)), 1, 1, 14, 14);
            g.FillEllipse(Brushes.White, 4, 4, 8, 8);
            g.FillEllipse(new SolidBrush(Color.FromArgb(200, 20, 20)), 5, 5, 6, 6);
        }
        else
        {
            // Blue camera-style icon
            g.FillRoundedRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)),
                new RectangleF(1, 3, 14, 10), 2);
            // Lens
            g.FillEllipse(Brushes.White,      5, 4, 6, 8);
            g.FillEllipse(new SolidBrush(Color.FromArgb(0, 90, 180)), 6, 5, 4, 6);
            // Viewfinder bump
            g.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), 5, 2, 4, 3);
        }

        IntPtr hIcon = bmp.GetHicon();
        return Icon.FromHandle(hIcon);
    }

    public void Dispose() => _tray?.Dispose();
}

// ── Extension for rounded rectangle ─────────────────────────────────────────
file static class GraphicsExt
{
    public static void FillRoundedRectangle(
        this Graphics g, Brush b, RectangleF rect, float radius)
    {
        using var path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddRoundedRectangle(rect, radius);
        g.FillPath(b, path);
    }

    public static void AddRoundedRectangle(
        this System.Drawing.Drawing2D.GraphicsPath path,
        RectangleF rect, float radius)
    {
        float d = radius * 2;
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
    }
}
