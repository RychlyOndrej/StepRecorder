using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using StepRecorder.Models;

namespace StepRecorder.Services;

/// <summary>Captures screen regions and saves them immediately to disk.</summary>
public sealed class ScreenCaptureService
{
    // ── Capture modes ─────────────────────────────────────────────────────

    /// <summary>Capture the window identified by <paramref name="hwnd"/>.</summary>
    public (Bitmap bitmap, Point clickInImage) CaptureWindow(
        IntPtr hwnd, int screenClickX, int screenClickY)
    {
        Rectangle bounds = GetWindowBounds(hwnd);

        bounds = ConstrainToPrimaryBounds(bounds);
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return CaptureFullScreen(screenClickX, screenClickY);

        var bmp = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);

        var clickInImage = new Point(
            screenClickX - bounds.X,
            screenClickY - bounds.Y);

        return (bmp, clickInImage);
    }

    /// <summary>
    /// Capture a fixed-size crop around the cursor constrained to the given window.
    /// If the window is smaller than crop size, remaining area is left empty.
    /// </summary>
    public (Bitmap bitmap, Point clickInImage) CaptureCropInWindow(
        IntPtr hwnd, int screenClickX, int screenClickY, int width, int height)
    {
        if (hwnd == IntPtr.Zero)
            return CaptureCrop(screenClickX, screenClickY, width, height);

        Rectangle windowBounds = ConstrainToPrimaryBounds(GetWindowBounds(hwnd));
        if (windowBounds.Width <= 0 || windowBounds.Height <= 0)
            return CaptureCrop(screenClickX, screenClickY, width, height);

        int x = screenClickX - width / 2;
        int y = screenClickY - height / 2;

        int minX = windowBounds.Left;
        int maxX = windowBounds.Right - width;
        int minY = windowBounds.Top;
        int maxY = windowBounds.Bottom - height;

        x = maxX >= minX ? Math.Clamp(x, minX, maxX) : minX;
        y = maxY >= minY ? Math.Clamp(y, minY, maxY) : minY;

        var requestedRect = new Rectangle(x, y, width, height);
        var sourceRect = Rectangle.Intersect(requestedRect, windowBounds);

        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Black);

        if (sourceRect.Width > 0 && sourceRect.Height > 0)
        {
            var destinationPoint = new Point(
                sourceRect.X - requestedRect.X,
                sourceRect.Y - requestedRect.Y);

            g.CopyFromScreen(sourceRect.Location, destinationPoint, sourceRect.Size);
        }

        var clickInImage = new Point(
            Math.Clamp(screenClickX - requestedRect.X, 0, width - 1),
            Math.Clamp(screenClickY - requestedRect.Y, 0, height - 1));

        return (bmp, clickInImage);
    }

    /// <summary>Capture the entire screen that contains the click point.</summary>
    public (Bitmap bitmap, Point clickInImage) CaptureFullScreen(
        int screenClickX, int screenClickY)
    {
        var screen = System.Windows.Forms.Screen.FromPoint(
            new Point(screenClickX, screenClickY));
        var bounds = screen.Bounds;

        var bmp = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);

        var clickInImage = new Point(
            screenClickX - bounds.X,
            screenClickY - bounds.Y);

        return (bmp, clickInImage);
    }

    /// <summary>Capture a crop centred on the mouse cursor.</summary>
    public (Bitmap bitmap, Point clickInImage) CaptureCrop(
        int screenClickX, int screenClickY, int width, int height)
    {
        int x = screenClickX - width  / 2;
        int y = screenClickY - height / 2;

        var total = AllScreensBounds();
        x = Math.Clamp(x, total.Left, Math.Max(total.Left, total.Right  - width));
        y = Math.Clamp(y, total.Top,  Math.Max(total.Top,  total.Bottom - height));

        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(x, y, 0, 0, new Size(width, height));

        var clickInImage = new Point(
            screenClickX - x,
            screenClickY - y);

        return (bmp, clickInImage);
    }

    // ── Persistence ───────────────────────────────────────────────────────

    /// <summary>
    /// Saves <paramref name="bitmap"/> to disk immediately and returns the path.
    /// The caller MUST dispose the bitmap afterwards to free RAM.
    /// </summary>
    public string SaveBitmap(Bitmap bitmap, string folder, string baseFileName,
        ImageSaveQuality quality)
    {
        Directory.CreateDirectory(folder);

        string path;
        switch (quality)
        {
            case ImageSaveQuality.JPEG85:
            case ImageSaveQuality.JPEG70:
            {
                path = Path.Combine(folder, baseFileName + ".jpg");
                long q = quality == ImageSaveQuality.JPEG85 ? 85L : 70L;
                var codec = ImageCodecInfo.GetImageEncoders()
                    .First(e => e.MimeType == "image/jpeg");
                using var ep = new EncoderParameters(1);
                ep.Param[0] = new EncoderParameter(Encoder.Quality, q);
                bitmap.Save(path, codec, ep);
                break;
            }
            default:
                path = Path.Combine(folder, baseFileName + ".png");
                bitmap.Save(path, ImageFormat.Png);
                break;
        }

        return path;
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static Rectangle ConstrainToPrimaryBounds(Rectangle r)
    {
        var all = AllScreensBounds();
        return Rectangle.Intersect(r, all);
    }

    private static Rectangle GetWindowBounds(IntPtr hwnd)
    {
        int hr = NativeMethods.DwmGetWindowAttribute(
            hwnd,
            NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
            out NativeMethods.RECT dwmRect,
            Marshal.SizeOf<NativeMethods.RECT>());

        if (hr == 0)
            return dwmRect.ToRectangle();

        NativeMethods.GetWindowRect(hwnd, out NativeMethods.RECT wr);
        return wr.ToRectangle();
    }

    private static Rectangle AllScreensBounds()
    {
        var result = Rectangle.Empty;
        foreach (var s in System.Windows.Forms.Screen.AllScreens)
            result = Rectangle.Union(result, s.Bounds);
        return result;
    }
}
