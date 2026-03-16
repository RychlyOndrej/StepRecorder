using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace StepRecorder.Services;

public sealed class WindowInfo
{
    public IntPtr  Handle      { get; init; }
    public string  Title       { get; init; } = string.Empty;
    public string  ProcessName { get; init; } = string.Empty;
    public uint    ProcessId   { get; init; }
    public System.Drawing.Rectangle Bounds { get; init; }

    public override string ToString() =>
        string.IsNullOrEmpty(Title)
            ? $"[{ProcessName}]"
            : $"{Title}  [{ProcessName}]";
}

public sealed class WindowDetectionService
{
    // ── Foreground window ────────────────────────────────────────────────

    public WindowInfo? GetForegroundWindowInfo()
    {
        var hwnd = NativeMethods.GetForegroundWindow();
        return hwnd == IntPtr.Zero ? null : BuildInfo(hwnd);
    }

    public WindowInfo? GetWindowInfo(IntPtr hwnd) =>
        hwnd == IntPtr.Zero ? null : BuildInfo(hwnd);

    // ── Enumerate visible windows ────────────────────────────────────────

    public List<WindowInfo> GetVisibleWindows()
    {
        var list = new List<WindowInfo>();

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            if (!NativeMethods.IsWindowVisible(hwnd)) return true;
            if (NativeMethods.IsIconic(hwnd))         return true;
            if (NativeMethods.GetWindowTextLength(hwnd) == 0) return true;

            var info = BuildInfo(hwnd);
            if (info != null)
                list.Add(info);

            return true; // continue enum
        }, IntPtr.Zero);

        return list
            .Where(w => !string.IsNullOrWhiteSpace(w.Title))
            .OrderBy(w => w.Title)
            .ToList();
    }

    // ── Internal ──────────────────────────────────────────────────────────

    private static WindowInfo? BuildInfo(IntPtr hwnd)
    {
        try
        {
            var sb = new StringBuilder(512);
            NativeMethods.GetWindowText(hwnd, sb, 512);

            NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);

            string procName = string.Empty;
            try { procName = Process.GetProcessById((int)pid).ProcessName; }
            catch { }

            // Prefer DWM accurate bounds (handles window shadows / DPI)
            System.Drawing.Rectangle bounds;
            int hrDwm = NativeMethods.DwmGetWindowAttribute(
                hwnd,
                NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
                out NativeMethods.RECT dwmRect,
                Marshal.SizeOf<NativeMethods.RECT>());

            if (hrDwm == 0)
            {
                bounds = dwmRect.ToRectangle();
            }
            else
            {
                NativeMethods.GetWindowRect(hwnd, out NativeMethods.RECT rect);
                bounds = rect.ToRectangle();
            }

            return new WindowInfo
            {
                Handle      = hwnd,
                Title       = sb.ToString().Trim(),
                ProcessName = procName,
                ProcessId   = pid,
                Bounds      = bounds
            };
        }
        catch
        {
            return null;
        }
    }
}
