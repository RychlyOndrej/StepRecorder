using System.IO;
using StepRecorder.Models;

namespace StepRecorder.Services;

/// <summary>
/// Orchestrates the full recording pipeline:
/// hook → capture → annotate → save → metadata only in RAM.
/// </summary>
public sealed class RecordingService : IDisposable
{
    private readonly AppSettings           _cfg;
    private readonly GlobalHookService     _hooks;
    private readonly ScreenCaptureService  _capture;
    private readonly WindowDetectionService _windows;
    private readonly ImageAnnotationService _annotate;

    private RecordingSession? _session;
    private IntPtr            _targetHwnd;
    private bool              _running;
    private long              _lastMouseClickTicks;
    private int               _lastMouseClickX;
    private int               _lastMouseClickY;
    private readonly long     _doubleClickWindowTicks;

    private const int DoubleClickPixelTolerance = 4;

    // ── Events ────────────────────────────────────────────────────────────
    public event Action<RecordingStep>? StepAdded;

    // ── Constructor ───────────────────────────────────────────────────────

    public RecordingService(AppSettings cfg)
    {
        _cfg      = cfg;
        _hooks    = new GlobalHookService();
        _capture  = new ScreenCaptureService();
        _windows  = new WindowDetectionService();
        _annotate = new ImageAnnotationService();

        _doubleClickWindowTicks = TimeSpan.FromMilliseconds(NativeMethods.GetDoubleClickTime()).Ticks;

        _hooks.LeftButtonDown += OnLeftButtonDown;
    }

    // ── Control ───────────────────────────────────────────────────────────

    public void StartRecording(RecordingSession session, IntPtr targetHwnd = default)
    {
        _session    = session;
        _targetHwnd = targetHwnd;
        _running    = true;
        _lastMouseClickTicks = 0;

        Directory.CreateDirectory(session.ImagesFolder);

        if (_cfg.CaptureOnMouseClick || _cfg.CaptureOnHotkey)
            _hooks.Start();
    }

    public void StopRecording()
    {
        _hooks.Stop();
        _session?.Finish();
        _running = false;
    }

    // Called by MainWindow via WndProc on Ctrl+Shift+F
    public void TriggerFullCapture()
    {
        if (!_running) return;
        NativeMethods.GetCursorPos(out var pt);
        ScheduleCapture(pt.X, pt.Y, CaptureSource.HotkeyFull);
    }

    // Called by MainWindow via WndProc on Ctrl+Shift+C
    public void TriggerCropCapture()
    {
        if (!_running) return;
        NativeMethods.GetCursorPos(out var pt);
        ScheduleCapture(pt.X, pt.Y, CaptureSource.HotkeyCrop);
    }

    // ── Mouse hook callback (UI thread) ───────────────────────────────────

    private void OnLeftButtonDown(int x, int y)
    {
        if (!_running || !_cfg.CaptureOnMouseClick) return;

        if (ShouldSkipAsDoubleClick(x, y))
            return;

        // If a target window is locked, only fire when it is foreground
        if (_targetHwnd != IntPtr.Zero &&
            NativeMethods.GetForegroundWindow() != _targetHwnd)
            return;

        ScheduleCapture(x, y, CaptureSource.MouseClick);
    }

    private bool ShouldSkipAsDoubleClick(int x, int y)
    {
        long nowTicks = DateTime.UtcNow.Ticks;

        if (_lastMouseClickTicks > 0)
        {
            bool isWithinWindow = nowTicks - _lastMouseClickTicks <= _doubleClickWindowTicks;
            bool isSameArea = Math.Abs(x - _lastMouseClickX) <= DoubleClickPixelTolerance &&
                              Math.Abs(y - _lastMouseClickY) <= DoubleClickPixelTolerance;

            if (isWithinWindow && isSameArea)
            {
                _lastMouseClickTicks = 0;
                return true;
            }
        }

        _lastMouseClickTicks = nowTicks;
        _lastMouseClickX = x;
        _lastMouseClickY = y;

        return false;
    }

    // ── Capture pipeline ──────────────────────────────────────────────────

    /// <summary>
    /// Grabs keys on the calling thread (UI), then offloads the heavy work
    /// (screenshot + save) to a ThreadPool thread so the hook
    /// callback returns in &lt; 300 ms.
    /// </summary>
    private void ScheduleCapture(int screenX, int screenY, CaptureSource source)
    {
        if (_session == null) return;

        // Snapshot volatile state on the UI thread
        string? keys      = _cfg.RecordKeystrokes ? _hooks.FlushKeys() : null;
        IntPtr  fgHwnd    = _targetHwnd != IntPtr.Zero
                                ? _targetHwnd
                                : NativeMethods.GetForegroundWindow();

        Task.Run(() => ExecuteCapture(screenX, screenY, source, keys, fgHwnd));
    }

    private void ExecuteCapture(int screenX, int screenY, CaptureSource source,
                                 string? keys, IntPtr fgHwnd)
    {
        if (_session == null) return;

        try
        {
            // 1. Get window metadata
            var winInfo = _windows.GetWindowInfo(fgHwnd);

            // 2. Capture bitmap
            System.Drawing.Bitmap bmp;
            System.Drawing.Point  clickInImg;

            if (source == CaptureSource.HotkeyFull)
            {
                if (_cfg.CaptureWindowOnly && fgHwnd != IntPtr.Zero)
                    (bmp, clickInImg) = _capture.CaptureWindow(fgHwnd, screenX, screenY);
                else
                    (bmp, clickInImg) = _capture.CaptureFullScreen(screenX, screenY);
            }
            else if (source == CaptureSource.HotkeyCrop)
            {
                if (_cfg.CaptureWindowOnly && fgHwnd != IntPtr.Zero)
                {
                    (bmp, clickInImg) = _capture.CaptureCropInWindow(
                        fgHwnd, screenX, screenY, _cfg.CropWidth, _cfg.CropHeight);
                }
                else
                {
                    (bmp, clickInImg) = _capture.CaptureCrop(
                        screenX, screenY, _cfg.CropWidth, _cfg.CropHeight);
                }
            }
            else // MouseClick
            {
                if (_cfg.CaptureWindowOnly && fgHwnd != IntPtr.Zero)
                    (bmp, clickInImg) = _capture.CaptureWindow(fgHwnd, screenX, screenY);
                else
                    (bmp, clickInImg) = _capture.CaptureFullScreen(screenX, screenY);
            }

            // 3. Annotate (in-place, no extra allocation)
            int nextNum = _session.StepCount + 1;

            if (_cfg.HighlightCursor && clickInImg.X >= 0)
                _annotate.AnnotateClick(bmp, clickInImg.X, clickInImg.Y,
                    _cfg.CursorHighlightColor, _cfg.CursorHighlightRadius);

            if (_cfg.ShowStepNumberBadge)
                _annotate.AddStepBadge(bmp, nextNum);

            // 4. Save to disk immediately → free RAM
            string imagePath = _capture.SaveBitmap(
                bmp,
                _session.ImagesFolder,
                $"step{nextNum:000}",
                _cfg.ImageQuality);

            bmp.Dispose(); // ← freed here, not held in memory

            // 5. Build lightweight step object
            var step = new RecordingStep
            {
                ImagePath      = imagePath,
                WindowTitle    = _cfg.RecordWindowName ? winInfo?.Title       : null,
                ProcessName    = _cfg.RecordWindowName ? winInfo?.ProcessName : null,
                ProcessId      = winInfo?.ProcessId ?? 0,
                ClickXInImage  = clickInImg.X,
                ClickYInImage  = clickInImg.Y,
                ScreenClickX   = screenX,
                ScreenClickY   = screenY,
                Source         = source,
                KeysPressed    = string.IsNullOrWhiteSpace(keys) ? null : keys
            };

            _session.AddStep(step);
            StepAdded?.Invoke(step);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Capture] Error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        _hooks.Dispose();
    }
}
