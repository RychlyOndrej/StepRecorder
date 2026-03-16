using System.IO;
using System.Windows.Input;
using StepRecorder.Helpers;

namespace StepRecorder.Models;

// ── Enums ─────────────────────────────────────────────────────────────────────

[Flags]
public enum ExportFormatFlags
{
    None = 0,
    PDF  = 1,
    MHT  = 2,
    Word = 4
}

public enum ImageSaveQuality { PNG, JPEG85, JPEG70 }
public enum DocumentLayout   { NumberedSteps, DetailedSteps, ScreenshotsOnly }

// ── HotkeyDefinition ──────────────────────────────────────────────────────────

public class HotkeyDefinition
{
    /// <summary>Combination of MOD_CONTROL / MOD_SHIFT / MOD_ALT / MOD_WIN.</summary>
    public uint Modifiers { get; set; }

    /// <summary>Windows virtual key code.</summary>
    public uint VirtualKey { get; set; }

    public string DisplayText
    {
        get
        {
            var parts = new List<string>();
            if ((Modifiers & NativeMethods.MOD_CONTROL) != 0) parts.Add("Ctrl");
            if ((Modifiers & NativeMethods.MOD_SHIFT)   != 0) parts.Add("Shift");
            if ((Modifiers & NativeMethods.MOD_ALT)     != 0) parts.Add("Alt");
            if ((Modifiers & NativeMethods.MOD_WIN)     != 0) parts.Add("Win");

            var key = KeyInterop.KeyFromVirtualKey((int)VirtualKey);
            parts.Add(((Key)key).ToString());

            return string.Join("+", parts);
        }
    }

    public static HotkeyDefinition CtrlShift(char letter) => new()
    {
        Modifiers  = NativeMethods.MOD_CONTROL | NativeMethods.MOD_SHIFT,
        VirtualKey = (uint)char.ToUpper(letter)
    };
}

// ── AppSettings ───────────────────────────────────────────────────────────────

public class AppSettings
{
    // ---- Session ----
    public string OutputFolder { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "StepRecorder");

    public ExportFormatFlags ExportFormats { get; set; } =
        ExportFormatFlags.PDF | ExportFormatFlags.MHT | ExportFormatFlags.Word;
    public DocumentLayout    Layout        { get; set; } = DocumentLayout.NumberedSteps;
    public bool IncludeTimestamps  { get; set; } = true;
    public bool IncludeWindowTitles { get; set; } = true;
    public string DocumentAuthor   { get; set; } = Environment.UserName;

    // ---- Capture triggers ----
    public bool CaptureOnMouseClick { get; set; } = true;
    public bool CaptureOnHotkey     { get; set; } = true;

    // ---- Window / region ----
    /// <summary>Capture only the foreground (or target) window, excluding taskbar.</summary>
    public bool CaptureWindowOnly { get; set; } = true;

    /// <summary>For crop-mode: width in physical pixels.</summary>
    public int CropWidth  { get; set; } = 640;
    /// <summary>For crop-mode: height in physical pixels.</summary>
    public int CropHeight { get; set; } = 640;

    // ---- Hotkeys ----
    public HotkeyDefinition HotkeyFullCapture   { get; set; } = HotkeyDefinition.CtrlShift('F');
    public HotkeyDefinition HotkeyCropCapture   { get; set; } = HotkeyDefinition.CtrlShift('C');
    public HotkeyDefinition HotkeyStopRecording { get; set; } = HotkeyDefinition.CtrlShift('S');

    // ---- Annotations ----
    public bool   HighlightCursor        { get; set; } = true;
    public bool   ShowStepNumberBadge    { get; set; } = true;
    public int    CursorHighlightRadius  { get; set; } = 25;
    public string CursorHighlightColor   { get; set; } = "#FFFF0000"; // ARGB hex
    public bool   RecordKeystrokes       { get; set; } = true;
    public bool   RecordWindowName       { get; set; } = true;

    // ---- Image quality ----
    public ImageSaveQuality ImageQuality { get; set; } = ImageSaveQuality.PNG;
}
