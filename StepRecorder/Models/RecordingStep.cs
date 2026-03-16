namespace StepRecorder.Models;

public enum CaptureSource { MouseClick, HotkeyFull, HotkeyCrop }

/// <summary>
/// Lightweight metadata for one recorded step.
/// The actual bitmap is already on disk — never stored in RAM.
/// </summary>
public class RecordingStep
{
    public int      StepNumber { get; set; }
    public DateTime Timestamp  { get; set; } = DateTime.Now;

    // ── Image ────────────────────────────────────────────────────────────
    /// <summary>Absolute path to the annotated screenshot saved on disk.</summary>
    public string? ImagePath { get; set; }

    // ── Context ──────────────────────────────────────────────────────────
    public string? WindowTitle  { get; set; }
    public string? ProcessName  { get; set; }
    public uint    ProcessId    { get; set; }

    // ── Click position ───────────────────────────────────────────────────
    /// <summary>Position relative to captured image (for annotation).</summary>
    public int ClickXInImage  { get; set; } = -1;
    public int ClickYInImage  { get; set; } = -1;

    /// <summary>Absolute screen coordinates.</summary>
    public int ScreenClickX   { get; set; }
    public int ScreenClickY   { get; set; }

    // ── Input ────────────────────────────────────────────────────────────
    public CaptureSource Source      { get; set; }
    public string?       KeysPressed { get; set; }

    // ── Description ──────────────────────────────────────────────────────
    public string? Description            { get; set; }
    public bool    DescriptionIsAiGenerated { get; set; }

    // ── Display helpers ───────────────────────────────────────────────────
    public string DisplayName =>
        !string.IsNullOrWhiteSpace(Description)
            ? Description
            : $"Krok {StepNumber}{(WindowTitle != null ? " – " + WindowTitle : "")}";

    public string TimeDisplay => Timestamp.ToString("HH:mm:ss");

    public string SourceIcon => Source switch
    {
        CaptureSource.MouseClick  => "🖱",
        CaptureSource.HotkeyFull  => "⌨ Celé",
        CaptureSource.HotkeyCrop  => "⌨ Výřez",
        _                         => ""
    };
}

// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Container for one recording session. Holds only lightweight metadata; all
/// bitmaps live on disk in <see cref="ImagesFolder"/>.
/// </summary>
public class RecordingSession
{
    public string   Id            { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public string   Name          { get; set; } = $"Nahrávka {DateTime.Now:yyyy-MM-dd HH-mm-ss}";
    public DateTime StartTime     { get; set; } = DateTime.Now;
    public DateTime? EndTime      { get; set; }

    public string SessionFolder   { get; set; } = string.Empty;
    public string ImagesFolder    { get; set; } = string.Empty;

    public List<RecordingStep> Steps { get; } = new();

    public int      StepCount => Steps.Count;
    public TimeSpan Duration  => (EndTime ?? DateTime.Now) - StartTime;

    public void AddStep(RecordingStep step)
    {
        step.StepNumber = Steps.Count + 1;
        Steps.Add(step);
    }

    public void Finish() => EndTime = DateTime.Now;
}
