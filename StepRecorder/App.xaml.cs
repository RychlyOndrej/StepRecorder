using System.Windows;

namespace StepRecorder;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(System.Windows.StartupEventArgs e)
    {
        // DPI awareness — must be set before any window is created
        try { NativeMethods.SetProcessDpiAwareness(2); } catch { }

        // QuestPDF community license
        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;

        base.OnStartup(e);
    }
}
