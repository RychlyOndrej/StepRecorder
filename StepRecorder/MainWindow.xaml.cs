using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using StepRecorder.Models;
using StepRecorder.Services;
using StepRecorder.Helpers;

namespace StepRecorder;

public partial class MainWindow : Window
{
    // ── Services ──────────────────────────────────────────────────────────
    private AppSettings          _cfg    = SettingsManager.Load();
    private RecordingService?    _recorder;
    private readonly ExportService  _exporter  = new();
    private readonly TrayService    _tray      = new();
    private readonly WindowDetectionService _winSvc = new();

    // ── Session state ─────────────────────────────────────────────────────
    private Models.RecordingSession? _session;
    private readonly ObservableCollection<Models.RecordingStep> _steps = new();

    // ── Hotkey IDs ────────────────────────────────────────────────────────
    private const int HK_FULL = 1001;
    private const int HK_CROP = 1002;
    private const int HK_STOP = 1003;

    // ── Currently viewed step ─────────────────────────────────────────────
    private Models.RecordingStep? _currentStep;

    // ── Suppress save during load ─────────────────────────────────────────
    private bool _loading = true;

    // ─────────────────────────────────────────────────────────────────────
    public MainWindow()
    {
        InitializeComponent();
        StepListBox.ItemsSource = _steps;

        _tray.Initialize();
        _tray.ShowRequested           += () => Dispatcher.Invoke(ShowWindow);
        _tray.StopAndExportRequested  += () => Dispatcher.Invoke(StopAndExportFromTray);
        _tray.ExitRequested           += () => Dispatcher.Invoke(() => { CleanUp(); System.Windows.Application.Current.Shutdown(); });

        LoadSettingsIntoUI();
        RefreshTargetWindows();
        UpdateExportSummary();

        // Wire auto-save to all setting controls
        Loaded += (_, _) => WireAutoSave();

        _loading = false;
    }

    private void WireAutoSave()
    {
        // CheckBoxes
        foreach (var cb in FindVisualDescendants<System.Windows.Controls.CheckBox>(this))
        {
            cb.Checked   += (_, _) => SaveSettingsFromUI();
            cb.Unchecked += (_, _) => SaveSettingsFromUI();
        }
        // RadioButtons
        foreach (var rb in FindVisualDescendants<System.Windows.Controls.RadioButton>(this))
            rb.Checked += (_, _) => SaveSettingsFromUI();
        // ComboBox
        ImageQualityCombo.SelectionChanged += (_, _) => SaveSettingsFromUI();
        // TextBoxes that matter for settings (not step description, not read-only)
        AuthorBox   .LostFocus += (_, _) => SaveSettingsFromUI();
        CropWidthBox.LostFocus += (_, _) => SaveSettingsFromUI();
        CropHeightBox.LostFocus += (_, _) => SaveSettingsFromUI();
        CursorRadiusBox.LostFocus += (_, _) => SaveSettingsFromUI();
    }

    private static IEnumerable<T> FindVisualDescendants<T>(System.Windows.DependencyObject parent)
        where T : System.Windows.DependencyObject
    {
        int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T t) yield return t;
            foreach (var d in FindVisualDescendants<T>(child)) yield return d;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  SETTINGS  ←→  UI
    // ══════════════════════════════════════════════════════════════════════

    private void LoadSettingsIntoUI()
    {
        _loading = true;

        // Relace tab
        SessionNameBox.Text   = $"Nahrávka {DateTime.Now:yyyy-MM-dd HH-mm}";
        OutputFolderBox.Text  = _cfg.OutputFolder;
        AuthorBox.Text        = _cfg.DocumentAuthor;

        ExportPdfCheck.IsChecked  = _cfg.ExportFormats.HasFlag(ExportFormatFlags.PDF);
        ExportMhtCheck.IsChecked  = _cfg.ExportFormats.HasFlag(ExportFormatFlags.MHT);
        ExportWordCheck.IsChecked = _cfg.ExportFormats.HasFlag(ExportFormatFlags.Word);

        LayoutNumbered      .IsChecked = _cfg.Layout == DocumentLayout.NumberedSteps;
        LayoutDetailed      .IsChecked = _cfg.Layout == DocumentLayout.DetailedSteps;
        LayoutScreenshotsOnly.IsChecked = _cfg.Layout == DocumentLayout.ScreenshotsOnly;

        IncludeTimestampsCheck  .IsChecked = _cfg.IncludeTimestamps;
        IncludeWindowTitlesCheck.IsChecked = _cfg.IncludeWindowTitles;

        ImageQualityCombo.SelectedIndex = _cfg.ImageQuality switch
        {
            ImageSaveQuality.JPEG85 => 1,
            ImageSaveQuality.JPEG70 => 2,
            _                       => 0
        };

        // Nahrávání tab
        CaptureWindowOnlyCheck.IsChecked = _cfg.CaptureWindowOnly;
        CaptureOnClickCheck   .IsChecked = _cfg.CaptureOnMouseClick;
        CaptureOnHotkeyCheck  .IsChecked = _cfg.CaptureOnHotkey;

        HotkeyFullText.Text = _cfg.HotkeyFullCapture.DisplayText;
        HotkeyCropText.Text = _cfg.HotkeyCropCapture.DisplayText;

        CropWidthBox .Text = _cfg.CropWidth .ToString();
        CropHeightBox.Text = _cfg.CropHeight.ToString();

        HighlightCursorCheck.IsChecked  = _cfg.HighlightCursor;
        ShowStepNumberBadgeCheck.IsChecked = _cfg.ShowStepNumberBadge;
        RecordKeystrokesCheck.IsChecked = _cfg.RecordKeystrokes;
        RecordWindowNameCheck.IsChecked = _cfg.RecordWindowName;
        CursorRadiusBox.Text            = _cfg.CursorHighlightRadius.ToString();

        try
        {
            ColorPreviewBorder.Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter
                    .ConvertFromString(_cfg.CursorHighlightColor));
        }
        catch { }

        UpdateHotkeyHint();
        _loading = false;
    }

    private void SaveSettingsFromUI()
    {
        if (_loading) return;

        _cfg.DocumentAuthor = AuthorBox.Text;
        _cfg.OutputFolder   = OutputFolderBox.Text;

        _cfg.ExportFormats =
            (ExportPdfCheck .IsChecked == true ? ExportFormatFlags.PDF  : 0) |
            (ExportMhtCheck .IsChecked == true ? ExportFormatFlags.MHT  : 0) |
            (ExportWordCheck.IsChecked == true ? ExportFormatFlags.Word : 0);

        _cfg.Layout = LayoutDetailed      .IsChecked == true ? DocumentLayout.DetailedSteps
                    : LayoutScreenshotsOnly.IsChecked == true ? DocumentLayout.ScreenshotsOnly
                                                               : DocumentLayout.NumberedSteps;

        _cfg.IncludeTimestamps   = IncludeTimestampsCheck  .IsChecked == true;
        _cfg.IncludeWindowTitles = IncludeWindowTitlesCheck.IsChecked == true;

        _cfg.ImageQuality = ImageQualityCombo.SelectedIndex switch
        {
            1 => ImageSaveQuality.JPEG85,
            2 => ImageSaveQuality.JPEG70,
            _ => ImageSaveQuality.PNG
        };

        _cfg.CaptureWindowOnly  = CaptureWindowOnlyCheck.IsChecked == true;
        _cfg.CaptureOnMouseClick= CaptureOnClickCheck   .IsChecked == true;
        _cfg.CaptureOnHotkey    = CaptureOnHotkeyCheck  .IsChecked == true;

        _cfg.CursorHighlightRadius = int.TryParse(CursorRadiusBox.Text, out int r) ? r : 25;
        _cfg.HighlightCursor  = HighlightCursorCheck .IsChecked == true;
        _cfg.ShowStepNumberBadge = ShowStepNumberBadgeCheck.IsChecked == true;
        _cfg.RecordKeystrokes = RecordKeystrokesCheck.IsChecked == true;
        _cfg.RecordWindowName = RecordWindowNameCheck.IsChecked == true;

        if (int.TryParse(CropWidthBox .Text, out int cw)) _cfg.CropWidth  = cw;
        if (int.TryParse(CropHeightBox.Text, out int ch)) _cfg.CropHeight = ch;

        SettingsManager.Save(_cfg);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  RECORDING  CONTROL
    // ══════════════════════════════════════════════════════════════════════

    private void StartRecording_Click(object sender, RoutedEventArgs e)
    {
        SaveSettingsFromUI();

        if ((_cfg.ExportFormats & (ExportFormatFlags.PDF | ExportFormatFlags.MHT | ExportFormatFlags.Word)) == 0)
        {
            System.Windows.MessageBox.Show("Prosím vyberte alespoň jeden výstupní formát na záložce Relace.",
                            "Žádný formát", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _steps.Clear();
        _currentStep = null;
        StepCountBadge.Text = "0";
        RecBadgeText.Text   = "NAHRÁVÁM  •  0 kroků";
        StepDetailContent.Visibility = Visibility.Collapsed;
        EmptyState       .Visibility = Visibility.Visible;
        ExportResultPanel.Visibility = Visibility.Collapsed;

        string sessionName = string.IsNullOrWhiteSpace(SessionNameBox.Text)
            ? $"Nahrávka {DateTime.Now:yyyy-MM-dd HH-mm}"
            : SessionNameBox.Text.Trim();

        string sessionFolder = GetUniqueDirectoryPath(Path.Combine(_cfg.OutputFolder, SanitizeFileName(sessionName)));

        // Create session
        _session = new Models.RecordingSession
        {
            Name          = sessionName,
            SessionFolder = sessionFolder,
        };
        _session.ImagesFolder = Path.Combine(_session.SessionFolder, "images");

        // Determine target window
        IntPtr targetHwnd = IntPtr.Zero;
        if (TargetWindowCombo.SelectedItem is WindowInfo wi)
            targetHwnd = wi.Handle;

        // Start recorder
        _recorder = new RecordingService(_cfg);
        _recorder.StepAdded += OnStepAdded;
        _recorder.StartRecording(_session, targetHwnd);

        // Register hotkeys
        if (_cfg.CaptureOnHotkey)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            NativeMethods.RegisterHotKey(hwnd, HK_FULL,
                _cfg.HotkeyFullCapture.Modifiers  | NativeMethods.MOD_NOREPEAT,
                _cfg.HotkeyFullCapture.VirtualKey);
            NativeMethods.RegisterHotKey(hwnd, HK_CROP,
                _cfg.HotkeyCropCapture.Modifiers  | NativeMethods.MOD_NOREPEAT,
                _cfg.HotkeyCropCapture.VirtualKey);
        }
        NativeMethods.RegisterHotKey(
            new WindowInteropHelper(this).Handle, HK_STOP,
            _cfg.HotkeyStopRecording.Modifiers | NativeMethods.MOD_NOREPEAT,
            _cfg.HotkeyStopRecording.VirtualKey);

        SetRecordingUiState(true);
        StatusText.Text = "Nahrávám…  Klikejte v cílové aplikaci.";

        _tray.Show();
        _tray.SetRecordingState(true, 0);
    }

    private void StopRecording_Click(object sender, RoutedEventArgs e) => StopRecording();

    private void StopRecording()
    {
        if (_recorder == null || _session == null) return;

        _recorder.StopRecording();
        _recorder.Dispose();
        _recorder = null;

        UnregisterHotkeys();

        SetRecordingUiState(false);
        _tray.SetRecordingState(false);

        StatusText.Text = $"Nahrávání zastaveno. {_session.StepCount} kroků.";
        UpdateExportSummary();

        // Auto-switch to export tab
        MainTabs.SelectedIndex = 3;
    }

    private void StopAndExportFromTray()
    {
        StopRecording();
        ShowWindow();
        Export_Click(this, new RoutedEventArgs());
    }

    // ── UI state helpers ──────────────────────────────────────────────────

    private void SetRecordingUiState(bool recording)
    {
        StartBtn          .IsEnabled  = !recording;
        StopBtn           .IsEnabled  = recording;
        MinimizeToTrayBtn .Visibility = recording ? Visibility.Visible : Visibility.Collapsed;
        RecordingBadge    .Visibility = recording ? Visibility.Visible : Visibility.Collapsed;
        ExportBtn         .IsEnabled  = !recording && (_session?.StepCount ?? 0) > 0;
        UpdateHotkeyHint();
    }

    private void UpdateHotkeyHint()
    {
        if (StartBtn.IsEnabled)
        {
            HotkeyHintText.Text = string.Empty;
        }
        else
        {
            HotkeyHintText.Text =
                $"Celé: {_cfg.HotkeyFullCapture.DisplayText}  " +
                $"Výřez: {_cfg.HotkeyCropCapture.DisplayText}  " +
                $"Stop: {_cfg.HotkeyStopRecording.DisplayText}";
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  STEP EVENTS
    // ══════════════════════════════════════════════════════════════════════

    private void OnStepAdded(Models.RecordingStep step)
    {
        Dispatcher.InvokeAsync(() =>
        {
            _steps.Add(step);
            StepCountBadge.Text    = _steps.Count.ToString();
            RecBadgeText.Text      = $"NAHRÁVÁM  •  {_steps.Count} kroků";
            _tray.SetRecordingState(true, _steps.Count);

            // Auto-select
            StepListBox.SelectedItem = step;
            StepListBox.ScrollIntoView(step);
        });
    }

    // ══════════════════════════════════════════════════════════════════════
    //  STEP LIST  &  DETAIL
    // ══════════════════════════════════════════════════════════════════════

    private void StepListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _currentStep = StepListBox.SelectedItem as Models.RecordingStep;
        ShowStepDetail(_currentStep);
    }

    private void ShowStepDetail(Models.RecordingStep? step)
    {
        if (step == null)
        {
            EmptyState      .Visibility = Visibility.Visible;
            StepDetailContent.Visibility = Visibility.Collapsed;
            OpenImageBtn.IsEnabled       = false;
            OpenFolderBtn.IsEnabled      = false;
            return;
        }

        EmptyState      .Visibility = Visibility.Collapsed;
        StepDetailContent.Visibility = Visibility.Visible;
        OpenImageBtn.IsEnabled       = step.ImagePath != null && File.Exists(step.ImagePath);
        OpenFolderBtn.IsEnabled      = step.ImagePath != null;

        // Screenshot
        StepScreenshot.Source = null;
        if (step.ImagePath != null && File.Exists(step.ImagePath))
        {
            try
            {
                var bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.UriSource     = new Uri(step.ImagePath);
                bmi.CacheOption   = BitmapCacheOption.OnLoad;
                bmi.DecodePixelWidth = 860;
                bmi.EndInit();
                bmi.Freeze();
                StepScreenshot.Source = bmi;
            }
            catch { }
        }

        // Metadata
        MetaTime  .Text = step.TimeDisplay;
        MetaWindow.Text = _cfg.IncludeWindowTitles && step.WindowTitle != null
            ? $"{step.WindowTitle}  [{step.ProcessName}]"
            : "—";
        MetaKeys  .Text = step.KeysPressed ?? "—";
        MetaSource.Text = step.Source switch
        {
            CaptureSource.MouseClick  => "Kliknutí myší",
            CaptureSource.HotkeyFull  => "Klávesová zkratka – celé okno",
            CaptureSource.HotkeyCrop  => "Klávesová zkratka – výřez",
            _                         => "?"
        };

        // Description (suppress TextChanged re-entry)
        _loading = true;
        StepDescriptionBox.Text = step.Description ?? string.Empty;
        _loading = false;
    }

    private void StepDescription_Changed(object sender, TextChangedEventArgs e)
    {
        if (_loading || _currentStep == null) return;
        _currentStep.Description = StepDescriptionBox.Text;
    }

    private void DeleteStep_Click(object sender, RoutedEventArgs e)
    {
        if (StepListBox.SelectedItem is not Models.RecordingStep step) return;
        if (System.Windows.MessageBox.Show($"Smazat krok {step.StepNumber}?",
                            "Potvrzení", MessageBoxButton.YesNo,
                            MessageBoxImage.Question) != MessageBoxResult.Yes) return;

        _steps.Remove(step);
        if (step.ImagePath != null) { try { File.Delete(step.ImagePath); } catch { } }

        // Renumber
        for (int i = 0; i < _steps.Count; i++)
            _steps[i].StepNumber = i + 1;

        StepCountBadge.Text = _steps.Count.ToString();
    }

    private void OpenImage_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStep?.ImagePath != null && File.Exists(_currentStep.ImagePath))
            Process.Start(new ProcessStartInfo(_currentStep.ImagePath) { UseShellExecute = true });
    }

    private void OpenFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStep?.ImagePath != null)
            Process.Start(new ProcessStartInfo("explorer.exe",
                $"/select,\"{_currentStep.ImagePath}\"") { UseShellExecute = true });
    }

    // ══════════════════════════════════════════════════════════════════════
    //  EXPORT
    // ══════════════════════════════════════════════════════════════════════

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_session == null || _session.StepCount == 0)
        {
            System.Windows.MessageBox.Show("Nejsou žádné kroky k exportu.",
                            "Prázdná relace", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SaveSettingsFromUI();
        _session.Name = SessionNameBox.Text.Trim();

        ExportBtn           .IsEnabled  = false;
        ExportProgressPanel .Visibility = Visibility.Visible;
        ExportResultPanel   .Visibility = Visibility.Collapsed;
        StatusText.Text = "Exportuji…";

        var progress = new Progress<(int current, int total, string label)>(t =>
        {
            ExportProgressBar.IsIndeterminate = false;
            ExportProgressBar.Maximum         = t.total;
            ExportProgressBar.Value           = t.current;
            ExportProgressLabel.Text          = t.label;
        });

        var files = await _exporter.ExportAsync(_session, _cfg, progress);

        ExportProgressPanel.Visibility = Visibility.Collapsed;

        if (files.Count > 0)
        {
            ExportedFilesList.ItemsSource = files;
            ExportResultPanel.Visibility  = Visibility.Visible;
            StatusText.Text = $"Export dokončen — {files.Count} soubor(y).";
            _tray.ShowBalloon("Export dokončen",
                $"Uloženo {files.Count} soubor(ů) do {_cfg.OutputFolder}");
        }
        else
        {
            StatusText.Text = "Export selhal. Zkontrolujte nastavení.";
            System.Windows.MessageBox.Show("Export nebyl úspěšný. Zkontrolujte, zda jsou nainstalované balíčky.",
                            "Chyba exportu", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        ExportBtn.IsEnabled = true;
    }

    private void ExportedFile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Documents.Hyperlink hl && hl.Tag is string path && File.Exists(path))
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = _session?.SessionFolder ?? _cfg.OutputFolder;
        if (Directory.Exists(folder))
            Process.Start(new ProcessStartInfo("explorer.exe", folder) { UseShellExecute = true });
        else if (Directory.Exists(_cfg.OutputFolder))
            Process.Start(new ProcessStartInfo("explorer.exe", _cfg.OutputFolder) { UseShellExecute = true });
    }

    private void UpdateExportSummary()
    {
        ExportSessionName .Text = _session?.Name       ?? "—";
        ExportStepCount   .Text = _session != null     ? $"{_session.StepCount} kroků" : "0";
        ExportOutputFolder.Text = _cfg.OutputFolder;
        ExportBtn.IsEnabled     = _session?.StepCount > 0 && _recorder == null;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  HOTKEYS  (WndProc)
    // ══════════════════════════════════════════════════════════════════════

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        HwndSource.FromHwnd(new WindowInteropHelper(this).Handle)
                  ?.AddHook(WndProc);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam,
                            ref bool handled)
    {
        if (msg == NativeMethods.WM_HOTKEY)
        {
            int id = wParam.ToInt32();
            switch (id)
            {
                case HK_FULL:
                    _recorder?.TriggerFullCapture();
                    handled = true;
                    break;
                case HK_CROP:
                    _recorder?.TriggerCropCapture();
                    handled = true;
                    break;
                case HK_STOP:
                    if (_recorder != null) Dispatcher.InvokeAsync(StopRecording);
                    handled = true;
                    break;
            }
        }
        return IntPtr.Zero;
    }

    private void UnregisterHotkeys()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        NativeMethods.UnregisterHotKey(hwnd, HK_FULL);
        NativeMethods.UnregisterHotKey(hwnd, HK_CROP);
        NativeMethods.UnregisterHotKey(hwnd, HK_STOP);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  HOTKEY CHANGE DIALOGS
    // ══════════════════════════════════════════════════════════════════════

    private void ChangeHotkeyFull_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new HotkeyPickerDialog(_cfg.HotkeyFullCapture) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _cfg.HotkeyFullCapture = dlg.Result!;
            HotkeyFullText.Text    = dlg.Result!.DisplayText;
            SettingsManager.Save(_cfg);
        }
    }

    private void ChangeHotkeyCrop_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new HotkeyPickerDialog(_cfg.HotkeyCropCapture) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _cfg.HotkeyCropCapture = dlg.Result!;
            HotkeyCropText.Text    = dlg.Result!.DisplayText;
            SettingsManager.Save(_cfg);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  COLOR PICKER
    // ══════════════════════════════════════════════════════════════════════

    private void ColorPreview_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dlg = new System.Windows.Forms.ColorDialog
        {
            FullOpen   = true,
            AnyColor   = true,
        };

        try
        {
            var c = (System.Windows.Media.Color)System.Windows.Media.ColorConverter
                       .ConvertFromString(_cfg.CursorHighlightColor);
            dlg.Color = System.Drawing.Color.FromArgb(c.A, c.R, c.G, c.B);
        }
        catch { }

        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var wc = System.Windows.Media.Color.FromArgb(
                         dlg.Color.A, dlg.Color.R, dlg.Color.G, dlg.Color.B);
            _cfg.CursorHighlightColor       = $"#{wc.A:X2}{wc.R:X2}{wc.G:X2}{wc.B:X2}";
            ColorPreviewBorder.Background   = new System.Windows.Media.SolidColorBrush(wc);
            SettingsManager.Save(_cfg);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  WINDOW / FOLDER  HELPERS
    // ══════════════════════════════════════════════════════════════════════

    private void BrowseFolder_Click(object sender, RoutedEventArgs e)
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog
        {
            Description          = "Vyberte výstupní složku",
            UseDescriptionForTitle = true,
            SelectedPath         = _cfg.OutputFolder
        };
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            OutputFolderBox.Text = dlg.SelectedPath;
            _cfg.OutputFolder    = dlg.SelectedPath;
            SettingsManager.Save(_cfg);
        }
    }

    private void RefreshTargetWindows()
    {
        var wins = _winSvc.GetVisibleWindows();
        TargetWindowCombo.ItemsSource   = wins;
        TargetWindowCombo.DisplayMemberPath = null;

        // Custom item template already set in XAML via DisplayMemberPath="."
        // so we set it via ToString() on WindowInfo
    }

    private void TargetWindowCombo_DropDownOpened(object sender, EventArgs e) =>
        RefreshTargetWindows();

    private void RefreshWindows_Click(object sender, RoutedEventArgs e) =>
        RefreshTargetWindows();

    // ══════════════════════════════════════════════════════════════════════
    //  TRAY / WINDOW  LIFECYCLE
    // ══════════════════════════════════════════════════════════════════════

    private void MinimizeToTray_Click(object sender, RoutedEventArgs e)
    {
        Hide();
        _tray.Show();
        _tray.ShowBalloon("StepRecorder",
            "Nahrávání pokračuje na pozadí.\nDvakrát klikněte pro zobrazení.");
    }

    private void ShowWindow()
    {
        Show();
        WindowState = WindowState.Normal;
        Activate();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_recorder != null)
        {
            var result = System.Windows.MessageBox.Show(
                "Nahrávání stále probíhá. Opravdu chcete ukončit?\n\nData budou ztracena.",
                "Ukončit?",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
                return;
            }
        }

        CleanUp();
        System.Windows.Application.Current.Shutdown();
    }

    private void CleanUp()
    {
        UnregisterHotkeys();
        _recorder?.StopRecording();
        _recorder?.Dispose();
        _tray.Hide();
        _tray.Dispose();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  MISC
    // ══════════════════════════════════════════════════════════════════════

    private static string SanitizeFileName(string s)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
        return s.Trim();
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
