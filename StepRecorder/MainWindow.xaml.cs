using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
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

    private static readonly IReadOnlyDictionary<string, string> EnUiMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Nahrávač postupů pro desktopové aplikace"] = "Step recorder for desktop applications",
            ["📁  Relace"] = "📁  Session",
            ["⏺  Nahrávání"] = "⏺  Recording",
            ["📋  Kroky"] = "📋  Steps",
            ["📤  Export"] = "📤  Export",
            ["Nastavení relace"] = "Session settings",
            ["Základní informace"] = "Basic information",
            ["Název relace:"] = "Session name:",
            ["Výstupní složka:"] = "Output folder:",
            ["📂 Procházet…"] = "📂 Browse…",
            ["Autor:"] = "Author:",
            ["Výstupní formáty"] = "Export formats",
            ["Rozložení dokumentu"] = "Document layout",
            ["Kvalita snímků"] = "Image quality",
            ["Formát uložení:"] = "Save format:",
            ["PDF  – číslované kroky se screenshoty a popisky"] = "PDF - numbered steps with screenshots and notes",
            ["MHT  – webový archiv, otevřít v Edge / IE"] = "MHT - web archive, open in Edge / IE",
            ["Word  – .docx dokument se screenshoty"] = "Word - .docx document with screenshots",
            ["Číslované kroky  (Krok 1, Krok 2 …)"] = "Numbered steps  (Step 1, Step 2 ...)",
            ["Detailní kroky  (screenshot + popis + metadata)"] = "Detailed steps  (screenshot + description + metadata)",
            ["Pouze screenshoty  (bez textového popisu)"] = "Screenshots only  (no text description)",
            ["Zahrnout časové razítko u každého kroku"] = "Include timestamp for each step",
            ["Zahrnout název okna a aplikace"] = "Include window and application name",
            ["PNG  (bezztrátový, větší soubor)"] = "PNG  (lossless, larger file)",
            ["JPEG 85 %  (výchozí, vhodná kvalita)"] = "JPEG 85%  (default, good quality)",
            ["JPEG 70 %  (nejmenší soubory)"] = "JPEG 70%  (smallest files)",
            ["Ovládání nahrávání"] = "Recording control",
            ["Cílová aplikace"] = "Target application",
            ["Zvolené okno:"] = "Selected window:",
            ["Zachycovat pouze oblast vybraného okna (bez taskbaru)"] = "Capture only selected window area (without taskbar)",
            ["Spouštěče zachycení"] = "Capture triggers",
            ["Screenshot při každém kliknutí levým tlačítkem myši"] = "Take screenshot on every left mouse click",
            ["Screenshot na klávesové zkratky:"] = "Take screenshot on hotkeys:",
            ["Celé okno / celá obrazovka:"] = "Full window / full screen:",
            ["Změnit"] = "Change",
            ["Výřez kolem myši:"] = "Crop around mouse:",
            ["Velikost výřezu (crop) kolem myši"] = "Mouse crop size",
            ["Rozměry (px):"] = "Dimensions (px):",
            ["Anotace snímků"] = "Screenshot annotations",
            ["Zvýraznit místo kliknutí"] = "Highlight click position",
            ["Zobrazit číslo kroku na screenshotu"] = "Show step number on screenshot",
            ["Zaznamenat stisknuté klávesy"] = "Record pressed keys",
            ["Zaznamenat název okna"] = "Record window title",
            ["Barva kruhu:"] = "Circle color:",
            ["Poloměr kruhu:"] = "Circle radius:",
            ["▶  Spustit nahrávání"] = "▶  Start recording",
            ["⏹  Zastavit a exportovat"] = "⏹  Stop and export",
            ["⬇  Minimalizovat do tray"] = "⬇  Minimize to tray",
            ["Zaznamenané kroky"] = "Recorded steps",
            ["🗑 Smazat krok"] = "🗑 Delete step",
            ["Vyberte krok ze seznamu"] = "Select a step from the list",
            ["nebo spusťte nahrávání na záložce Nahrávání"] = "or start recording on the Recording tab",
            ["Metadata kroku"] = "Step metadata",
            ["Čas:"] = "Time:",
            ["Aplikace:"] = "Application:",
            ["Klávesy:"] = "Keys:",
            ["Zdroj:"] = "Source:",
            ["Popis kroku"] = "Step description",
            ["🖼 Otevřít obrázek"] = "🖼 Open image",
            ["📂 Otevřít složku"] = "📂 Open folder",
            ["Exportovat nahrávku"] = "Export recording",
            ["Aktuální relace"] = "Current session",
            ["Název:"] = "Name:",
            ["Počet kroků:"] = "Step count:",
            ["✅  Export dokončen!"] = "✅  Export completed!",
            ["📤  Exportovat nyní"] = "📤  Export now",
            ["📂  Otevřít výstupní složku"] = "📂  Open output folder",
            ["Připraven k nahrávání."] = "Ready to record.",
            ["NAHRÁVÁM  •  0 kroků"] = "RECORDING  •  0 steps",
            ["Obnovit seznam"] = "Refresh list"
        };

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
    private bool _changingLanguage;
    private bool _changingTheme;

    private bool IsEnglishUi => string.Equals(_cfg.UiLanguage, "en", StringComparison.OrdinalIgnoreCase);

    // ─────────────────────────────────────────────────────────────────────
    public MainWindow()
    {
        InitializeComponent();
        StepListBox.ItemsSource = _steps;

        _tray.Initialize();
        _tray.ShowRequested           += () => Dispatcher.Invoke(ShowWindow);
        _tray.StopAndExportRequested  += () => Dispatcher.Invoke(StopAndExportFromTray);
        _tray.ExitRequested           += () => Dispatcher.Invoke(() => { CleanUp(); System.Windows.Application.Current.Shutdown(); });

        ApplyLocalizedText();
        LoadSettingsIntoUI();
        ThemeManager.Apply(this, _cfg.UiTheme);
        RefreshTargetWindows();
        UpdateExportSummary();

        // Wire auto-save to all setting controls
        Loaded += (_, _) => WireAutoSave();

        _loading = false;
    }

    private void ApplyLocalizedText()
    {
        LanguageLabel.Content = L10n.T("LanguageLabel");
        ThemeLabel.Content = L10n.T("ThemeLabel");

        if (LanguageCombo.Items.Count >= 2)
        {
            if (LanguageCombo.Items[0] is ComboBoxItem cs)
                cs.Content = L10n.T("LanguageCzech");
            if (LanguageCombo.Items[1] is ComboBoxItem en)
                en.Content = L10n.T("LanguageEnglish");
        }

        if (ThemeCombo.Items.Count >= 2)
        {
            if (ThemeCombo.Items[0] is ComboBoxItem light)
                light.Content = L10n.T("ThemeLight");
            if (ThemeCombo.Items[1] is ComboBoxItem dark)
                dark.Content = L10n.T("ThemeDark");
        }

        if (!IsEnglishUi)
            return;

        TranslateEnglishRecursive(this);
    }

    private static void TranslateEnglishRecursive(DependencyObject root)
    {
        switch (root)
        {
            case TextBlock tb when TryTranslate(tb.Text, out var translatedText):
                tb.Text = translatedText;
                break;
            case ContentControl cc when cc.Content is string content && TryTranslate(content, out var translatedContent):
                cc.Content = translatedContent;
                break;
            case HeaderedContentControl hcc when hcc.Header is string header && TryTranslate(header, out var translatedHeader):
                hcc.Header = translatedHeader;
                break;
            case Run run when TryTranslate(run.Text, out var translatedRun):
                run.Text = translatedRun;
                break;
            case System.Windows.Controls.ToolTip toolTip when toolTip.Content is string tip && TryTranslate(tip, out var translatedTip):
                toolTip.Content = translatedTip;
                break;
        }

        foreach (var child in LogicalTreeHelper.GetChildren(root))
        {
            if (child is DependencyObject dependencyObject)
                TranslateEnglishRecursive(dependencyObject);
        }
    }

    private static bool TryTranslate(string? source, out string translated)
    {
        translated = string.Empty;
        if (string.IsNullOrWhiteSpace(source))
            return false;

        return EnUiMap.TryGetValue(source, out translated!);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  SETTINGS  ←→  UI
    // ══════════════════════════════════════════════════════════════════════

    private void LoadSettingsIntoUI()
    {
        _loading = true;

        // Relace tab
        SessionNameBox.Text   = $"{(IsEnglishUi ? "Recording" : "Nahrávka")} {DateTime.Now:yyyy-MM-dd HH-mm}";
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

        _changingLanguage = true;
        foreach (var item in LanguageCombo.Items)
        {
            if (item is ComboBoxItem comboItem &&
                string.Equals(comboItem.Tag?.ToString(), _cfg.UiLanguage, StringComparison.OrdinalIgnoreCase))
            {
                LanguageCombo.SelectedItem = comboItem;
                break;
            }
        }
        _changingLanguage = false;

        _changingTheme = true;
        foreach (var item in ThemeCombo.Items)
        {
            if (item is ComboBoxItem comboItem &&
                string.Equals(comboItem.Tag?.ToString(), _cfg.UiTheme, StringComparison.OrdinalIgnoreCase))
            {
                ThemeCombo.SelectedItem = comboItem;
                break;
            }
        }
        _changingTheme = false;

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

        if (LanguageCombo.SelectedItem is ComboBoxItem selectedLanguage)
            _cfg.UiLanguage = selectedLanguage.Tag?.ToString() ?? "cs";

        if (ThemeCombo.SelectedItem is ComboBoxItem selectedTheme)
            _cfg.UiTheme = selectedTheme.Tag?.ToString() ?? "light";

        SettingsManager.Save(_cfg);
    }

    private void WireAutoSave()
    {
        foreach (var cb in FindVisualDescendants<System.Windows.Controls.CheckBox>(this))
        {
            cb.Checked   += (_, _) => SaveSettingsFromUI();
            cb.Unchecked += (_, _) => SaveSettingsFromUI();
        }

        foreach (var rb in FindVisualDescendants<System.Windows.Controls.RadioButton>(this))
            rb.Checked += (_, _) => SaveSettingsFromUI();

        ImageQualityCombo.SelectionChanged += (_, _) => SaveSettingsFromUI();
        AuthorBox.LostFocus += (_, _) => SaveSettingsFromUI();
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
    //  RECORDING  CONTROL
    // ══════════════════════════════════════════════════════════════════════

    private void StartRecording_Click(object sender, RoutedEventArgs e)
    {
        SaveSettingsFromUI();

        if ((_cfg.ExportFormats & (ExportFormatFlags.PDF | ExportFormatFlags.MHT | ExportFormatFlags.Word)) == 0)
        {
            System.Windows.MessageBox.Show(
                L10n.T("NoExportFormatMessage"),
                L10n.T("NoExportFormatTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        _steps.Clear();
        _currentStep = null;
        StepCountBadge.Text = "0";
        RecBadgeText.Text   = IsEnglishUi ? "RECORDING  •  0 steps" : "NAHRÁVÁM  •  0 kroků";
        StepDetailContent.Visibility = Visibility.Collapsed;
        EmptyState       .Visibility = Visibility.Visible;
        ExportResultPanel.Visibility = Visibility.Collapsed;

        string sessionName = string.IsNullOrWhiteSpace(SessionNameBox.Text)
            ? $"{(IsEnglishUi ? "Recording" : "Nahrávka")} {DateTime.Now:yyyy-MM-dd HH-mm}"
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
        StatusText.Text = L10n.T("RecordingStatus");

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

        StatusText.Text = string.Format(L10n.T("RecordingStoppedStatusFormat"), _session.StepCount);
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
            HotkeyHintText.Text = IsEnglishUi
                ? $"Full: {_cfg.HotkeyFullCapture.DisplayText}  Crop: {_cfg.HotkeyCropCapture.DisplayText}  Stop: {_cfg.HotkeyStopRecording.DisplayText}"
                : $"Celé: {_cfg.HotkeyFullCapture.DisplayText}  Výřez: {_cfg.HotkeyCropCapture.DisplayText}  Stop: {_cfg.HotkeyStopRecording.DisplayText}";
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
            RecBadgeText.Text      = IsEnglishUi ? $"RECORDING  •  {_steps.Count} steps" : $"NAHRÁVÁM  •  {_steps.Count} kroků";
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
            CaptureSource.MouseClick  => IsEnglishUi ? "Mouse click" : "Kliknutí myší",
            CaptureSource.HotkeyFull  => IsEnglishUi ? "Hotkey - full window" : "Klávesová zkratka – celé okno",
            CaptureSource.HotkeyCrop  => IsEnglishUi ? "Hotkey - crop" : "Klávesová zkratka – výřez",
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
        if (System.Windows.MessageBox.Show(
                string.Format(L10n.T("DeleteStepConfirmMessageFormat"), step.StepNumber),
                L10n.T("DeleteStepConfirmTitle"),
                MessageBoxButton.YesNo,
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
            System.Windows.MessageBox.Show(
                L10n.T("EmptySessionMessage"),
                L10n.T("EmptySessionTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        SaveSettingsFromUI();
        _session.Name = SessionNameBox.Text.Trim();

        ExportBtn           .IsEnabled  = false;
        ExportProgressPanel .Visibility = Visibility.Visible;
        ExportResultPanel   .Visibility = Visibility.Collapsed;
        StatusText.Text = L10n.T("ExportStatus");

        var progress = new Progress<(int current, int total, string label)>(t =>
        {
            ExportProgressBar.IsIndeterminate = false;
            ExportProgressBar.Maximum         = t.total;
            ExportProgressBar.Value           = t.current;
            ExportProgressLabel.Text          = LocalizeProgressLabel(t.label);
        });

        var files = await _exporter.ExportAsync(_session, _cfg, progress);

        ExportProgressPanel.Visibility = Visibility.Collapsed;

        if (files.Count > 0)
        {
            ExportedFilesList.ItemsSource = files;
            ExportResultPanel.Visibility  = Visibility.Visible;
            StatusText.Text = string.Format(L10n.T("ExportDoneStatusFormat"), files.Count);
            _tray.ShowBalloon(
                L10n.T("ExportDoneTrayTitle"),
                string.Format(L10n.T("ExportDoneTrayMessageFormat"), files.Count, _cfg.OutputFolder));
        }
        else
        {
            StatusText.Text = L10n.T("ExportFailedStatus");
            System.Windows.MessageBox.Show(
                L10n.T("ExportErrorMessage"),
                L10n.T("ExportErrorTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        ExportBtn.IsEnabled = true;
    }

    private string LocalizeProgressLabel(string label)
    {
        if (!IsEnglishUi)
            return label;

        return label switch
        {
            "Generuji PDF…" => "Generating PDF...",
            "Generuji MHT…" => "Generating MHT...",
            "Generuji Word…" => "Generating Word...",
            "Hotovo" => "Done",
            _ => label
        };
    }

    private void UpdateExportSummary()
    {
        ExportSessionName .Text = _session?.Name       ?? "—";
        ExportStepCount   .Text = _session != null
            ? (IsEnglishUi ? $"{_session.StepCount} steps" : $"{_session.StepCount} kroků")
            : "0";
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
            Description          = IsEnglishUi ? "Select output folder" : "Vyberte výstupní složku",
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

    private void LanguageCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || _changingLanguage) return;
        if (LanguageCombo.SelectedItem is not ComboBoxItem selected) return;

        var newLanguage = selected.Tag?.ToString() ?? "cs";
        if (string.Equals(_cfg.UiLanguage, newLanguage, StringComparison.OrdinalIgnoreCase))
            return;

        _cfg.UiLanguage = newLanguage;
        SettingsManager.Save(_cfg);

        var result = System.Windows.MessageBox.Show(
            L10n.T("LanguageRestartMessage"),
            L10n.T("LanguageRestartTitle"),
            MessageBoxButton.YesNo,
            MessageBoxImage.Information);

        if (result == MessageBoxResult.Yes)
        {
            var exe = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(exe))
                Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true });

            CleanUp();
            System.Windows.Application.Current.Shutdown();
        }
    }

    private void ThemeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || _changingTheme) return;
        if (ThemeCombo.SelectedItem is not ComboBoxItem selected) return;

        var newTheme = selected.Tag?.ToString() ?? "light";
        if (string.Equals(_cfg.UiTheme, newTheme, StringComparison.OrdinalIgnoreCase))
            return;

        _cfg.UiTheme = newTheme;
        SettingsManager.Save(_cfg);
        ThemeManager.Apply(this, _cfg.UiTheme);
    }

    private void MinimizeToTray_Click(object sender, RoutedEventArgs e)
    {
        Hide();
        _tray.Show();
        _tray.ShowBalloon("StepRecorder", L10n.T("TrayRecordingBackground"));
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
                L10n.T("ExitConfirmMessage"),
                L10n.T("ExitConfirmTitle"),
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

    private void ExportedFile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Hyperlink hl && hl.Tag is string path && File.Exists(path))
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

    private void RefreshTargetWindows()
    {
        var wins = _winSvc.GetVisibleWindows();
        TargetWindowCombo.ItemsSource = wins;
        TargetWindowCombo.DisplayMemberPath = null;
    }

    private void TargetWindowCombo_DropDownOpened(object sender, EventArgs e) =>
        RefreshTargetWindows();

    private void RefreshWindows_Click(object sender, RoutedEventArgs e) =>
        RefreshTargetWindows();
}
