using System.Windows;

namespace StepRecorder.Helpers;

internal static class ThemeManager
{
    private static readonly System.Windows.Media.SolidColorBrush DarkWindowBrush = NewBrush(0x1E, 0x1E, 0x1E);
    private static readonly System.Windows.Media.SolidColorBrush DarkSurfaceBrush = NewBrush(0x2D, 0x2D, 0x30);
    private static readonly System.Windows.Media.SolidColorBrush DarkSurfaceAltBrush = NewBrush(0x25, 0x25, 0x28);
    private static readonly System.Windows.Media.SolidColorBrush DarkBorderBrush = NewBrush(0x3F, 0x3F, 0x46);
    private static readonly System.Windows.Media.SolidColorBrush DarkTextBrush = NewBrush(0xF0, 0xF0, 0xF0);
    private static readonly System.Windows.Media.SolidColorBrush DarkMutedTextBrush = NewBrush(0xC8, 0xC8, 0xC8);

    private static readonly System.Windows.Media.SolidColorBrush LightWindowBrush = NewBrush(0xF3, 0xF3, 0xF3);
    private static readonly System.Windows.Media.SolidColorBrush LightSurfaceBrush = NewBrush(0xFF, 0xFF, 0xFF);
    private static readonly System.Windows.Media.SolidColorBrush LightSurfaceAltBrush = NewBrush(0xF7, 0xF7, 0xF7);
    private static readonly System.Windows.Media.SolidColorBrush LightBorderBrush = NewBrush(0xE0, 0xE0, 0xE0);
    private static readonly System.Windows.Media.SolidColorBrush LightTextBrush = NewBrush(0x11, 0x11, 0x11);
    private static readonly System.Windows.Media.SolidColorBrush LightMutedTextBrush = NewBrush(0x66, 0x66, 0x66);

    private static readonly System.Windows.Media.SolidColorBrush DarkTabSelectedBrush = NewBrush(0x35, 0x42, 0x50);
    private static readonly System.Windows.Media.SolidColorBrush DarkTabHoverBrush = NewBrush(0x2F, 0x37, 0x40);
    private static readonly System.Windows.Media.SolidColorBrush DarkListSelectedBrush = NewBrush(0x35, 0x42, 0x50);

    private static readonly System.Windows.Media.SolidColorBrush LightTabSelectedBrush = NewBrush(0xEF, 0xF6, 0xFC);
    private static readonly System.Windows.Media.SolidColorBrush LightTabHoverBrush = NewBrush(0xF5, 0xF5, 0xF5);
    private static readonly System.Windows.Media.SolidColorBrush LightListSelectedBrush = NewBrush(0xEF, 0xF6, 0xFC);

    public static void Apply(Window window, string? theme)
    {
        bool dark = string.Equals(theme, "dark", StringComparison.OrdinalIgnoreCase);
        ApplyThemeResources(dark);

        // keep fallback direct coloring for controls with hardcoded locals
        ApplyRecursive(window, dark);
    }

    private static void ApplyThemeResources(bool dark)
    {
        var resources = System.Windows.Application.Current?.Resources;
        if (resources == null) return;

        resources["ThemeWindowBrush"] = dark ? DarkWindowBrush : LightWindowBrush;
        resources["ThemeSurfaceBrush"] = dark ? DarkSurfaceBrush : LightSurfaceBrush;
        resources["ThemeSurfaceAltBrush"] = dark ? DarkSurfaceAltBrush : LightSurfaceAltBrush;
        resources["ThemeBorderBrush"] = dark ? DarkBorderBrush : LightBorderBrush;
        resources["ThemeTextBrush"] = dark ? DarkTextBrush : LightTextBrush;
        resources["ThemeMutedTextBrush"] = dark ? DarkMutedTextBrush : LightMutedTextBrush;
        resources["ThemeTabSelectedBrush"] = dark ? DarkTabSelectedBrush : LightTabSelectedBrush;
        resources["ThemeTabHoverBrush"] = dark ? DarkTabHoverBrush : LightTabHoverBrush;
        resources["ThemeListSelectedBrush"] = dark ? DarkListSelectedBrush : LightListSelectedBrush;

        // Default control/popup system keys used by WPF templates.
        resources[System.Windows.SystemColors.WindowBrushKey] = dark ? DarkSurfaceBrush : LightSurfaceBrush;
        resources[System.Windows.SystemColors.WindowTextBrushKey] = dark ? DarkTextBrush : LightTextBrush;
        resources[System.Windows.SystemColors.ControlBrushKey] = dark ? DarkSurfaceAltBrush : LightSurfaceAltBrush;
        resources[System.Windows.SystemColors.ControlTextBrushKey] = dark ? DarkTextBrush : LightTextBrush;
        resources[System.Windows.SystemColors.HighlightBrushKey] = dark ? DarkListSelectedBrush : LightListSelectedBrush;
        resources[System.Windows.SystemColors.HighlightTextBrushKey] = dark ? DarkTextBrush : LightTextBrush;
        resources[System.Windows.SystemColors.InactiveSelectionHighlightBrushKey] = dark ? DarkSurfaceAltBrush : LightSurfaceAltBrush;
        resources[System.Windows.SystemColors.InactiveSelectionHighlightTextBrushKey] = dark ? DarkTextBrush : LightTextBrush;

        // Additional keys used by ComboBox default control template gradients/borders.
        resources[System.Windows.SystemColors.ControlLightBrushKey] = dark ? DarkSurfaceBrush : LightSurfaceBrush;
        resources[System.Windows.SystemColors.ControlLightLightBrushKey] = dark ? DarkSurfaceBrush : LightSurfaceBrush;
        resources[System.Windows.SystemColors.ControlDarkBrushKey] = dark ? DarkBorderBrush : LightBorderBrush;
        resources[System.Windows.SystemColors.ControlDarkDarkBrushKey] = dark ? DarkBorderBrush : LightBorderBrush;
        resources[System.Windows.SystemColors.GrayTextBrushKey] = dark ? DarkMutedTextBrush : LightMutedTextBrush;
    }

    private static void ApplyRecursive(DependencyObject root, bool dark)
    {
        switch (root)
        {
            case Window w:
                SetOrClear(w, Window.BackgroundProperty, dark ? DarkWindowBrush : null);
                SetOrClear(w, Window.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.Border border:
                if (border.Name is not "RecordingBadge")
                {
                    if (ShouldOverrideBrush(border.Background))
                        SetOrClear(border, System.Windows.Controls.Border.BackgroundProperty, dark ? DarkSurfaceAltBrush : null);

                    if (ShouldOverrideBrush(border.BorderBrush))
                        SetOrClear(border, System.Windows.Controls.Border.BorderBrushProperty, dark ? DarkBorderBrush : null);
                }
                break;

            case System.Windows.Controls.Panel panel:
                if (ShouldOverrideBrush(panel.Background))
                    SetOrClear(panel, System.Windows.Controls.Panel.BackgroundProperty, dark ? DarkWindowBrush : null);
                break;

            case System.Windows.Controls.TextBox tb:
                SetOrClear(tb, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceBrush : null);
                SetOrClear(tb, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                SetOrClear(tb, System.Windows.Controls.Control.BorderBrushProperty, dark ? DarkBorderBrush : null);
                break;

            case System.Windows.Controls.PasswordBox pb:
                SetOrClear(pb, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceBrush : null);
                SetOrClear(pb, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                SetOrClear(pb, System.Windows.Controls.Control.BorderBrushProperty, dark ? DarkBorderBrush : null);
                break;

            case System.Windows.Controls.ComboBox cb:
                SetOrClear(cb, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceBrush : null);
                SetOrClear(cb, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                SetOrClear(cb, System.Windows.Controls.Control.BorderBrushProperty, dark ? DarkBorderBrush : null);
                break;

            case System.Windows.Controls.ComboBoxItem cbi:
                SetOrClear(cbi, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceAltBrush : null);
                SetOrClear(cbi, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.ListBox lb:
                SetOrClear(lb, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceAltBrush : null);
                SetOrClear(lb, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                SetOrClear(lb, System.Windows.Controls.Control.BorderBrushProperty, dark ? DarkBorderBrush : null);
                break;

            case System.Windows.Controls.ListBoxItem lbi:
                SetOrClear(lbi, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkSurfaceAltBrush : null);
                SetOrClear(lbi, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.TabControl tc:
                SetOrClear(tc, System.Windows.Controls.Control.BackgroundProperty, dark ? DarkWindowBrush : null);
                SetOrClear(tc, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.TabItem ti:
                SetOrClear(ti, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.GroupBox gb:
                SetOrClear(gb, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                SetOrClear(gb, System.Windows.Controls.Control.BorderBrushProperty, dark ? DarkBorderBrush : null);
                break;

            case System.Windows.Controls.Label label:
                SetOrClear(label, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.CheckBox checkBox:
                SetOrClear(checkBox, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.RadioButton radioButton:
                SetOrClear(radioButton, System.Windows.Controls.Control.ForegroundProperty, dark ? DarkTextBrush : null);
                break;

            case System.Windows.Controls.TextBlock textBlock when textBlock.Name is not "RecBadgeText" and not "StepCountBadge":
                if (ShouldOverrideBrush(textBlock.Foreground))
                    SetOrClear(textBlock, System.Windows.Controls.TextBlock.ForegroundProperty, dark ? DarkTextBrush : null);
                break;
        }

        foreach (var child in LogicalTreeHelper.GetChildren(root))
        {
            if (child is DependencyObject dependencyObject)
                ApplyRecursive(dependencyObject, dark);
        }
    }

    private static bool ShouldOverrideBrush(System.Windows.Media.Brush? brush)
    {
        if (brush is null)
            return true;

        if (brush is not System.Windows.Media.SolidColorBrush solid)
            return false;

        var c = solid.Color;

        // Always allow reverting previously applied dark brushes.
        if (c == DarkWindowBrush.Color || c == DarkSurfaceBrush.Color || c == DarkSurfaceAltBrush.Color || c == DarkBorderBrush.Color || c == DarkTextBrush.Color)
            return true;

        // Keep accent colors (blue/red) unchanged; only replace neutral grays.
        bool looksAccent = c.B > c.R + 20 || c.R > c.G + 35;
        if (looksAccent)
            return false;

        int max = Math.Max(c.R, Math.Max(c.G, c.B));
        int min = Math.Min(c.R, Math.Min(c.G, c.B));
        bool nearNeutral = (max - min) <= 18;
        bool nearWhiteOrGray = max >= 140;
        bool nearBlackOrGray = max <= 120;

        return nearNeutral && (nearWhiteOrGray || nearBlackOrGray);
    }

    private static void SetOrClear(DependencyObject target, DependencyProperty property, object? value)
    {
        if (value is null)
            target.ClearValue(property);
        else
            target.SetValue(property, value);
    }

    private static System.Windows.Media.SolidColorBrush NewBrush(byte r, byte g, byte b)
    {
        var brush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }
}