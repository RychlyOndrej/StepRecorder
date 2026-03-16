using System.Windows;
using System.Windows.Input;

namespace StepRecorder;

public partial class HotkeyPickerDialog : Window
{
    private HotkeyDefinition? _pending;

    public HotkeyDefinition? Result { get; private set; }

    public HotkeyPickerDialog(HotkeyDefinition current)
    {
        InitializeComponent();

        // Show current hotkey
        HotkeyPreviewText.Text  = current.DisplayText;
        HotkeyPreviewText.Foreground = System.Windows.Media.Brushes.Black;

        // Capture keys in the border area
        HotkeyBorder.MouseLeftButtonDown += (_, _) => Focus();
        PreviewKeyDown += Dialog_PreviewKeyDown;
    }

    private void Dialog_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        e.Handled = true;

        // Resolve actual key (handle Alt → System)
        Key k = e.Key == Key.System ? e.SystemKey : e.Key;

        // Skip pure modifiers
        if (k is Key.LeftCtrl or Key.RightCtrl
               or Key.LeftShift or Key.RightShift
               or Key.LeftAlt or Key.RightAlt
               or Key.LWin or Key.RWin)
            return;

        // Build modifier flags
        uint mods = 0;
        if (Keyboard.IsKeyDown(Key.LeftCtrl)  || Keyboard.IsKeyDown(Key.RightCtrl))  mods |= NativeMethods.MOD_CONTROL;
        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) mods |= NativeMethods.MOD_SHIFT;
        if (Keyboard.IsKeyDown(Key.LeftAlt)   || Keyboard.IsKeyDown(Key.RightAlt))   mods |= NativeMethods.MOD_ALT;

        uint vk = (uint)KeyInterop.VirtualKeyFromKey(k);

        // Require at least one modifier
        if (mods == 0)
        {
            HotkeyErrorText.Text       = "Kombinace musí obsahovat Ctrl, Shift nebo Alt.";
            HotkeyErrorText.Visibility = Visibility.Visible;
            OkBtn.IsEnabled            = false;
            return;
        }

        _pending = new HotkeyDefinition { Modifiers = mods, VirtualKey = vk };

        HotkeyPreviewText.Text       = _pending.DisplayText;
        HotkeyPreviewText.Foreground = System.Windows.Media.Brushes.Black;
        HotkeyErrorText.Visibility   = Visibility.Collapsed;
        OkBtn.IsEnabled              = true;
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        Result       = _pending;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) =>
        DialogResult = false;
}
