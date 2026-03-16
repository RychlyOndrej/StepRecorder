using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace StepRecorder.Services;

/// <summary>
/// Installs WH_MOUSE_LL and WH_KEYBOARD_LL global hooks.
/// Must be started on the UI thread (has a Windows message pump).
/// </summary>
public sealed class GlobalHookService : IDisposable
{
    // Delegates kept alive to prevent GC collection
    private NativeMethods.LowLevelMouseProc?    _mouseProc;
    private NativeMethods.LowLevelKeyboardProc? _keyboardProc;

    private IntPtr _mouseHookId    = IntPtr.Zero;
    private IntPtr _keyboardHookId = IntPtr.Zero;

    private readonly StringBuilder _keyBuffer    = new();
    private DateTime               _lastKeyTime  = DateTime.MinValue;
    private const int              KeyTimeoutMs  = 2000;

    private bool _active;

    // ── Events ────────────────────────────────────────────────────────────

    /// <summary>Fired on the UI thread when the left mouse button is pressed.</summary>
    public event Action<int, int>? LeftButtonDown;   // screenX, screenY

    // ── Start / Stop ─────────────────────────────────────────────────────

    public void Start()
    {
        if (_active) return;

        _mouseProc    = MouseCallback;
        _keyboardProc = KeyboardCallback;

        using var proc   = Process.GetCurrentProcess();
        using var module = proc.MainModule!;
        IntPtr hMod      = NativeMethods.GetModuleHandle(module.ModuleName!);

        _mouseHookId    = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL,    _mouseProc,    hMod, 0);
        _keyboardHookId = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, _keyboardProc, hMod, 0);

        _active = true;
    }

    public void Stop()
    {
        if (!_active) return;

        if (_mouseHookId    != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_mouseHookId);    _mouseHookId    = IntPtr.Zero; }
        if (_keyboardHookId != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_keyboardHookId); _keyboardHookId = IntPtr.Zero; }

        _active = false;
    }

    // ── Key buffer ────────────────────────────────────────────────────────

    /// <summary>
    /// Returns and clears the current key buffer.
    /// Call this just before creating a RecordingStep.
    /// </summary>
    public string FlushKeys()
    {
        var result = _keyBuffer.ToString();
        _keyBuffer.Clear();
        return result;
    }

    // ── Hook callbacks ────────────────────────────────────────────────────

    private IntPtr MouseCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)NativeMethods.WM_LBUTTONDOWN)
        {
            var s = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
            LeftButtonDown?.Invoke(s.pt.X, s.pt.Y);
        }
        return NativeMethods.CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
    }

    private IntPtr KeyboardCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 &&
            (wParam == (IntPtr)NativeMethods.WM_KEYDOWN ||
             wParam == (IntPtr)NativeMethods.WM_SYSKEYDOWN))
        {
            var s  = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);
            var vk = (System.Windows.Input.Key)
                     System.Windows.Input.KeyInterop.KeyFromVirtualKey((int)s.vkCode);

            // Skip pure modifier keys
            if (IsModifier(vk))
                return NativeMethods.CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);

            var now = DateTime.Now;

            // Auto-flush stale buffer after timeout
            if ((now - _lastKeyTime).TotalMilliseconds > KeyTimeoutMs && _keyBuffer.Length > 0)
                _keyBuffer.Clear();

            _lastKeyTime = now;
            AppendKey(vk, s.vkCode);
        }

        return NativeMethods.CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static bool IsModifier(System.Windows.Input.Key k) =>
        k is System.Windows.Input.Key.LeftCtrl  or System.Windows.Input.Key.RightCtrl  or
             System.Windows.Input.Key.LeftShift or System.Windows.Input.Key.RightShift or
             System.Windows.Input.Key.LeftAlt   or System.Windows.Input.Key.RightAlt   or
             System.Windows.Input.Key.LWin      or System.Windows.Input.Key.RWin       or
             System.Windows.Input.Key.System;

    private void AppendKey(System.Windows.Input.Key k, uint vk)
    {
        switch (k)
        {
            case System.Windows.Input.Key.Back:
                if (_keyBuffer.Length > 0) _keyBuffer.Remove(_keyBuffer.Length - 1, 1);
                break;
            case System.Windows.Input.Key.Return: _keyBuffer.Append("↵"); break;
            case System.Windows.Input.Key.Tab:    _keyBuffer.Append("⇥"); break;
            case System.Windows.Input.Key.Space:  _keyBuffer.Append(" ");  break;
            case System.Windows.Input.Key.Escape: _keyBuffer.Append("[Esc]"); break;
            case System.Windows.Input.Key.Delete: _keyBuffer.Append("[Del]"); break;
            default:
                string keyStr = k.ToString();
                if (keyStr.Length == 1)
                    _keyBuffer.Append(keyStr);
                else if (keyStr.StartsWith("D") && keyStr.Length == 2 && char.IsDigit(keyStr[1]))
                    _keyBuffer.Append(keyStr[1]);
                else if (keyStr.StartsWith("NumPad") && keyStr.Length == 7)
                    _keyBuffer.Append(keyStr[6]);
                else
                    _keyBuffer.Append($"[{keyStr}]");
                break;
        }
    }

    public void Dispose() => Stop();
}
