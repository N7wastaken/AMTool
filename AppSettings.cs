using System.Windows.Input;

namespace AMTool;

public sealed class AppSettings
{
    public ModifierKeys HotkeyModifiers { get; set; } = ModifierKeys.Control | ModifierKeys.Shift;

    public Key HotkeyKey { get; set; } = Key.Q;

    public bool AutoStartEnabled { get; set; }
}
