using System.Windows.Input;

namespace AMTool;

public sealed class AppSettings
{
    public ModifierKeys HotkeyModifiers { get; set; } = ModifierKeys.Control | ModifierKeys.Shift;

    public HotkeyInputKind HotkeyInputKind { get; set; } = HotkeyInputKind.Keyboard;

    public Key HotkeyKey { get; set; } = Key.Q;

    public HotkeyMouseButton HotkeyMouseButton { get; set; }

    public bool AutoStartEnabled { get; set; }

    public bool HasCompletedTutorial { get; set; }
}
