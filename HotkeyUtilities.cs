using System.Collections.Generic;
using System.Windows.Input;

namespace AMTool;

internal static class HotkeyUtilities
{
    private const ModifierKeys SupportedModifiers =
        ModifierKeys.Control | ModifierKeys.Shift | ModifierKeys.Alt | ModifierKeys.Windows;

    internal static ModifierKeys SanitizeModifiers(ModifierKeys modifiers)
    {
        return modifiers & SupportedModifiers;
    }

    internal static bool IsValidHotkey(ModifierKeys modifiers, Key key)
    {
        return IsValidHotkey(modifiers, key, HotkeyInputKind.Keyboard, HotkeyMouseButton.None);
    }

    internal static bool IsValidHotkey(
        ModifierKeys modifiers,
        Key key,
        HotkeyInputKind inputKind,
        HotkeyMouseButton mouseButton)
    {
        return inputKind switch
        {
            HotkeyInputKind.Keyboard => !IsModifierOnlyKey(key) && key != Key.None,
            HotkeyInputKind.MouseButton => mouseButton != HotkeyMouseButton.None,
            _ => false
        };
    }

    internal static string FormatHotkey(ModifierKeys modifiers, Key key)
    {
        return FormatHotkey(modifiers, key, HotkeyInputKind.Keyboard, HotkeyMouseButton.None);
    }

    internal static string FormatHotkey(
        ModifierKeys modifiers,
        Key key,
        HotkeyInputKind inputKind,
        HotkeyMouseButton mouseButton)
    {
        var parts = new List<string>(5);
        ModifierKeys sanitizedModifiers = SanitizeModifiers(modifiers);

        if (sanitizedModifiers.HasFlag(ModifierKeys.Control))
        {
            parts.Add("Ctrl");
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Shift))
        {
            parts.Add("Shift");
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Alt))
        {
            parts.Add("Alt");
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Windows))
        {
            parts.Add("Win");
        }

        string triggerLabel = inputKind switch
        {
            HotkeyInputKind.Keyboard when key != Key.None => GetKeyLabel(key),
            HotkeyInputKind.MouseButton when mouseButton != HotkeyMouseButton.None => GetMouseButtonLabel(mouseButton),
            _ => string.Empty
        };

        if (!string.IsNullOrWhiteSpace(triggerLabel))
        {
            parts.Add(triggerLabel);
        }

        return parts.Count == 0 ? "Brak" : string.Join("+", parts);
    }

    internal static string GetKeyLabel(Key key)
    {
        return key switch
        {
            >= Key.D0 and <= Key.D9 => ((char)('0' + (key - Key.D0))).ToString(),
            >= Key.NumPad0 and <= Key.NumPad9 => $"Num{key - Key.NumPad0}",
            Key.Next => "PageDown",
            Key.Prior => "PageUp",
            Key.Return => "Enter",
            Key.Escape => "Esc",
            _ => key.ToString()
        };
    }

    internal static bool IsModifierOnlyKey(Key key)
    {
        return key is Key.LeftCtrl
            or Key.RightCtrl
            or Key.LeftShift
            or Key.RightShift
            or Key.LeftAlt
            or Key.RightAlt
            or Key.LWin
            or Key.RWin
            or Key.System;
    }

    internal static string GetMouseButtonLabel(HotkeyMouseButton mouseButton)
    {
        return mouseButton switch
        {
            HotkeyMouseButton.Left => "LPM",
            HotkeyMouseButton.Right => "PPM",
            HotkeyMouseButton.Middle => "Srodkowy",
            HotkeyMouseButton.XButton1 => "Mouse4",
            HotkeyMouseButton.XButton2 => "Mouse5",
            _ => string.Empty
        };
    }

    internal static HotkeyMouseButton? ToHotkeyMouseButton(MouseButton mouseButton)
    {
        return mouseButton switch
        {
            MouseButton.Left => HotkeyMouseButton.Left,
            MouseButton.Right => HotkeyMouseButton.Right,
            MouseButton.Middle => HotkeyMouseButton.Middle,
            MouseButton.XButton1 => HotkeyMouseButton.XButton1,
            MouseButton.XButton2 => HotkeyMouseButton.XButton2,
            _ => null
        };
    }
}
