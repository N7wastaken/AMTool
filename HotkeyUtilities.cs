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
        return !IsModifierOnlyKey(key) && key != Key.None;
    }

    internal static string FormatHotkey(ModifierKeys modifiers, Key key)
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

        if (key != Key.None)
        {
            parts.Add(GetKeyLabel(key));
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
}
