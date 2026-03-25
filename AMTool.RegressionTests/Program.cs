using System;
using System.Windows;
using System.Windows.Input;
using AMTool;

return RegressionTestRunner.RunAll();

internal static class RegressionTestRunner
{
    private static int _failures;

    internal static int RunAll()
    {
        Run(nameof(FormatHotkey_IncludesModifiersAndKey), FormatHotkey_IncludesModifiersAndKey);
        Run(nameof(IsValidHotkey_AcceptsSingleKeyAndRejectsModifierOnly), IsValidHotkey_AcceptsSingleKeyAndRejectsModifierOnly);
        Run(nameof(GetHoverBubbleText_ReturnsShortcutNameOrAddLabel), GetHoverBubbleText_ReturnsShortcutNameOrAddLabel);
        Run(nameof(CalculateHoverBubblePosition_ClampsWithinScene), CalculateHoverBubblePosition_ClampsWithinScene);
        Run(nameof(BuildInfoTooltip_UsesCorrectScrollMessage), BuildInfoTooltip_UsesCorrectScrollMessage);
        Run(nameof(VisibilityDeactivationBehavior_DefersDuringAnimationAndHidesAfterFocusLoss), VisibilityDeactivationBehavior_DefersDuringAnimationAndHidesAfterFocusLoss);

        if (_failures == 0)
        {
            Console.WriteLine("All regression checks passed.");
            return 0;
        }

        Console.Error.WriteLine($"{_failures} regression check(s) failed.");
        return 1;
    }

    private static void Run(string name, Action test)
    {
        try
        {
            test();
            Console.WriteLine($"PASS {name}");
        }
        catch (Exception ex)
        {
            _failures++;
            Console.Error.WriteLine($"FAIL {name}: {ex.Message}");
        }
    }

    private static void FormatHotkey_IncludesModifiersAndKey()
    {
        string formatted = HotkeyUtilities.FormatHotkey(ModifierKeys.Control | ModifierKeys.Shift, Key.Q);
        AssertEqual("Ctrl+Shift+Q", formatted);
    }

    private static void IsValidHotkey_AcceptsSingleKeyAndRejectsModifierOnly()
    {
        AssertTrue(HotkeyUtilities.IsValidHotkey(ModifierKeys.None, Key.F4));
        AssertFalse(HotkeyUtilities.IsValidHotkey(ModifierKeys.Control, Key.LeftCtrl));
    }

    private static void GetHoverBubbleText_ReturnsShortcutNameOrAddLabel()
    {
        string shortcutText = OrbitUiUtilities.GetHoverBubbleText(new AppShortcutEntry { DisplayName = "Notepad" });
        string addText = OrbitUiUtilities.GetHoverBubbleText(null);

        AssertEqual("Notepad", shortcutText);
        AssertEqual("Dodaj skrot", addText);
    }

    private static void CalculateHoverBubblePosition_ClampsWithinScene()
    {
        Point nearTopLeft = OrbitUiUtilities.CalculateHoverBubblePosition(
            new Point(0, 0),
            new Size(88, 88),
            new Size(140, 32),
            new Size(780, 270));

        Point nearRightEdge = OrbitUiUtilities.CalculateHoverBubblePosition(
            new Point(730, 20),
            new Size(88, 88),
            new Size(140, 32),
            new Size(780, 270));

        AssertEqual(0d, nearTopLeft.X);
        AssertEqual(100d, nearTopLeft.Y);
        AssertEqual(640d, nearRightEdge.X);
        AssertEqual(120d, nearRightEdge.Y);
    }

    private static void BuildInfoTooltip_UsesCorrectScrollMessage()
    {
        string compactTooltip = OrbitUiUtilities.BuildInfoTooltip("Ctrl+Shift+Q", autoStartEnabled: true, shortcutCount: 3, totalTileCount: 4, shortcutScrollIndex: 0);
        string scrollTooltip = OrbitUiUtilities.BuildInfoTooltip("Q", autoStartEnabled: false, shortcutCount: 7, totalTileCount: 8, shortcutScrollIndex: 2);

        AssertContains("Do 5 skrotow widac od razu bez przewijania.", compactTooltip);
        AssertContains("Widok 3 z 8.", scrollTooltip);
        AssertContains("Autostart: wylaczony.", scrollTooltip);
    }

    private static void VisibilityDeactivationBehavior_DefersDuringAnimationAndHidesAfterFocusLoss()
    {
        AssertTrue(WindowVisibilityUtilities.ShouldDeferHideUntilAnimationCompletes(isVisibilityAnimationRunning: true));
        AssertFalse(WindowVisibilityUtilities.CanHideAfterDeactivation(
            isExitRequested: false,
            isVisible: true,
            isActive: true,
            isVisibilityAnimationRunning: false,
            blockingInteractionDepth: 0,
            isAnyShortcutContextMenuOpen: false));
        AssertTrue(WindowVisibilityUtilities.CanHideAfterDeactivation(
            isExitRequested: false,
            isVisible: true,
            isActive: false,
            isVisibilityAnimationRunning: false,
            blockingInteractionDepth: 0,
            isAnyShortcutContextMenuOpen: false));
    }

    private static void AssertTrue(bool value)
    {
        if (!value)
        {
            throw new InvalidOperationException("Expected true.");
        }
    }

    private static void AssertFalse(bool value)
    {
        if (value)
        {
            throw new InvalidOperationException("Expected false.");
        }
    }

    private static void AssertEqual<T>(T expected, T actual)
    {
        if (!Equals(expected, actual))
        {
            throw new InvalidOperationException($"Expected '{expected}', got '{actual}'.");
        }
    }

    private static void AssertContains(string expectedSubstring, string actual)
    {
        if (!actual.Contains(expectedSubstring, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Expected substring '{expectedSubstring}' was not found.");
        }
    }
}
