namespace AMTool;

internal static class WindowVisibilityUtilities
{
    internal static bool CanToggleVisibilityFromHotkey(int blockingInteractionDepth)
    {
        return blockingInteractionDepth == 0;
    }

    internal static bool ShouldDeferHideUntilAnimationCompletes(bool isVisibilityAnimationRunning)
    {
        return isVisibilityAnimationRunning;
    }

    internal static bool CanHideAfterDeactivation(
        bool isExitRequested,
        bool isVisible,
        bool isActive,
        bool isVisibilityAnimationRunning,
        int blockingInteractionDepth,
        bool isAnyShortcutContextMenuOpen)
    {
        return !isExitRequested
            && isVisible
            && !isActive
            && !isVisibilityAnimationRunning
            && blockingInteractionDepth == 0
            && !isAnyShortcutContextMenuOpen;
    }
}
