using System;

namespace AMTool;

internal static class ShortcutAssignmentUtilities
{
    private const int MaxVisibleOrbitTiles = 5;
    private const int PreferredVisibleSlotIndex = 2;

    internal static int GetScrollIndexToRevealShortcut(int shortcutCount, int shortcutIndex)
    {
        if (shortcutCount <= 0 || shortcutIndex < 0)
        {
            return 0;
        }

        int totalTileCount = shortcutCount + 1;
        int maxScrollIndex = totalTileCount > MaxVisibleOrbitTiles
            ? totalTileCount - 1
            : 0;
        int preferredScrollIndex = shortcutIndex - PreferredVisibleSlotIndex;

        return Math.Clamp(preferredScrollIndex, 0, maxScrollIndex);
    }
}
