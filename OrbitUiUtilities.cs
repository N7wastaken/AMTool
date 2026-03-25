using System;

namespace AMTool;

internal static class OrbitUiUtilities
{
    private const string AddShortcutHoverText = "Dodaj skrot";
    private const int MaxVisibleOrbitTiles = 5;
    private const double HoverBubbleOffsetY = 12;

    internal static string GetHoverBubbleText(object? orbitItem)
    {
        return orbitItem is AppShortcutEntry shortcut
            ? shortcut.DisplayName
            : AddShortcutHoverText;
    }

    internal static System.Windows.Point CalculateHoverBubblePosition(
        System.Windows.Point tilePosition,
        System.Windows.Size tileSize,
        System.Windows.Size bubbleSize,
        System.Windows.Size sceneSize)
    {
        double x = tilePosition.X + ((tileSize.Width - bubbleSize.Width) / 2);
        double y = tilePosition.Y - bubbleSize.Height - HoverBubbleOffsetY;

        x = Math.Clamp(x, 0, Math.Max(0, sceneSize.Width - bubbleSize.Width));

        if (y < 0)
        {
            y = tilePosition.Y + tileSize.Height + HoverBubbleOffsetY;
        }

        y = Math.Clamp(y, 0, Math.Max(0, sceneSize.Height - bubbleSize.Height));
        return new System.Windows.Point(x, y);
    }

    internal static string BuildInfoTooltip(
        string hotkeyText,
        bool autoStartEnabled,
        int shortcutCount,
        int totalTileCount,
        int shortcutScrollIndex)
    {
        string scrollInfo = totalTileCount > MaxVisibleOrbitTiles
            ? $"Kolko myszy przewija kolejne skroty w petli. Widok {shortcutScrollIndex + 1} z {GetMaxShortcutScrollIndex(totalTileCount) + 1}."
            : "Do 5 skrotow widac od razu bez przewijania.";

        return string.Join(
            Environment.NewLine,
            $"{hotkeyText} wywoluje AMTool.",
            "LPM uruchamia aplikacje.",
            "PPM usuwa skrot z listy.",
            "PPM na ( i ) otwiera ustawienia.",
            $"Autostart: {(autoStartEnabled ? "wlaczony" : "wylaczony")}.",
            $"Liczba skrotow: {shortcutCount}.",
            scrollInfo);
    }

    private static int GetMaxShortcutScrollIndex(int totalTileCount)
    {
        return totalTileCount > MaxVisibleOrbitTiles
            ? totalTileCount - 1
            : 0;
    }
}
