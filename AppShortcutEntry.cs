namespace AMTool;

public sealed class AppShortcutEntry
{
    public string DisplayName { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public string ShortcutPath { get; set; } = string.Empty;

    public string? TargetPath { get; set; }
}
