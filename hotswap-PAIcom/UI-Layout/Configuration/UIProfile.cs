using System.Collections.Generic;

namespace PAIcomPatcher.UI.Configuration;

public sealed class UIProfile
{
    public string Version { get; set; } = "1.0";
    public ProfileMetadata Metadata { get; set; } = new();
    public AppearanceSettings Appearance { get; set; } = new();
    public LayoutSettings Layout { get; set; } = new();
    public Dictionary<string, ButtonConfig> CustomButtons { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, TabConfig> Tabs { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class ProfileMetadata
{
    public string Name { get; set; } = "Default Profile";
    public string Description { get; set; } = "Factory default theme and layout";
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public DateTime Modified { get; set; } = DateTime.UtcNow;
}

public sealed class AppearanceSettings
{
    public string BackgroundColor { get; set; } = "#2D2D30";
    public string TextColor { get; set; } = "#FFFFFF";
    public string ButtonBackgroundColor { get; set; } = "#3F3F46";
    public string ButtonHoverColor { get; set; } = "#505090";
    public string ButtonBorderColor { get; set; } = "#6D6D6D";
    public int ButtonBorderSize { get; set; } = 1;
    public string AccentColor { get; set; } = "#0E639C";
}

public sealed class LayoutSettings
{
    public int WindowWidth { get; set; } = 850;
    public int WindowHeight { get; set; } = 600;
    public int ButtonWidth { get; set; } = 140;
    public int ButtonHeight { get; set; } = 52;
    public int ButtonPadding { get; set; } = 6;
    public bool PreserveAspectRatio { get; set; }
}

public sealed class ButtonConfig
{
    public string DisplayName { get; set; } = string.Empty;
    public string GroupTab { get; set; } = "General";
    public int Position { get; set; }
    public bool Visible { get; set; } = true;
    public string? BackgroundColor { get; set; }
    public string? Notes { get; set; }
}

public sealed class TabConfig
{
    public int DisplayOrder { get; set; }
    public string? BackgroundColor { get; set; }
}

public sealed record CommandEntry(string Phrase, string CommandName, int Index);