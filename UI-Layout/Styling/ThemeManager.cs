using PAIcomPatcher.UI.Configuration;

namespace PAIcomPatcher.UI.Styling;

public sealed class ThemeManager
{
    private readonly UIProfile _profile;

    public ThemeManager(UIProfile profile)
    {
        _profile = profile;
    }

    public Color BackgroundColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.BackgroundColor, Color.FromArgb(45, 45, 48));
    public Color TextColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.TextColor, Color.White);
    public Color ButtonBackgroundColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.ButtonBackgroundColor, Color.FromArgb(63, 63, 70));
    public Color ButtonHoverColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.ButtonHoverColor, Color.FromArgb(80, 80, 90));
    public Color ButtonBorderColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.ButtonBorderColor, Color.FromArgb(109, 109, 109));
    public Color AccentColor => ColorUtilities.ParseColorOrDefault(_profile.Appearance.AccentColor, Color.FromArgb(14, 99, 156));
    public int ButtonBorderSize => Math.Max(0, _profile.Appearance.ButtonBorderSize);

    public void ApplyToForm(Form form)
    {
        form.BackColor = BackgroundColor;
        form.ForeColor = TextColor;
    }

    public Color ResolveButtonBackColor(ButtonConfig config)
    {
        return ColorUtilities.TryParseColor(config.BackgroundColor, out var custom)
            ? custom
            : ButtonBackgroundColor;
    }

    public Color ResolveTabBackColor(string? tabColor)
    {
        return ColorUtilities.ParseColorOrDefault(tabColor, BackgroundColor);
    }
}