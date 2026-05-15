using System.Drawing;

namespace PAIcomPatcher.UI.Styling;

public static class ColorUtilities
{
    public static Color ParseColorOrDefault(string? value, Color fallback)
    {
        if (TryParseColor(value, out var color))
            return color;

        return fallback;
    }

    public static bool TryParseColor(string? value, out Color color)
    {
        color = Color.Empty;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        var text = value.Trim();
        if (text.StartsWith('#'))
        {
            return TryParseHex(text, out color);
        }

        if (text.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) && text.EndsWith(')'))
        {
            return TryParseRgb(text, out color);
        }

        try
        {
            color = ColorTranslator.FromHtml(text);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static string ToHex(Color color)
    {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    private static bool TryParseHex(string value, out Color color)
    {
        color = Color.Empty;
        var hex = value[1..];
        if (hex.Length is not (6 or 8))
            return false;

        if (!int.TryParse(hex[..2], System.Globalization.NumberStyles.HexNumber, null, out var r))
            return false;
        if (!int.TryParse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g))
            return false;
        if (!int.TryParse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
            return false;

        var a = 255;
        if (hex.Length == 8)
        {
            if (!int.TryParse(hex.Substring(6, 2), System.Globalization.NumberStyles.HexNumber, null, out a))
                return false;
        }

        color = Color.FromArgb(a, r, g, b);
        return true;
    }

    private static bool TryParseRgb(string value, out Color color)
    {
        color = Color.Empty;
        var inner = value[4..^1];
        var parts = inner.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 3)
            return false;

        if (!int.TryParse(parts[0], out var r)) return false;
        if (!int.TryParse(parts[1], out var g)) return false;
        if (!int.TryParse(parts[2], out var b)) return false;

        var a = 255;
        if (parts.Length > 3 && !int.TryParse(parts[3], out a))
            return false;

        color = Color.FromArgb(a, r, g, b);
        return true;
    }
}