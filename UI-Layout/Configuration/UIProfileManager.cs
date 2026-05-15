using System.Text.Json;
using System.Text.Json.Serialization;

namespace PAIcomPatcher.UI.Configuration;

public sealed class UIProfileManager
{
    private readonly string _baseDir;
    private readonly string _profilePath;
    private readonly string _defaultProfilePath;
    private readonly string _commandsPath;
    private readonly JsonSerializerOptions _serializerOptions;
    private UIProfile? _currentProfile;
    private IReadOnlyList<CommandEntry>? _commandCache;

    public UIProfileManager(string baseDir)
    {
        _baseDir = baseDir;
        _profilePath = Path.Combine(baseDir, "custom-commands", "ui-profile.json");
        _defaultProfilePath = Path.Combine(baseDir, "custom-commands", "ui-profile.default.json");
        _commandsPath = Path.Combine(baseDir, "custom-commands", "commands.txt");
        _serializerOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNameCaseInsensitive = true,
        };
    }

    public string ProfilePath => _profilePath;
    public string DefaultProfilePath => _defaultProfilePath;
    public string CommandsPath => _commandsPath;

    public UIProfile LoadProfile()
    {
        if (_currentProfile != null)
            return _currentProfile;

        Directory.CreateDirectory(Path.GetDirectoryName(_profilePath)!);

        if (!File.Exists(_profilePath))
        {
            _currentProfile = GetDefaultProfile();
            SaveProfile(_currentProfile);
            WriteDefaultProfileTemplate();
            return _currentProfile;
        }

        try
        {
            var json = File.ReadAllText(_profilePath);
            var profile = JsonSerializer.Deserialize<UIProfile>(json, _serializerOptions) ?? GetDefaultProfile();
            _currentProfile = NormalizeProfile(profile);
            ValidateProfile(_currentProfile);
            return _currentProfile;
        }
        catch
        {
            _currentProfile = GetDefaultProfile();
            SaveProfile(_currentProfile);
            return _currentProfile;
        }
    }

    public IReadOnlyList<CommandEntry> LoadAllCommands()
    {
        if (_commandCache != null)
            return _commandCache;

        var commands = new List<CommandEntry>();
        if (!File.Exists(_commandsPath))
        {
            _commandCache = commands;
            return _commandCache;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var index = 0;

        foreach (var rawLine in File.ReadAllLines(_commandsPath))
        {
            var trimmed = rawLine.Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#'))
                continue;

            var (phrase, commandName) = ParseCommandLine(trimmed);
            if (seen.Add(phrase))
            {
                commands.Add(new CommandEntry(phrase, commandName, index));
                index++;
            }
        }

        _commandCache = commands;
        return _commandCache;
    }

    public ButtonConfig GetButtonMetadata(string phrase)
    {
        var profile = LoadProfile();
        if (profile.CustomButtons.TryGetValue(phrase, out var custom))
            return CloneButtonConfig(custom);

        var commandEntry = LoadAllCommands().FirstOrDefault(entry => string.Equals(entry.Phrase, phrase, StringComparison.OrdinalIgnoreCase));
        return CreateDefaultButtonConfig(phrase, commandEntry?.Index ?? 0);
    }

    public void SaveCustomButtonMetadata(string phrase, ButtonConfig metadata)
    {
        var profile = LoadProfile();
        profile.CustomButtons[phrase] = NormalizeButtonConfig(metadata, phrase, metadata.Position);
        profile.Metadata.Modified = DateTime.UtcNow;
        SaveProfile(profile);
    }

    public void RemoveCustomButtonMetadata(string phrase)
    {
        var profile = LoadProfile();
        if (profile.CustomButtons.Remove(phrase))
        {
            profile.Metadata.Modified = DateTime.UtcNow;
            SaveProfile(profile);
        }
    }

    public void ResetToDefault()
    {
        CreateBackup();

        var profile = LoadProfile();
        profile.CustomButtons.Clear();
        profile.Metadata.Modified = DateTime.UtcNow;
        SaveProfile(profile);
    }

    public void CreateBackup()
    {
        if (File.Exists(_profilePath))
        {
            File.Copy(_profilePath, _profilePath + ".backup", overwrite: true);
        }
    }

    public bool IsButtonCustomized(string phrase)
    {
        return LoadProfile().CustomButtons.ContainsKey(phrase);
    }

    public UIProfile GetDefaultProfile()
    {
        var profile = new UIProfile();
        profile.Metadata = new ProfileMetadata
        {
            Name = "Default Profile",
            Description = "Factory default theme and layout",
            Created = DateTime.UtcNow,
            Modified = DateTime.UtcNow,
        };
        profile.CustomButtons = new Dictionary<string, ButtonConfig>(StringComparer.OrdinalIgnoreCase);
        profile.Tabs = new Dictionary<string, TabConfig>(StringComparer.OrdinalIgnoreCase)
        {
            ["General"] = new TabConfig { DisplayOrder = 0 }
        };
        return profile;
    }

    public void SaveProfile(UIProfile profile)
    {
        var normalized = NormalizeProfile(profile);
        var json = JsonSerializer.Serialize(normalized, _serializerOptions);
        File.WriteAllText(_profilePath, json);
        _currentProfile = normalized;
    }

    public void RefreshCommandsCache()
    {
        _commandCache = null;
    }

    public ButtonConfig CreateDefaultButtonConfig(string phrase, int position)
    {
        return new ButtonConfig
        {
            DisplayName = GenerateDisplayName(phrase),
            GroupTab = "General",
            Position = position,
            Visible = true,
            BackgroundColor = null,
        };
    }

    public static string GenerateDisplayName(string phrase)
    {
        var stripped = phrase.Trim();
        const string prefix = "hey paicom ";
        if (stripped.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            stripped = stripped[prefix.Length..].Trim();

        stripped = stripped.Replace("[", string.Empty).Replace("]", string.Empty);
        stripped = stripped.Replace("(", string.Empty).Replace(")", string.Empty);
        stripped = string.Join(' ', stripped.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

        if (stripped.Length == 0)
            return "Command";

        return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(stripped.ToLowerInvariant());
    }

    private static ButtonConfig CloneButtonConfig(ButtonConfig source)
    {
        return new ButtonConfig
        {
            DisplayName = source.DisplayName,
            GroupTab = source.GroupTab,
            Position = source.Position,
            Visible = source.Visible,
            BackgroundColor = source.BackgroundColor,
            Notes = source.Notes,
        };
    }

    private UIProfile NormalizeProfile(UIProfile profile)
    {
        profile.CustomButtons = profile.CustomButtons != null
            ? new Dictionary<string, ButtonConfig>(profile.CustomButtons, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, ButtonConfig>(StringComparer.OrdinalIgnoreCase);

        profile.Tabs = profile.Tabs != null
            ? new Dictionary<string, TabConfig>(profile.Tabs, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, TabConfig>(StringComparer.OrdinalIgnoreCase);

        if (profile.Metadata == null)
            profile.Metadata = new ProfileMetadata();

        if (profile.Appearance == null)
            profile.Appearance = new AppearanceSettings();

        if (profile.Layout == null)
            profile.Layout = new LayoutSettings();

        if (!profile.Tabs.ContainsKey("General"))
            profile.Tabs["General"] = new TabConfig { DisplayOrder = 0 };

        foreach (var key in profile.CustomButtons.Keys.ToList())
        {
            profile.CustomButtons[key] = NormalizeButtonConfig(profile.CustomButtons[key], key, profile.CustomButtons[key].Position);
        }

        return profile;
    }

    private static ButtonConfig NormalizeButtonConfig(ButtonConfig config, string phrase, int fallbackPosition)
    {
        return new ButtonConfig
        {
            DisplayName = string.IsNullOrWhiteSpace(config.DisplayName)
                ? GenerateDisplayName(phrase)
                : config.DisplayName.Trim(),
            GroupTab = string.IsNullOrWhiteSpace(config.GroupTab)
                ? "General"
                : config.GroupTab.Trim(),
            Position = config.Position >= 0 ? config.Position : fallbackPosition,
            Visible = config.Visible,
            BackgroundColor = string.IsNullOrWhiteSpace(config.BackgroundColor) ? null : config.BackgroundColor.Trim(),
            Notes = string.IsNullOrWhiteSpace(config.Notes) ? null : config.Notes.Trim(),
        };
    }

    private void ValidateProfile(UIProfile profile)
    {
        if (!ProfileValidator.Validate(profile, out var errors))
        {
            var defaultProfile = GetDefaultProfile();
            defaultProfile.CustomButtons = profile.CustomButtons;
            defaultProfile.Tabs = profile.Tabs;
            defaultProfile.Metadata = profile.Metadata;
            defaultProfile.Appearance = profile.Appearance;
            defaultProfile.Layout = profile.Layout;
            _currentProfile = defaultProfile;
            SaveProfile(defaultProfile);
            return;
        }
    }

    private void WriteDefaultProfileTemplate()
    {
        if (!File.Exists(_defaultProfilePath))
        {
            var json = JsonSerializer.Serialize(GetDefaultProfile(), _serializerOptions);
            File.WriteAllText(_defaultProfilePath, json);
        }
    }

    private static (string Phrase, string CommandName) ParseCommandLine(string line)
    {
        var parenOpen = line.LastIndexOf('(');
        var parenClose = line.LastIndexOf(')');

        if (parenOpen > 0 && parenClose > parenOpen)
        {
            var phrase = line[..parenOpen].Trim();
            var fileRef = line[(parenOpen + 1)..parenClose].Trim();
            if (fileRef.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                fileRef = fileRef[..^4];
            return (phrase, fileRef);
        }

        return (line.Trim(), line.Trim());
    }
}