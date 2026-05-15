using PAIcomPatcher.UI.Configuration;

namespace PAIcomPatcher.UI.Layout;

public sealed class LayoutManager
{
    private readonly UIProfileManager _profileManager;
    private readonly IReadOnlyList<CommandEntry> _commands;

    public LayoutManager(UIProfileManager profileManager, IReadOnlyList<CommandEntry> commands)
    {
        _profileManager = profileManager;
        _commands = commands;
    }

    public IReadOnlyList<CommandButtonDefinition> GetAllButtons()
    {
        var profile = _profileManager.LoadProfile();
        var buttons = new List<CommandButtonDefinition>(_commands.Count);

        foreach (var command in _commands)
        {
            var isCustomized = profile.CustomButtons.ContainsKey(command.Phrase);
            var metadata = _profileManager.GetButtonMetadata(command.Phrase);
            metadata.Position = metadata.Position < 0 ? command.Index : metadata.Position;

            buttons.Add(new CommandButtonDefinition
            {
                Phrase = command.Phrase,
                CommandName = command.CommandName,
                Config = metadata,
                IsCustomized = isCustomized,
            });
        }

        return buttons;
    }

    public Dictionary<string, List<CommandButtonDefinition>> GetButtonsByTab()
    {
        var buttons = GetAllButtons();
        var grouped = new Dictionary<string, List<CommandButtonDefinition>>(StringComparer.OrdinalIgnoreCase);

        foreach (var button in buttons.Where(button => button.Config.Visible))
        {
            if (!grouped.TryGetValue(button.Config.GroupTab, out var list))
            {
                list = new List<CommandButtonDefinition>();
                grouped[button.Config.GroupTab] = list;
            }

            list.Add(button);
        }

        foreach (var list in grouped.Values)
        {
            list.Sort(CompareButtons);
        }

        return grouped;
    }

    public List<string> GetTabDisplayOrder()
    {
        var profile = _profileManager.LoadProfile();
        var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        if (!order.ContainsKey("General"))
            order["General"] = 0;

        foreach (var pair in profile.Tabs)
        {
            order[pair.Key] = pair.Value.DisplayOrder;
        }

        var discoveryIndex = order.Count + 1;
        foreach (var button in GetAllButtons())
        {
            if (!order.ContainsKey(button.Config.GroupTab))
            {
                order[button.Config.GroupTab] = discoveryIndex++;
            }
        }

        return order
            .OrderBy(pair => pair.Value)
            .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => pair.Key)
            .ToList();
    }

    public string GetButtonDisplayName(string phrase)
    {
        return _profileManager.GetButtonMetadata(phrase).DisplayName;
    }

    public (string Tab, int Index) GetButtonPosition(string phrase)
    {
        var button = GetAllButtons().FirstOrDefault(button => string.Equals(button.Phrase, phrase, StringComparison.OrdinalIgnoreCase));
        if (button == null)
            return ("General", 0);

        var grouped = GetButtonsByTab();
        if (!grouped.TryGetValue(button.Config.GroupTab, out var list))
            return (button.Config.GroupTab, 0);

        var index = list.FindIndex(candidate => string.Equals(candidate.Phrase, phrase, StringComparison.OrdinalIgnoreCase));
        return (button.Config.GroupTab, Math.Max(0, index));
    }

    private static int CompareButtons(CommandButtonDefinition left, CommandButtonDefinition right)
    {
        var position = left.Config.Position.CompareTo(right.Config.Position);
        if (position != 0)
            return position;

        return string.Compare(left.Phrase, right.Phrase, StringComparison.OrdinalIgnoreCase);
    }
}