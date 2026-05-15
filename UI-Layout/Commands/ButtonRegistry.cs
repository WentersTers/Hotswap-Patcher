using PAIcomPatcher.UI.Configuration;
using PAIcomPatcher.UI.Layout;

namespace PAIcomPatcher.UI.Commands;

public sealed class ButtonRegistry
{
    private readonly UIProfileManager _profileManager;
    private readonly LayoutManager _layoutManager;

    public ButtonRegistry(UIProfileManager profileManager, LayoutManager layoutManager)
    {
        _profileManager = profileManager;
        _layoutManager = layoutManager;
    }

    public CommandButtonDefinition GetButtonMetadata(string phrase)
    {
        return _layoutManager.GetAllButtons().First(button => string.Equals(button.Phrase, phrase, StringComparison.OrdinalIgnoreCase));
    }

    public bool TryGetButtonMetadata(string phrase, out CommandButtonDefinition? definition)
    {
        definition = _layoutManager.GetAllButtons().FirstOrDefault(button => string.Equals(button.Phrase, phrase, StringComparison.OrdinalIgnoreCase));
        return definition != null;
    }

    public void UpdateButtonMetadata(string phrase, ButtonConfig metadata)
    {
        _profileManager.SaveCustomButtonMetadata(phrase, metadata);
    }

    public void RemoveCustomization(string phrase)
    {
        _profileManager.RemoveCustomButtonMetadata(phrase);
    }
}