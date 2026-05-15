using PAIcomPatcher.UI.Configuration;

namespace PAIcomPatcher.UI.Layout;

public sealed class CommandButtonDefinition
{
    public required string Phrase { get; init; }
    public required string CommandName { get; init; }
    public required ButtonConfig Config { get; init; }
    public required bool IsCustomized { get; init; }
}