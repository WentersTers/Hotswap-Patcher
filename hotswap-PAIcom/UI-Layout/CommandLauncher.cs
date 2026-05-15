using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PAIcomPatcher.UI.Commands;
using PAIcomPatcher.UI.Components;
using PAIcomPatcher.UI.Configuration;
using PAIcomPatcher.UI.Layout;
using PAIcomPatcher.UI.Styling;

namespace PAIcomPatcher.UILayout;

/// <summary>
/// WinForms-based command launcher that creates a button grid from custom-commands/commands.txt.
/// If commands.txt is not found, falls back to scanning animations/ directory for .txt files.
/// Each button, when clicked, writes the command phrase to command_input.txt and dispatches it.
/// </summary>
public class CommandLauncher : Form
{
    private readonly string _baseDir;
    private readonly string _animationsDir;
    private readonly string _inputFilePath;
    private readonly Action<string>? _dispatchDelegate;

    private UIProfileManager _profileManager;
    private ThemeManager _themeManager;
    private LayoutManager _layoutManager;
    private ButtonRegistry _buttonRegistry;
    private UIProfile _profile;
    private IReadOnlyList<CommandEntry> _commands;

    private UITabControl _tabControl = null!;
    private Label _statusLabel = null!;
    private MenuStrip _menuStrip = null!;
    private ToolStripMenuItem _refreshMenuItem = null!;
    private ToolStripMenuItem _resetMenuItem = null!;
    private ToolStripMenuItem _exitMenuItem = null!;

    public CommandLauncher(string? baseDir = null, Action<string>? dispatchDelegate = null)
    {
        _baseDir = baseDir ?? GetDefaultBaseDir();
        _animationsDir = Path.Combine(_baseDir, "animations");
        _inputFilePath = Path.Combine(_baseDir, "command_input.txt");
        _dispatchDelegate = dispatchDelegate;

        _profileManager = new UIProfileManager(_baseDir);
        _themeManager = new ThemeManager(_profileManager.GetDefaultProfile());
        _layoutManager = new LayoutManager(_profileManager, Array.Empty<CommandEntry>());
        _buttonRegistry = new ButtonRegistry(_profileManager, _layoutManager);
        _profile = _profileManager.GetDefaultProfile();
        _commands = Array.Empty<CommandEntry>();

        InitializeComponent();
        RefreshCommands();
    }

    private string GetDefaultBaseDir()
    {
        var location = System.Reflection.Assembly.GetExecutingAssembly().Location;
        var dir = Path.GetDirectoryName(location) ?? AppDomain.CurrentDomain.BaseDirectory;

        var searchDir = new DirectoryInfo(dir);
        for (int i = 0; i < 5; i++)
        {
            if (searchDir == null)
                break;

            var commandsPath = Path.Combine(searchDir.FullName, "custom-commands", "commands.txt");
            if (File.Exists(commandsPath))
                return searchDir.FullName;

            searchDir = searchDir.Parent;
        }

        return dir;
    }

    private void InitializeComponent()
    {
        SuspendLayout();

        Text = "PAIcom Command Launcher";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        Font = new Font("Segoe UI", 9F);
        ClientSize = new Size(850, 600);
        MinimumSize = new Size(680, 420);

        _menuStrip = BuildMenuStrip();
        _tabControl = new UITabControl
        {
            Dock = DockStyle.Fill,
        };
        _statusLabel = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 28,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(10, 0, 0, 0),
        };

        Controls.Add(_tabControl);
        Controls.Add(_statusLabel);
        Controls.Add(_menuStrip);

        MainMenuStrip = _menuStrip;

        ResumeLayout(performLayout: true);
    }

    private MenuStrip BuildMenuStrip()
    {
        var menuStrip = new MenuStrip();

        _refreshMenuItem = new ToolStripMenuItem("Refresh", null, (_, _) => RefreshCommands());
        _resetMenuItem = new ToolStripMenuItem("Reset UI Settings", null, (_, _) => ResetCustomizations());
        _exitMenuItem = new ToolStripMenuItem("Exit", null, (_, _) => Close());

        var fileMenu = new ToolStripMenuItem("File");
        fileMenu.DropDownItems.Add(_refreshMenuItem);
        fileMenu.DropDownItems.Add(new ToolStripSeparator());
        fileMenu.DropDownItems.Add(_resetMenuItem);
        fileMenu.DropDownItems.Add(new ToolStripSeparator());
        fileMenu.DropDownItems.Add(_exitMenuItem);

        menuStrip.Items.Add(fileMenu);
        return menuStrip;
    }

    private void LoadLauncherState()
    {
        _profileManager = new UIProfileManager(_baseDir);
        _profile = _profileManager.LoadProfile();
        _themeManager = new ThemeManager(_profile);

        _profileManager.RefreshCommandsCache();
        _commands = _profileManager.LoadAllCommands();
        if (_commands.Count == 0)
            _commands = LoadFallbackCommands();

        _layoutManager = new LayoutManager(_profileManager, _commands);
        _buttonRegistry = new ButtonRegistry(_profileManager, _layoutManager);

        if (_profile.Layout.WindowWidth > 0 && _profile.Layout.WindowHeight > 0)
            ClientSize = new Size(_profile.Layout.WindowWidth, _profile.Layout.WindowHeight);

        Text = string.IsNullOrWhiteSpace(_profile.Metadata?.Name)
            ? "PAIcom Command Launcher"
            : $"PAIcom Command Launcher - {_profile.Metadata.Name}";
    }

    private IReadOnlyList<CommandEntry> LoadFallbackCommands()
    {
        if (!Directory.Exists(_animationsDir))
            return Array.Empty<CommandEntry>();

        var commands = new List<CommandEntry>();
        var index = 0;
        foreach (var filePath in Directory.GetFiles(_animationsDir, "*.txt"))
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var phrase = $"hey paicom {fileName}";
            commands.Add(new CommandEntry(phrase, fileName, index));
            index++;
        }

        return commands;
    }

    private void ApplyTheme()
    {
        _themeManager.ApplyToForm(this);
        _tabControl.TabBackColor = _themeManager.BackgroundColor;
        _tabControl.TabSelectedBackColor = _themeManager.AccentColor;
        _tabControl.TabTextColor = _themeManager.TextColor;
        _menuStrip.BackColor = _themeManager.BackgroundColor;
        _menuStrip.ForeColor = _themeManager.TextColor;
        _statusLabel.BackColor = ColorUtilities.ParseColorOrDefault("#3C3C41", _themeManager.BackgroundColor);
        _statusLabel.ForeColor = _themeManager.TextColor;
    }

    private void RenderButtons()
    {
        _tabControl.SuspendLayout();
        _tabControl.TabPages.Clear();

        var tabOrder = _layoutManager.GetTabDisplayOrder();
        var groupedButtons = _layoutManager.GetButtonsByTab();

        if (tabOrder.Count == 0)
            tabOrder.Add("General");

        foreach (var tabName in tabOrder)
        {
            groupedButtons.TryGetValue(tabName, out var buttonsInTab);
            var page = BuildTabPage(tabName, buttonsInTab ?? new List<CommandButtonDefinition>());
            _tabControl.TabPages.Add(page);
        }

        if (_tabControl.TabPages.Count == 0)
        {
            _tabControl.TabPages.Add(BuildTabPage("General", Array.Empty<CommandButtonDefinition>()));
        }

        _tabControl.ResumeLayout(performLayout: true);
    }

    private TabPage BuildTabPage(string tabName, IReadOnlyList<CommandButtonDefinition> buttons)
    {
        var tabBackground = _profile.Tabs.TryGetValue(tabName, out var tabConfig)
            ? _themeManager.ResolveTabBackColor(tabConfig.BackgroundColor)
            : _themeManager.BackgroundColor;

        var page = new TabPage(tabName)
        {
            BackColor = tabBackground,
            ForeColor = _themeManager.TextColor,
            Padding = new Padding(10),
        };

        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            WrapContents = true,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = tabBackground,
            Padding = new Padding(_profile.Layout.ButtonPadding),
        };

        if (buttons.Count == 0)
        {
            panel.Controls.Add(new Label
            {
                AutoSize = true,
                ForeColor = _themeManager.TextColor,
                Text = "No commands in this tab.",
                Padding = new Padding(8),
            });
        }
        else
        {
            foreach (var buttonDefinition in buttons)
            {
                panel.Controls.Add(BuildButton(buttonDefinition));
            }
        }

        page.Controls.Add(panel);
        return page;
    }

    private Button BuildButton(CommandButtonDefinition definition)
    {
        var baseColor = _themeManager.ResolveButtonBackColor(definition.Config);
        var button = new Button
        {
            Text = definition.Config.DisplayName,
            Width = _profile.Layout.ButtonWidth,
            Height = _profile.Layout.ButtonHeight,
            Margin = new Padding(_profile.Layout.ButtonPadding),
            FlatStyle = FlatStyle.Flat,
            BackColor = baseColor,
            ForeColor = _themeManager.TextColor,
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter,
            AutoEllipsis = true,
            Tag = definition,
        };

        button.FlatAppearance.BorderSize = _profile.Appearance.ButtonBorderSize;
        button.FlatAppearance.BorderColor = _themeManager.ButtonBorderColor;

        button.MouseEnter += (_, _) => button.BackColor = _themeManager.ButtonHoverColor;
        button.MouseLeave += (_, _) => button.BackColor = baseColor;
        button.Click += (_, _) => OnCommandButtonClicked(definition.Phrase, definition.CommandName);
        button.ContextMenuStrip = BuildButtonContextMenu(definition);

        return button;
    }

    private ContextMenuStrip BuildButtonContextMenu(CommandButtonDefinition definition)
    {
        var menu = new ContextMenuStrip
        {
            BackColor = _themeManager.BackgroundColor,
            ForeColor = _themeManager.TextColor,
        };

        var editItem = new ToolStripMenuItem("Edit...");
        editItem.Click += (_, _) => EditButton(definition);

        var revertItem = new ToolStripMenuItem("Revert to Default")
        {
            Enabled = definition.IsCustomized,
        };
        revertItem.Click += (_, _) => RevertButton(definition.Phrase);

        menu.Items.Add(editItem);
        menu.Items.Add(revertItem);
        return menu;
    }

    private void EditButton(CommandButtonDefinition definition)
    {
        try
        {
            using var dialog = new CommandLauncherEditDialog(definition.Phrase, definition.CommandName, definition.Config, _themeManager.ButtonBackgroundColor);
            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            _buttonRegistry.UpdateButtonMetadata(definition.Phrase, dialog.EditedConfig);
            UpdateStatus($"Saved customization for {definition.Phrase}");
            RefreshCommands();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to edit the button: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RevertButton(string phrase)
    {
        _buttonRegistry.RemoveCustomization(phrase);
        UpdateStatus($"Reverted {phrase} to the default layout.");
        RefreshCommands();
    }

    private void ResetCustomizations()
    {
        var result = MessageBox.Show(
            this,
            "Reset custom button overrides only? Theme and layout settings will remain intact.",
            "Reset UI Settings",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes)
            return;

        try
        {
            _profileManager.ResetToDefault();
            UpdateStatus("Custom button overrides cleared.");
            RefreshCommands();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to reset the profile: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void OnCommandButtonClicked(string phrase, string commandName)
    {
        try
        {
            WriteCommandInputFile(phrase);
            _dispatchDelegate?.Invoke(phrase);
            UpdateStatus($"Dispatched {phrase} ({commandName})");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Error dispatching command: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void WriteCommandInputFile(string phrase)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_inputFilePath)!);
            File.WriteAllText(_inputFilePath, phrase.Trim());
        }
        catch (IOException)
        {
            System.Threading.Thread.Sleep(100);
            File.WriteAllText(_inputFilePath, phrase.Trim());
        }
    }

    private void UpdateStatus(string message)
    {
        _statusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
    }

    /// <summary>
    /// Reload the profile and command list from disk.
    /// </summary>
    public void RefreshCommands()
    {
        LoadLauncherState();
        ApplyTheme();
        RenderButtons();

        if (_commands.Count == 0)
            UpdateStatus("No commands were found in commands.txt or animations/*.txt.");
        else
            UpdateStatus($"Loaded {_commands.Count} commands.");
    }
}
