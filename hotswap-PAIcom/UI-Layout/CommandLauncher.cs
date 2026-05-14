using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PAIcomPatcher.UILayout;

/// <summary>
/// WinForms-based command launcher that creates a button grid from custom-commands/commands.txt.
/// If commands.txt is not found, falls back to scanning animations/ directory for .txt files.
/// Each button, when clicked, writes the command phrase to command_input.txt and dispatches it.
/// </summary>
public class CommandLauncher : Form
{
    private readonly string _baseDir;
    private readonly string _commandsFilePath;
    private readonly string _inputFilePath;
    private readonly string _animationsDir;
    private FlowLayoutPanel? _flowPanel;
    private Dictionary<string, string> _commands; // phrase -> command name mapping
    private Action<string>? _dispatchDelegate;
    
    // Constants for UI spacing  
    private const int ButtonWidth = 120;
    private const int ButtonHeight = 50;
    private const int ButtonPadding = 5;
    private const int ColumnsPerRow = 5;

    public CommandLauncher(string? baseDir = null, Action<string>? dispatchDelegate = null)
    {
        _baseDir = baseDir ?? GetDefaultBaseDir();
        _commandsFilePath = Path.Combine(_baseDir, "custom-commands", "commands.txt");
        _inputFilePath = Path.Combine(_baseDir, "command_input.txt");
        _animationsDir = Path.Combine(_baseDir, "animations");
        _dispatchDelegate = dispatchDelegate;
        _commands = new Dictionary<string, string>();

        InitializeComponent();
        LoadCommands();
        CreateButtonGrid();
    }

    private string GetDefaultBaseDir()
    {
        // Try to use the directory of the executing assembly
        var location = System.Reflection.Assembly.GetExecutingAssembly().Location;
        var dir = Path.GetDirectoryName(location) ?? AppDomain.CurrentDomain.BaseDirectory;

        // If we're in bin/Release/net8.0-windows (or similar), walk up to find the project root
        // Look for custom-commands/commands.txt by walking up the directory tree
        var searchDir = new DirectoryInfo(dir);
        for (int i = 0; i < 5; i++)  // Search up to 5 levels
        {
            if (searchDir == null) break;
            
            var commandsPath = Path.Combine(searchDir.FullName, "custom-commands", "commands.txt");
            if (File.Exists(commandsPath))
                return searchDir.FullName;
            
            searchDir = searchDir.Parent;
        }

        // Fallback: return the assembly directory
        return dir;
    }

    private void InitializeComponent()
    {
        this.Text = "PAIcom Command Launcher";
        this.Size = new Size(850, 600);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.Font = new Font("Segoe UI", 9F);
        this.BackColor = Color.FromArgb(45, 45, 48); // Dark theme
        this.ForeColor = Color.White;

        // Create a scroll container
        _flowPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            AutoSize = false,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(ButtonPadding),
            BackColor = Color.FromArgb(45, 45, 48),
        };

        this.Controls.Add(_flowPanel);

        // Add a status label at the bottom
        var statusLabel = new Label
        {
            Dock = DockStyle.Bottom,
            Text = "Ready. Commands will be dispatched immediately.",
            Height = 30,
            BackColor = Color.FromArgb(60, 60, 65),
            ForeColor = Color.LightGray,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(10, 0, 0, 0),
        };
        this.Controls.Add(statusLabel);
    }

    private void LoadCommands()
    {
        _commands.Clear();

        // Try loading from commands.txt first
        if (File.Exists(_commandsFilePath))
        {
            try
            {
                foreach (var line in File.ReadAllLines(_commandsFilePath))
                {
                    var trimmed = line.Trim();
                    if (trimmed.Length == 0 || trimmed.StartsWith("#"))
                        continue;

                    // Parse format: "hey paicom do something (file.txt)"
                    var parenOpen = trimmed.LastIndexOf('(');
                    var parenClose = trimmed.LastIndexOf(')');

                    string phrase;
                    string commandName;

                    if (parenOpen > 0 && parenClose > parenOpen)
                    {
                        phrase = trimmed.Substring(0, parenOpen).Trim();
                        var fileRef = trimmed.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();
                        
                        // Strip .txt extension
                        if (fileRef.ToLowerInvariant().EndsWith(".txt"))
                            fileRef = fileRef.Substring(0, fileRef.Length - 4);
                        
                        commandName = fileRef;
                    }
                    else
                    {
                        phrase = trimmed;
                        commandName = trimmed;
                    }

                    if (!_commands.ContainsKey(phrase))
                        _commands[phrase] = commandName;
                }
                return; // Successfully loaded from commands.txt, don't use fallback
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading commands: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // Fallback: scan animations/ directory for .txt files
        LoadCommandsFromAnimationsDirectory();
    }

    private void LoadCommandsFromAnimationsDirectory()
    {
        if (!Directory.Exists(_animationsDir))
        {
            MessageBox.Show($"No commands.txt found and animations directory not found: {_animationsDir}", 
                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var txtFiles = Directory.GetFiles(_animationsDir, "*.txt");
            foreach (var filePath in txtFiles)
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                
                // Create a command phrase from the filename
                // Example: "youtube.txt" -> "hey paicom youtube" as phrase, "youtube" as command name
                var phrase = $"hey paicom {fileName}";
                var commandName = fileName;
                
                if (!_commands.ContainsKey(phrase))
                    _commands[phrase] = commandName;
            }

            if (_commands.Count > 0)
            {
                var label = this.Text;
                if (!label.Contains("(Animations)"))
                    this.Text += " (Animations Fallback)";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error scanning animations directory: {ex.Message}", 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void CreateButtonGrid()
    {
        if (_flowPanel == null) return;
        
        _flowPanel.Controls.Clear();

        if (_commands.Count == 0)
        {
            var label = new Label
            {
                Text = "No commands found (commands.txt or animations/*.txt)",
                AutoSize = true,
                ForeColor = Color.LightGray,
            };
            _flowPanel.Controls.Add(label);
            return;
        }

        int index = 0;
        foreach (var kvp in _commands)
        {
            var phrase = kvp.Key;
            var commandName = kvp.Value;

            var button = new Button
            {
                Text = ShortenPhrase(phrase, 15),
                Width = ButtonWidth,
                Height = ButtonHeight,
                Margin = new Padding(ButtonPadding),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(63, 63, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Tag = new Tuple<string, string>(phrase, commandName),
            };

            button.FlatAppearance.BorderColor = Color.FromArgb(109, 109, 109);
            button.FlatAppearance.BorderSize = 1;

            // Add hover effect
            button.MouseEnter += (s, e) =>
            {
                button.BackColor = Color.FromArgb(80, 80, 90);
            };
            button.MouseLeave += (s, e) =>
            {
                button.BackColor = Color.FromArgb(63, 63, 70);
            };

            button.Click += (s, e) => OnCommandButtonClicked(phrase, commandName);

            _flowPanel.Controls.Add(button);
            index++;
        }
    }

    private string ShortenPhrase(string phrase, int maxLength)
    {
        if (phrase.Length <= maxLength)
            return phrase;

        // Remove "hey paicom " prefix if present
        var stripped = phrase.StartsWith("hey paicom ", StringComparison.OrdinalIgnoreCase)
            ? phrase.Substring(11).Trim()
            : phrase;

        if (stripped.Length <= maxLength)
            return stripped;

        return stripped.Substring(0, maxLength - 3) + "…";
    }

    private void OnCommandButtonClicked(string phrase, string commandName)
    {
        try
        {
            // Write to command_input.txt for backward compatibility with the watcher
            WriteCommandInputFile(phrase);

            // If a dispatch delegate was registered, invoke it directly
            if (_dispatchDelegate != null)
            {
                _dispatchDelegate(phrase);
            }

            LogMessage($"Dispatched: {phrase} ({commandName})");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error dispatching command: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void WriteCommandInputFile(string phrase)
    {
        try
        {
            // Ensure the file exists
            if (!File.Exists(_inputFilePath))
            {
                File.WriteAllText(_inputFilePath, "");
            }

            // Write the phrase to trigger the watcher
            File.WriteAllText(_inputFilePath, phrase.Trim());
        }
        catch (IOException)
        {
            // Retry once if file is locked
            System.Threading.Thread.Sleep(100);
            File.WriteAllText(_inputFilePath, phrase.Trim());
        }
    }

    private void LogMessage(string message)
    {
        // Find the status label and update it
        foreach (Control ctrl in this.Controls)
        {
            if (ctrl is Label label && label.Dock == DockStyle.Bottom)
            {
                label.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
                break;
            }
        }
    }

    /// <summary>
    /// Reload commands from disk (useful if commands.txt was edited externally).
    /// </summary>
    public void RefreshCommands()
    {
        LoadCommands();
        CreateButtonGrid();
    }
}
