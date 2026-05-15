using PAIcomPatcher.UI.Configuration;
using PAIcomPatcher.UI.Styling;

namespace PAIcomPatcher.UILayout;

public sealed class CommandLauncherEditDialog : Form
{
    private readonly TextBox _displayNameBox;
    private readonly TextBox _groupTabBox;
    private readonly NumericUpDown _positionBox;
    private readonly CheckBox _visibleBox;
    private readonly TextBox _backgroundColorBox;
    private readonly TextBox _notesBox;
    private readonly Button _colorButton;
    private readonly ColorDialog _colorDialog;

    public ButtonConfig EditedConfig { get; private set; }

    public CommandLauncherEditDialog(string phrase, string commandName, ButtonConfig initialConfig, Color fallbackButtonColor)
    {
        Text = $"Edit Button - {phrase}";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(460, 330);
        Font = new Font("Segoe UI", 9F);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 7,
            Padding = new Padding(12),
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));

        layout.Controls.Add(new Label { Text = "Display name", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        _displayNameBox = new TextBox { Text = initialConfig.DisplayName, Dock = DockStyle.Fill };
        layout.SetColumnSpan(_displayNameBox, 2);
        layout.Controls.Add(_displayNameBox, 1, 0);

        layout.Controls.Add(new Label { Text = "Group tab", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        _groupTabBox = new TextBox { Text = initialConfig.GroupTab, Dock = DockStyle.Fill };
        layout.SetColumnSpan(_groupTabBox, 2);
        layout.Controls.Add(_groupTabBox, 1, 1);

        layout.Controls.Add(new Label { Text = "Position", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
        _positionBox = new NumericUpDown { Minimum = 0, Maximum = 9999, Value = Math.Max(0, initialConfig.Position), Dock = DockStyle.Left, Width = 100 };
        layout.SetColumnSpan(_positionBox, 2);
        layout.Controls.Add(_positionBox, 1, 2);

        layout.Controls.Add(new Label { Text = "Visible", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
        _visibleBox = new CheckBox { Checked = initialConfig.Visible, Dock = DockStyle.Left, AutoSize = true };
        layout.SetColumnSpan(_visibleBox, 2);
        layout.Controls.Add(_visibleBox, 1, 3);

        layout.Controls.Add(new Label { Text = "Background", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 4);
        _backgroundColorBox = new TextBox { Text = initialConfig.BackgroundColor ?? string.Empty, Dock = DockStyle.Fill };
        layout.Controls.Add(_backgroundColorBox, 1, 4);
        _colorButton = new Button { Text = "Pick...", Dock = DockStyle.Fill };
        layout.Controls.Add(_colorButton, 2, 4);

        layout.Controls.Add(new Label { Text = "Notes", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 5);
        _notesBox = new TextBox { Text = initialConfig.Notes ?? string.Empty, Dock = DockStyle.Fill };
        layout.SetColumnSpan(_notesBox, 2);
        layout.Controls.Add(_notesBox, 1, 5);

        var buttonRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
        };
        var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
        var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };
        buttonRow.Controls.Add(okButton);
        buttonRow.Controls.Add(cancelButton);
        layout.Controls.Add(buttonRow, 0, 6);
        layout.SetColumnSpan(buttonRow, 3);

        Controls.Add(layout);

        AcceptButton = okButton;
        CancelButton = cancelButton;

        _colorDialog = new ColorDialog { FullOpen = true, Color = fallbackButtonColor };
        _colorButton.Click += (_, _) => PickColor();
        okButton.Click += (_, _) => Commit();

        EditedConfig = new ButtonConfig
        {
            DisplayName = initialConfig.DisplayName,
            GroupTab = initialConfig.GroupTab,
            Position = initialConfig.Position,
            Visible = initialConfig.Visible,
            BackgroundColor = initialConfig.BackgroundColor,
            Notes = initialConfig.Notes,
        };
    }

    private void PickColor()
    {
        if (_colorDialog.ShowDialog(this) != DialogResult.OK)
            return;

        _backgroundColorBox.Text = ColorUtilities.ToHex(_colorDialog.Color);
    }

    private void Commit()
    {
        if (string.IsNullOrWhiteSpace(_displayNameBox.Text))
        {
            MessageBox.Show(this, "Display name cannot be empty.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
            return;
        }

        EditedConfig = new ButtonConfig
        {
            DisplayName = _displayNameBox.Text.Trim(),
            GroupTab = string.IsNullOrWhiteSpace(_groupTabBox.Text) ? "General" : _groupTabBox.Text.Trim(),
            Position = (int)_positionBox.Value,
            Visible = _visibleBox.Checked,
            BackgroundColor = string.IsNullOrWhiteSpace(_backgroundColorBox.Text) ? null : _backgroundColorBox.Text.Trim(),
            Notes = string.IsNullOrWhiteSpace(_notesBox.Text) ? null : _notesBox.Text.Trim(),
        };
    }
}