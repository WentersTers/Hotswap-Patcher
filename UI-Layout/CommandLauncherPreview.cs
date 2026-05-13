using System;
using System.Windows.Forms;

namespace PAIcomPatcher.UILayout;

/// <summary>
/// Standalone test/preview form for the command launcher.
/// This can be opened independently to visually verify the button grid layout
/// and test command dispatching without needing the full injected runtime.
/// 
/// Usage:
///   CommandLauncherPreview.ShowPreview();
/// 
/// This form reuses the exact same CommandLauncher UI code as the injected runtime version,
/// so the visual layout can be verified before shipping.
/// </summary>
public class CommandLauncherPreview
{
    /// <summary>
    /// Opens the preview form in a standalone window for visual inspection.
    /// </summary>
    public static void ShowPreview(string? baseDir = null)
    {
        try
        {
            // Create a dispatch delegate that logs to the console and command_input.txt
            Action<string> testDispatch = (phrase) =>
            {
                MessageBox.Show(
                    $"Test dispatch would execute: {phrase}\n\n(In production, this would invoke the injected runtime's DispatchCommandPhrase method)",
                    "Preview: Command Dispatch",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            };

            // Create the launcher form (reuses the same UI code)
            var launcher = new CommandLauncher(baseDir, testDispatch);
            launcher.Text = "Command Launcher - PREVIEW MODE";

            // Show as a modal dialog
            launcher.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error loading preview form:\n\n{ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
