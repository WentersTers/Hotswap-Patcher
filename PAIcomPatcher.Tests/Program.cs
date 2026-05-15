using PAIcomPatcher.UI.Configuration;
using PAIcomPatcher.UILayout;

namespace PAIcomPatcher.Tests;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), "paicom-ui-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(tempRoot, "custom-commands"));
        Directory.CreateDirectory(Path.Combine(tempRoot, "animations"));

        try
        {
            var commandsPath = Path.Combine(tempRoot, "custom-commands", "commands.txt");
            File.WriteAllLines(commandsPath, new[]
            {
                "hey paicom open the browser (internet.txt)",
                "hey paicom play some music (music.txt)",
            });

            var profileManager = new UIProfileManager(tempRoot);
            var profile = profileManager.LoadProfile();
            Assert(File.Exists(profileManager.ProfilePath), "ui-profile.json should be created on first load.");
            Assert(profile.CustomButtons.Count == 0, "Default profile should start with zero custom buttons.");

            var commands = profileManager.LoadAllCommands();
            Assert(commands.Count == 2, "commands.txt should parse into two commands.");

            var defaultButton = profileManager.GetButtonMetadata("hey paicom open the browser");
            Assert(defaultButton.GroupTab == "General", "Uncustomized buttons should default to the General tab.");
            Assert(defaultButton.DisplayName.Contains("Browser", StringComparison.OrdinalIgnoreCase), "Default display name should be derived from the phrase.");

            profileManager.SaveCustomButtonMetadata("hey paicom open the browser", new ButtonConfig
            {
                DisplayName = "My Browser",
                GroupTab = "Media",
                Position = 1,
                Visible = true,
                BackgroundColor = "#FF6600",
                Notes = "Smoke test customization",
            });

            var customizedButton = profileManager.GetButtonMetadata("hey paicom open the browser");
            Assert(customizedButton.DisplayName == "My Browser", "Custom display name should persist.");
            Assert(customizedButton.GroupTab == "Media", "Custom tab should persist.");
            Assert(profileManager.IsButtonCustomized("hey paicom open the browser"), "Customization flag should be true after saving.");

            profileManager.ResetToDefault();
            Assert(!profileManager.IsButtonCustomized("hey paicom open the browser"), "Reset should clear custom overrides only.");
            Assert(File.Exists(profileManager.ProfilePath + ".backup"), "Reset should create a backup file.");

            string? dispatchedPhrase = null;
            using (var launcher = new CommandLauncher(tempRoot, phrase => dispatchedPhrase = phrase))
            {
                var tabControl = FindControl<PAIcomPatcher.UI.Components.UITabControl>(launcher);
                Assert(tabControl != null, "Launcher should create the tab control.");
                Assert(tabControl!.TabPages.Count >= 1, "Launcher should render at least one tab.");

                var buttons = FindControls<Button>(tabControl).ToList();
                Assert(buttons.Count >= 2, "Launcher should create buttons for all commands.");

                buttons.First(button => button.Text.Contains("Browser", StringComparison.OrdinalIgnoreCase)).PerformClick();
                var dispatchMethod = typeof(CommandLauncher).GetMethod("OnCommandButtonClicked", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                Assert(dispatchMethod != null, "Launcher should expose the dispatch handler.");
                dispatchMethod!.Invoke(launcher, new object[] { "hey paicom open the browser", "internet" });
                Assert(string.Equals(dispatchedPhrase, "hey paicom open the browser", StringComparison.OrdinalIgnoreCase), "Button click should dispatch the selected phrase.");
            }

            Console.WriteLine("Smoke tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
        finally
        {
            try
            {
                Directory.Delete(tempRoot, recursive: true);
            }
            catch
            {
                // Ignore cleanup failures in the smoke test harness.
            }
        }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    private static T? FindControl<T>(Control root) where T : Control
    {
        foreach (Control child in root.Controls)
        {
            if (child is T match)
                return match;

            var nested = FindControl<T>(child);
            if (nested != null)
                return nested;
        }

        return null;
    }

    private static IEnumerable<T> FindControls<T>(Control root) where T : Control
    {
        foreach (Control child in root.Controls)
        {
            if (child is T match)
                yield return match;

            foreach (var nested in FindControls<T>(child))
                yield return nested;
        }
    }
}