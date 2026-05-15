namespace PAIcomPatcher.UI.Configuration;

public static class ProfileValidator
{
    public static bool Validate(UIProfile profile, out List<string> errors)
    {
        errors = new List<string>();

        if (profile is null)
        {
            errors.Add("Profile is null.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(profile.Version))
            errors.Add("Profile version is missing.");

        if (profile.Layout.WindowWidth <= 0)
            errors.Add("Window width must be greater than zero.");

        if (profile.Layout.WindowHeight <= 0)
            errors.Add("Window height must be greater than zero.");

        if (profile.Layout.ButtonWidth <= 0)
            errors.Add("Button width must be greater than zero.");

        if (profile.Layout.ButtonHeight <= 0)
            errors.Add("Button height must be greater than zero.");

        if (profile.Layout.ButtonPadding < 0)
            errors.Add("Button padding cannot be negative.");

        return errors.Count == 0;
    }
}