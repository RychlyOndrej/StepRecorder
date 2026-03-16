using System.Globalization;
using System.Reflection;
using System.Resources;

namespace StepRecorder.Helpers;

internal static class L10n
{
    private static readonly ResourceManager ResourceManager =
        new("StepRecorder.Resources.Strings", Assembly.GetExecutingAssembly());

    public static void ApplyLanguage(string languageCode)
    {
        var culture = languageCode?.Trim().ToLowerInvariant() switch
        {
            "en" => new CultureInfo("en"),
            _ => new CultureInfo("cs")
        };

        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
    }

    public static string T(string key)
    {
        var value = ResourceManager.GetString(key, CultureInfo.CurrentUICulture);
        return string.IsNullOrWhiteSpace(value) ? key : value;
    }
}