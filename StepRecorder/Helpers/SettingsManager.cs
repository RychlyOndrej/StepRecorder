using System.IO;
using System.Text.Json;
using StepRecorder.Models;

namespace StepRecorder.Helpers;

public static class SettingsManager
{
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "StepRecorder", "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented             = true,
        PropertyNameCaseInsensitive = true,
        Converters                = { }
    };

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
            }
        }
        catch { /* return defaults on any error */ }

        return new AppSettings();
    }

    public static void Save(AppSettings settings)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, JsonOptions));
        }
        catch { /* silently ignore save errors */ }
    }
}
