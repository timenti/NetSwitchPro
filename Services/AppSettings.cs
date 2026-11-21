using System.Text.Json;

namespace NetworkAdapterSwitcher.Services;

internal sealed class AppSettings
{
    public string Theme { get; set; } = "dark";
    public string Language { get; set; } = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "ru" ? "Russian" : "English";
    public bool ShowVirtual { get; set; }
    public bool ShowBluetooth { get; set; }
    public bool ShowAdapters { get; set; }
    public Guid? AdapterAId { get; set; }
    public Guid? AdapterBId { get; set; }
    public string? AdapterA { get; set; }
    public string? AdapterB { get; set; }
    public Guid? WifiAdapterId { get; set; }
    public string? WifiA { get; set; }
    public string? WifiB { get; set; }
    public bool IsLight => Theme.Equals("light", StringComparison.OrdinalIgnoreCase);

    private static string Local => Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
    private static string FilePath => Path.Combine(Local, "NetSwitchPro", "settings.v2.json");

    public static AppSettings Load()
    {
        var candidates = new[]
        {
            Path.Combine(Local, "NetSwitchPro", "settings.json"),
            Path.Combine(Local, "Temp", "NetSwitchPro", "settings.json"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NetSwitchPro", "settings.json")
        };
        foreach (var path in candidates.Where(File.Exists).OrderByDescending(File.GetLastWriteTimeUtc).Prepend(FilePath))
        {
            try
            {
                if (!File.Exists(path)) continue;
                if (JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path)) is { } found)
                {
                    found.Theme ??= "dark";
                    found.Language ??= "English";
                    if (path != FilePath)
                        found.ShowAdapters = found.AdapterA is not null && found.AdapterB is not null;
                    return found;
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            catch (JsonException) { }
        }
        return new();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        string temporary = FilePath + "." + Guid.NewGuid().ToString("N");
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temporary, FilePath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
