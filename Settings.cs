using System.Text.Json;

namespace V2XAmbilight;

public sealed class Settings
{
    public int   MonitorIndex      { get; set; } = 0;
    public int   StripPercent      { get; set; } = 5;
    public int   FrameRate         { get; set; } = 20;
    public bool  Enabled           { get; set; } = true;
    public bool  StartWithWindows  { get; set; } = false;
    public float Brightness        { get; set; } = 1.0f;
    public float Saturation        { get; set; } = 1.5f;
    public float Smoothing         { get; set; } = 0.0f;
    public bool  SuppressConflictWarning { get; set; } = false;

    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "V2XAmbilight", "settings.json");

    public static Settings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new();
        }
        catch { }
        return new();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(this,
            new JsonSerializerOptions { WriteIndented = true }));
    }
}
