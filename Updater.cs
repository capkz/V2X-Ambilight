using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;

namespace V2XAmbilight;

internal static class Updater
{
    const string ApiUrl = "https://api.github.com/repos/capkz/V2X-Ambilight/releases/latest";

    public static string CurrentVersion =>
        (Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "0.0.0")
        .Split('+')[0]; // strip commit hash suffix the SDK appends

    public static async Task<(bool HasUpdate, string Tag, string DownloadUrl)> CheckAsync()
    {
        try
        {
            using var http = MakeClient();
            string json = await http.GetStringAsync(ApiUrl).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            string tag = root.GetProperty("tag_name").GetString() ?? "";
            string url = root.GetProperty("assets")[0]
                             .GetProperty("browser_download_url").GetString() ?? "";

            return (IsNewer(tag, CurrentVersion), tag, url);
        }
        catch { return (false, "", ""); }
    }

    static readonly string InstallPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "V2XAmbilight", "V2XAmbilight.exe");

    public static async Task ApplyAsync(string downloadUrl, Action<string> log)
    {
        string tmp = Path.Combine(Path.GetTempPath(), "V2XAmbilight_new.exe");
        // Always update the install location; if running from elsewhere update that too
        string current = File.Exists(InstallPath) ? InstallPath : Application.ExecutablePath;
        string bat     = Path.Combine(Path.GetTempPath(), "V2XAmbilight_update.bat");

        log("Downloading update…");
        using var http = MakeClient();
        byte[] data = await http.GetByteArrayAsync(downloadUrl).ConfigureAwait(false);
        await File.WriteAllBytesAsync(tmp, data).ConfigureAwait(false);

        log("Download complete — restarting…");

        await File.WriteAllTextAsync(bat, $"""
            @echo off
            timeout /t 2 /nobreak >nul
            copy /y "{tmp}" "{current}"
            start "" "{current}"
            del "{tmp}"
            del "%~f0"
            """).ConfigureAwait(false);

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName        = bat,
            UseShellExecute = true,
            WindowStyle     = System.Diagnostics.ProcessWindowStyle.Hidden,
        });

        Application.Exit();
    }

    static bool IsNewer(string remoteTag, string current)
    {
        string remote = remoteTag.TrimStart('v');
        return Version.TryParse(remote, out var r)
            && Version.TryParse(current,  out var c)
            && r > c;
    }

    static HttpClient MakeClient()
    {
        var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.Add(
            new ProductInfoHeaderValue("V2XAmbilight", CurrentVersion));
        return http;
    }
}
