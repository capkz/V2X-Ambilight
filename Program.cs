namespace V2XAmbilight;

static class Program
{
    static readonly string CrashLog = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "V2XAmbilight", "crash.log");

    [STAThread]
    static void Main(string[] args)
    {
        // Elevated kill helper — launched by the app itself via runas
        if (args.Length == 2 && args[0] == "--kill")
        {
            foreach (var p in System.Diagnostics.Process.GetProcessesByName(args[1]))
                try { p.Kill(); } catch { }
            return;
        }

        // Single instance guard
        using var mutex = new Mutex(true, "V2XAmbilight_SingleInstance", out bool first);
        if (!first) return;

        Application.ThreadException                        += (_, e) => LogCrash(e.Exception);
        AppDomain.CurrentDomain.UnhandledException         += (_, e) => LogCrash(e.ExceptionObject as Exception);
        TaskScheduler.UnobservedTaskException              += (_, e) => { LogCrash(e.Exception); e.SetObserved(); };

        ApplicationConfiguration.Initialize();
        Application.Run(new TrayApp());
    }

    static void LogCrash(Exception? ex)
    {
        if (ex is null) return;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(CrashLog)!);
            File.AppendAllText(CrashLog,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex}\n\n");

            MessageBox.Show(
                $"V2X Ambilight crashed:\n\n{ex.Message}\n\nSee: {CrashLog}",
                "V2X Ambilight — Crash", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch { }
    }
}
