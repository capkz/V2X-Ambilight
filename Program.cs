namespace V2XAmbilight;

static class Program
{
    [STAThread]
    static void Main()
    {
        // Single instance guard
        using var mutex = new Mutex(true, "V2XAmbilight_SingleInstance", out bool first);
        if (!first) return;

        ApplicationConfiguration.Initialize();
        Application.Run(new TrayApp());
    }
}
