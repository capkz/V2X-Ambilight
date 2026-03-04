using Microsoft.Win32;

namespace V2XAmbilight;

/// <summary>
/// System tray application: manages the capture timer, device connection loop,
/// and all user-facing settings via the NotifyIcon context menu.
/// </summary>
public sealed class TrayApp : ApplicationContext
{
    // -------------------------------------------------------------------------
    // State
    // -------------------------------------------------------------------------

    private Settings _settings = Settings.Load();
    private readonly KatanaDevice _device;
    private readonly ScreenSampler _sampler;
    private readonly CancellationTokenSource _cts = new();

    private enum Status { Disconnected, Connecting, Connected }
    private Status _status = Status.Disconnected;

    // -------------------------------------------------------------------------
    // UI controls
    // -------------------------------------------------------------------------

    private readonly NotifyIcon _tray;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly ToolStripMenuItem _enabledItem;
    private readonly ToolStripMenuItem _monitorMenu;
    private readonly ToolStripMenuItem _stripMenu;
    private readonly ToolStripMenuItem _startupItem;

    // -------------------------------------------------------------------------
    // Constructor
    // -------------------------------------------------------------------------

    public TrayApp()
    {
        _device = new KatanaDevice(_ => { }); // status conveyed via tray icon
        _sampler = new ScreenSampler();

        // --- Tray icon ---
        _tray = new NotifyIcon
        {
            Icon    = MakeIcon(Color.Gray),
            Visible = true,
            Text    = "V2X Ambilight — starting…",
        };

        // --- Menu items ---
        _enabledItem = new ToolStripMenuItem("Enabled", null, OnToggleEnabled)
            { Checked = _settings.Enabled, CheckOnClick = true };

        _monitorMenu = new ToolStripMenuItem("Monitor");
        _stripMenu   = new ToolStripMenuItem("Strip Height");

        _startupItem = new ToolStripMenuItem("Start with Windows", null, OnToggleStartup)
            { Checked = _settings.StartWithWindows, CheckOnClick = true };

        BuildMonitorMenu();
        BuildStripMenu();

        _tray.ContextMenuStrip = new ContextMenuStrip();
        _tray.ContextMenuStrip.Items.AddRange([
            new ToolStripMenuItem("Katana V2X Ambilight") { Enabled = false },
            new ToolStripSeparator(),
            _enabledItem,
            new ToolStripSeparator(),
            _monitorMenu,
            _stripMenu,
            new ToolStripSeparator(),
            _startupItem,
            new ToolStripSeparator(),
            new ToolStripMenuItem("Exit", null, OnExit),
        ]);

        // Rebuild monitor list each time menu opens (handles plug/unplug)
        _tray.ContextMenuStrip.Opening += (_, _) => BuildMonitorMenu();

        // --- Capture timer (UI thread) ---
        _timer = new System.Windows.Forms.Timer
        {
            Interval = 1000 / Math.Clamp(_settings.FrameRate, 1, 60),
        };
        _timer.Tick += OnTick;

        // Start device connection loop in background
        _ = ConnectLoopAsync(_cts.Token);
    }

    // -------------------------------------------------------------------------
    // Capture tick
    // -------------------------------------------------------------------------

    private void OnTick(object? sender, EventArgs e)
    {
        if (!_settings.Enabled) return;

        var screens = Screen.AllScreens;
        int idx = Math.Clamp(_settings.MonitorIndex, 0, screens.Length - 1);

        try
        {
            byte[] colors = _sampler.Sample(screens[idx], _settings.StripPercent);
            _device.SetColors(colors);
        }
        catch
        {
            // Device disconnected mid-session — stop timer, let connect loop retry
            _device.Disconnect();
            _timer.Stop();
            SetStatus(Status.Disconnected);
        }
    }

    // -------------------------------------------------------------------------
    // Device connection loop (background task)
    // -------------------------------------------------------------------------

    private async Task ConnectLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            if (_device.IsConnected)
            {
                await Task.Delay(1000, ct).ConfigureAwait(false);
                continue;
            }

            SetStatus(Status.Connecting);
            try
            {
                await _device.ConnectAsync(ct).ConfigureAwait(false);
                SetStatus(Status.Connected);

                // Start timer on UI thread
                _tray.ContextMenuStrip!.Invoke(_timer.Start);
            }
            catch (OperationCanceledException) { break; }
            catch
            {
                SetStatus(Status.Disconnected);
                await Task.Delay(5000, ct).ConfigureAwait(false); // retry every 5 s
            }
        }
    }

    // -------------------------------------------------------------------------
    // Status / icon
    // -------------------------------------------------------------------------

    private void SetStatus(Status status)
    {
        if (_status == status) return;
        _status = status;

        // Icon updates must be on the UI thread
        if (_tray.ContextMenuStrip?.InvokeRequired == true)
            _tray.ContextMenuStrip.Invoke(() => ApplyStatus(status));
        else
            ApplyStatus(status);
    }

    private void ApplyStatus(Status status)
    {
        var (color, tip) = status switch
        {
            Status.Connected    => (Color.Lime,   "V2X Ambilight — Connected"),
            Status.Connecting   => (Color.Orange,  "V2X Ambilight — Connecting…"),
            _                   => (Color.Gray,    "V2X Ambilight — Device not found"),
        };

        var old = _tray.Icon;
        _tray.Icon = MakeIcon(color);
        _tray.Text = tip;
        old?.Dispose();
    }

    private static Icon MakeIcon(Color color)
    {
        using var bmp = new Bitmap(16, 16);
        using var g   = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using var brush = new SolidBrush(color);
        g.FillEllipse(brush, 1, 1, 13, 13);
        return Icon.FromHandle(bmp.GetHicon());
    }

    // -------------------------------------------------------------------------
    // Menu handlers
    // -------------------------------------------------------------------------

    private void BuildMonitorMenu()
    {
        _monitorMenu.DropDownItems.Clear();
        var screens = Screen.AllScreens;
        for (int i = 0; i < screens.Length; i++)
        {
            int idx = i;
            var s   = screens[i];
            string label = s.Primary
                ? $"Primary ({s.Bounds.Width}×{s.Bounds.Height})"
                : $"Monitor {i + 1} ({s.Bounds.Width}×{s.Bounds.Height})";

            _monitorMenu.DropDownItems.Add(
                new ToolStripMenuItem(label, null, (_, _) => SelectMonitor(idx))
                    { Checked = _settings.MonitorIndex == i });
        }
    }

    private void BuildStripMenu()
    {
        _stripMenu.DropDownItems.Clear();
        foreach (int pct in new[] { 3, 5, 10, 15 })
        {
            int p = pct;
            _stripMenu.DropDownItems.Add(
                new ToolStripMenuItem($"{pct}%", null, (_, _) => SelectStrip(p))
                    { Checked = _settings.StripPercent == pct });
        }
    }

    private void SelectMonitor(int idx)
    {
        _settings.MonitorIndex = idx;
        _settings.Save();
        BuildMonitorMenu();
    }

    private void SelectStrip(int pct)
    {
        _settings.StripPercent = pct;
        _settings.Save();
        BuildStripMenu();
    }

    private void OnToggleEnabled(object? sender, EventArgs e)
    {
        _settings.Enabled = _enabledItem.Checked;
        _settings.Save();

        if (!_settings.Enabled)
            _device.SetColors(new byte[21]); // turn LEDs off
    }

    private void OnToggleStartup(object? sender, EventArgs e)
    {
        _settings.StartWithWindows = _startupItem.Checked;
        _settings.Save();
        ApplyAutoStart(_settings.StartWithWindows);
    }

    private static void ApplyAutoStart(bool enable)
    {
        using var key = Registry.CurrentUser.OpenSubKey(
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", writable: true)!;
        if (enable)
            key.SetValue("V2XAmbilight", $"\"{Application.ExecutablePath}\"");
        else
            key.DeleteValue("V2XAmbilight", throwOnMissingValue: false);
    }

    private void OnExit(object? sender, EventArgs e)
    {
        _cts.Cancel();
        _timer.Stop();

        try { _device.SetColors(new byte[21]); } catch { } // LEDs off
        _tray.Visible = false;
        Application.Exit();
    }

    // -------------------------------------------------------------------------
    // Cleanup
    // -------------------------------------------------------------------------

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _cts.Cancel();
            _cts.Dispose();
            _timer.Dispose();
            _device.Dispose();
            _sampler.Dispose();
            _tray.Dispose();
        }
        base.Dispose(disposing);
    }
}
