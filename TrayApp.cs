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
    private readonly LogWindow _log;
    private byte[] _lastColors = new byte[21];

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
    private readonly ToolStripMenuItem _brightnessMenu;
    private readonly ToolStripMenuItem _saturationMenu;
    private readonly ToolStripMenuItem _smoothingMenu;
    private readonly ToolStripMenuItem _startupItem;
    private readonly ToolStripMenuItem _updateItem;

    // -------------------------------------------------------------------------
    // Constructor
    // -------------------------------------------------------------------------

    public TrayApp()
    {
        _log    = new LogWindow();
        _device = new KatanaDevice(_log.Append);
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

        _updateItem     = new ToolStripMenuItem("Update Available!", null, OnUpdateClick)
            { Visible = false };
        _brightnessMenu = new ToolStripMenuItem("Brightness");
        _saturationMenu = new ToolStripMenuItem("Vibrancy");
        _smoothingMenu  = new ToolStripMenuItem("Smoothing");

        BuildMonitorMenu();
        BuildStripMenu();
        BuildBrightnessMenu();
        BuildSaturationMenu();
        BuildSmoothingMenu();

        _tray.ContextMenuStrip = new ContextMenuStrip();
        _tray.ContextMenuStrip.Items.AddRange([
            new ToolStripMenuItem($"Katana V2X Ambilight  v{Updater.CurrentVersion}") { Enabled = false },
            new ToolStripSeparator(),
            _updateItem,
            _enabledItem,
            new ToolStripSeparator(),
            _monitorMenu,
            _stripMenu,
            _brightnessMenu,
            _saturationMenu,
            _smoothingMenu,
            new ToolStripSeparator(),
            _startupItem,
            new ToolStripMenuItem("Check for Updates", null, OnCheckUpdate),
            new ToolStripMenuItem("Show Log",          null, (_, _) => _log.Show()),
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

        // Check for updates shortly after startup
        _ = CheckForUpdateAsync();
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
            ProcessColors(colors, _settings.Brightness, _settings.Saturation);
            ApplySmoothing(colors, _lastColors, _settings.Smoothing);
            _device.SetColors(colors);
        }
        catch (Exception ex)
        {
            // Device disconnected mid-session — stop timer, let connect loop retry
            _log.Append($"Device error: {ex.Message}");
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
                _log.Invoke(_timer.Start);
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                bool denied = ex.Message.Contains("denied", StringComparison.OrdinalIgnoreCase)
                           || ex.Message.Contains("Access", StringComparison.OrdinalIgnoreCase);

                string? blocker = denied ? FindPortBlocker() : null;

                if (blocker != null)
                {
                    _log.Append($"Port blocked by '{blocker}' — close it, then the app will reconnect automatically.");
                    _log.BeginInvoke(() =>
                        _tray.ShowBalloonTip(6000, "V2X Ambilight",
                            $"Close '{blocker}' to allow V2X Ambilight to connect.", ToolTipIcon.Warning));
                }
                else
                {
                    _log.Append($"Connect failed: {ex.Message} — retrying in 5 s");
                }

                SetStatus(Status.Disconnected);
                await Task.Delay(5000, ct).ConfigureAwait(false);
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
        if (_log.InvokeRequired)
            _log.BeginInvoke(() => ApplyStatus(status));
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
        int[] presets = [3, 5, 10, 15];
        foreach (int pct in presets)
        {
            int p = pct;
            _stripMenu.DropDownItems.Add(
                new ToolStripMenuItem($"{pct}%", null, (_, _) => SelectStrip(p))
                    { Checked = _settings.StripPercent == pct });
        }
        // Show custom value if it doesn't match a preset
        bool isCustom = !Array.Exists(presets, p => p == _settings.StripPercent);
        _stripMenu.DropDownItems.Add(new ToolStripSeparator());
        _stripMenu.DropDownItems.Add(
            new ToolStripMenuItem(isCustom ? $"Custom ({_settings.StripPercent}%)…" : "Custom…",
                null, (_, _) =>
                {
                    if (PromptPercent("Strip Height (1–50%)", _settings.StripPercent, 1, 50, out int val))
                        SelectStrip(val);
                }) { Checked = isCustom });
    }

    private void BuildBrightnessMenu()
    {
        _brightnessMenu.DropDownItems.Clear();
        float[] presets = [0.5f, 0.75f, 1.0f, 1.25f, 1.5f, 2.0f];
        foreach (float v in presets)
        {
            float val = v;
            string label = v == 1.0f ? "100% (default)" : $"{(int)(v * 100)}%";
            _brightnessMenu.DropDownItems.Add(
                new ToolStripMenuItem(label, null, (_, _) => SelectBrightness(val))
                    { Checked = Math.Abs(_settings.Brightness - v) < 0.01f });
        }
        bool isCustom = !Array.Exists(presets, p => Math.Abs(p - _settings.Brightness) < 0.01f);
        _brightnessMenu.DropDownItems.Add(new ToolStripSeparator());
        _brightnessMenu.DropDownItems.Add(
            new ToolStripMenuItem(isCustom ? $"Custom ({(int)(_settings.Brightness * 100)}%)…" : "Custom…",
                null, (_, _) =>
                {
                    if (PromptPercent("Brightness (10–500%)", (int)(_settings.Brightness * 100), 10, 500, out int val))
                        SelectBrightness(val / 100f);
                }) { Checked = isCustom });
    }

    private void BuildSaturationMenu()
    {
        _saturationMenu.DropDownItems.Clear();
        float[] presets = [0.5f, 1.0f, 1.5f, 2.0f, 3.0f];
        foreach (float v in presets)
        {
            float val = v;
            string label = v == 1.0f ? "100% (default)" : $"{(int)(v * 100)}%";
            _saturationMenu.DropDownItems.Add(
                new ToolStripMenuItem(label, null, (_, _) => SelectSaturation(val))
                    { Checked = Math.Abs(_settings.Saturation - v) < 0.01f });
        }
        bool isCustom = !Array.Exists(presets, p => Math.Abs(p - _settings.Saturation) < 0.01f);
        _saturationMenu.DropDownItems.Add(new ToolStripSeparator());
        _saturationMenu.DropDownItems.Add(
            new ToolStripMenuItem(isCustom ? $"Custom ({(int)(_settings.Saturation * 100)}%)…" : "Custom…",
                null, (_, _) =>
                {
                    if (PromptPercent("Vibrancy (10–500%)", (int)(_settings.Saturation * 100), 10, 500, out int val))
                        SelectSaturation(val / 100f);
                }) { Checked = isCustom });
    }

    private void BuildSmoothingMenu()
    {
        _smoothingMenu.DropDownItems.Clear();
        float[] presets = [0f, 0.05f, 0.1f, 0.2f, 0.3f];
        foreach (float v in presets)
        {
            float val = v;
            string label = v == 0f ? "Off (default)" : $"{(int)(v * 100)}%";
            _smoothingMenu.DropDownItems.Add(
                new ToolStripMenuItem(label, null, (_, _) => SelectSmoothing(val))
                    { Checked = Math.Abs(_settings.Smoothing - v) < 0.01f });
        }
        bool isCustom = !Array.Exists(presets, p => Math.Abs(p - _settings.Smoothing) < 0.01f);
        _smoothingMenu.DropDownItems.Add(new ToolStripSeparator());
        _smoothingMenu.DropDownItems.Add(
            new ToolStripMenuItem(isCustom ? $"Custom ({(int)(_settings.Smoothing * 100)}%)…" : "Custom…",
                null, (_, _) =>
                {
                    if (PromptPercent("Smoothing (0–99%)", (int)(_settings.Smoothing * 100), 0, 99, out int val))
                        SelectSmoothing(val / 100f);
                }) { Checked = isCustom });
    }

    private void SelectBrightness(float v)
    {
        _settings.Brightness = v;
        _settings.Save();
        BuildBrightnessMenu();
    }

    private void SelectSaturation(float v)
    {
        _settings.Saturation = v;
        _settings.Save();
        BuildSaturationMenu();
    }

    private void SelectSmoothing(float v)
    {
        _settings.Smoothing = v;
        _settings.Save();
        BuildSmoothingMenu();
    }

    private static void ProcessColors(byte[] colors, float brightness, float saturation)
    {
        for (int i = 0; i < colors.Length; i += 3)
        {
            float r = colors[i], g = colors[i + 1], b = colors[i + 2];

            // Saturation: interpolate between luminance and original color
            if (Math.Abs(saturation - 1f) > 0.01f)
            {
                float lum = 0.299f * r + 0.587f * g + 0.114f * b;
                r = lum + saturation * (r - lum);
                g = lum + saturation * (g - lum);
                b = lum + saturation * (b - lum);
            }

            // Brightness: scale all channels
            if (Math.Abs(brightness - 1f) > 0.01f)
            {
                r *= brightness;
                g *= brightness;
                b *= brightness;
            }

            colors[i]     = (byte)Math.Clamp(r, 0, 255);
            colors[i + 1] = (byte)Math.Clamp(g, 0, 255);
            colors[i + 2] = (byte)Math.Clamp(b, 0, 255);
        }
    }

    private static void ApplySmoothing(byte[] current, byte[] last, float smoothing)
    {
        if (smoothing < 0.01f) { Array.Copy(current, last, current.Length); return; }
        for (int i = 0; i < current.Length; i++)
        {
            current[i] = (byte)Math.Round(last[i] * smoothing + current[i] * (1f - smoothing));
            last[i] = current[i];
        }
    }

    // Shows a small prompt dialog. Returns true if the user entered a valid value.
    private static bool PromptPercent(string title, int current, int min, int max, out int result)
    {
        result = current;
        using var form = new Form
        {
            Text = title, Width = 260, Height = 120,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition   = FormStartPosition.CenterScreen,
            MinimizeBox = false, MaximizeBox = false,
        };
        var tb  = new TextBox  { Text = current.ToString(), Left = 10, Top = 10, Width = 220 };
        var ok  = new Button   { Text = "OK",     Left = 60,  Top = 42, Width = 80, DialogResult = DialogResult.OK };
        var cancel = new Button{ Text = "Cancel", Left = 150, Top = 42, Width = 80, DialogResult = DialogResult.Cancel };
        form.Controls.AddRange([tb, ok, cancel]);
        form.AcceptButton = ok;
        form.CancelButton = cancel;

        if (form.ShowDialog() != DialogResult.OK) return false;
        if (!int.TryParse(tb.Text.Trim().TrimEnd('%'), out int v) || v < min || v > max)
        {
            MessageBox.Show($"Enter a number between {min} and {max}.", title,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        result = v;
        return true;
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

    static readonly string InstallPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "V2XAmbilight", "V2XAmbilight.exe");

    private static void ApplyAutoStart(bool enable)
    {
        using var key = Registry.CurrentUser.OpenSubKey(
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", writable: true)!;

        if (enable)
        {
            // Copy exe to permanent location if not already there
            string current = Application.ExecutablePath;
            if (!string.Equals(current, InstallPath, StringComparison.OrdinalIgnoreCase))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(InstallPath)!);
                File.Copy(current, InstallPath, overwrite: true);
            }
            key.SetValue("V2XAmbilight", $"\"{InstallPath}\"");
        }
        else
        {
            key.DeleteValue("V2XAmbilight", throwOnMissingValue: false);
        }
    }

    private async Task CheckForUpdateAsync(bool silent = true)
    {
        if (silent) await Task.Delay(3000).ConfigureAwait(false);
        _log.Append("Checking for updates…");
        var (hasUpdate, tag, url) = await Updater.CheckAsync().ConfigureAwait(false);

        _log.Invoke(() =>
        {
            if (hasUpdate)
            {
                _log.Append($"Update available: {tag} (current: v{Updater.CurrentVersion})");
                _updateItem.Text    = $"Update to {tag}";
                _updateItem.Tag     = url;
                _updateItem.Visible = true;

                if (!silent)
                {
                    var result = MessageBox.Show(
                        $"Version {tag} is available.\n\nInstall now? The app will restart automatically.",
                        "V2X Ambilight — Update Available",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                        _ = Updater.ApplyAsync(url, _log.Append);
                }
                else
                {
                    _tray.ShowBalloonTip(6000, "V2X Ambilight Update",
                        $"Version {tag} is available. Click the tray menu to update.", ToolTipIcon.Info);
                }
            }
            else
            {
                _log.Append($"Up to date (v{Updater.CurrentVersion})");
                if (!silent)
                    MessageBox.Show(
                        $"You're up to date!\n\nCurrent version: v{Updater.CurrentVersion}",
                        "V2X Ambilight — No Updates",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        });
    }

    private void OnCheckUpdate(object? sender, EventArgs e)
        => _ = CheckForUpdateAsync(silent: false);

    private void OnUpdateClick(object? sender, EventArgs e)
    {
        string url = _updateItem.Tag as string ?? "";
        if (string.IsNullOrEmpty(url)) return;
        _ = Updater.ApplyAsync(url, _log.Append);
    }

    private static string? FindPortBlocker()
    {
        // Known Creative Sound Blaster processes that may hold the COM port
        string[] candidates = ["SBCommand", "CTAudSvc", "V2XBridge", "SBConsole"];
        foreach (string name in candidates)
        {
            if (System.Diagnostics.Process.GetProcessesByName(name).Length > 0)
                return name;
        }
        return null;
    }

    private void OnExit(object? sender, EventArgs e)
    {
        _cts.Cancel();
        _timer.Stop();

        try { _device.SetColors(new byte[21]); } catch { } // LEDs off
        _tray.Visible = false;
        _log.AllowClose();
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
            _log.Dispose();
        }
        base.Dispose(disposing);
    }
}
