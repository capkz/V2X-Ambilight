namespace V2XAmbilight;

/// <summary>
/// Simple floating log window accessible from the tray menu.
/// Thread-safe: Append() can be called from any thread.
/// </summary>
internal sealed class LogWindow : Form
{
    private readonly RichTextBox _box;
    private const int MaxLines = 500;

    public LogWindow()
    {
        Text            = "V2X Ambilight — Log";
        Size            = new Size(560, 340);
        MinimumSize     = new Size(360, 200);
        StartPosition   = FormStartPosition.Manual;
        Location        = new Point(
            Screen.PrimaryScreen!.WorkingArea.Right  - Width  - 16,
            Screen.PrimaryScreen!.WorkingArea.Bottom - Height - 16);
        FormBorderStyle = FormBorderStyle.SizableToolWindow;
        ShowInTaskbar   = false;

        _box = new RichTextBox
        {
            Dock        = DockStyle.Fill,
            ReadOnly    = true,
            BackColor   = Color.FromArgb(18, 18, 18),
            ForeColor   = Color.FromArgb(220, 220, 220),
            Font        = new Font("Consolas", 9f),
            BorderStyle = BorderStyle.None,
            WordWrap    = false,
            ScrollBars  = RichTextBoxScrollBars.Both,
        };

        var clear = new Button
        {
            Text   = "Clear",
            Dock   = DockStyle.Bottom,
            Height = 26,
        };
        clear.Click += (_, _) => _box.Clear();

        Controls.Add(_box);
        Controls.Add(clear);

        // Don't destroy on close — just hide, so log persists
        FormClosing += (_, e) => { e.Cancel = true; Hide(); };

        // Force handle creation so BeginInvoke works before the window is ever shown
        _ = Handle;

        Append("Log started.");
    }

    public void Append(string message)
    {
        string line = $"[{DateTime.Now:HH:mm:ss}] {message}";

        if (_box.InvokeRequired)
        {
            _box.BeginInvoke(() => AppendLine(line));
        }
        else
        {
            AppendLine(line);
        }
    }

    private void AppendLine(string line)
    {
        // Trim oldest lines if over limit
        if (_box.Lines.Length > MaxLines)
        {
            _box.Select(0, _box.GetFirstCharIndexFromLine(MaxLines / 2));
            _box.SelectedText = "";
        }

        _box.AppendText(line + Environment.NewLine);
        _box.ScrollToCaret();
    }
}
