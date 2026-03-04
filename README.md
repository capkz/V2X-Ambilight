# V2X Ambilight

Lightweight Windows tray app that samples the bottom strip of your screen and drives the 7 LEDs on a **Creative Sound Blaster Katana V2X** in real time — no V2XBridge service, no drivers, talks directly to the device.

![tray icon green](https://img.shields.io/badge/status-active-brightgreen) ![license](https://img.shields.io/badge/license-GPL--v3-blue) [![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/chaul)

---

## Features

- Screen capture at 20 fps with GDI (works with desktop, browser, video, borderless games)
- Adjustable strip height (3 / 5 / 10 / 15% from bottom of screen)
- Brightness and Vibrancy (saturation) controls
- Multi-monitor support
- Start with Windows (installs itself to a permanent location automatically)
- Auto-update — checks GitHub Releases on startup, one-click install
- Log window for diagnostics

## Requirements

- Windows 10 / 11
- [.NET 8 Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) (x64) — prompted automatically if missing
- Creative Sound Blaster Katana V2X plugged in via USB

## Download

Grab the latest `V2XAmbilight.exe` from [Releases](../../releases/latest).

> **First run:** Windows may show a SmartScreen warning because the exe is not yet code-signed.
> Click **More info → Run anyway**, or right-click the file → **Properties → Unblock → OK**.

## Usage

1. **Close the Creative Sound Blaster Command app** before starting (it holds the COM port)
2. Run `V2XAmbilight.exe`
3. A tray icon appears — gray = searching, orange = connecting, green = running
4. Right-click the tray icon to adjust settings

### Start with Windows

Enable **Start with Windows** in the tray menu. The exe is automatically copied to
`%LocalAppData%\V2XAmbilight\V2XAmbilight.exe` and registered in the Windows startup registry.
You can then delete the original exe from wherever you downloaded it.

### Settings

| Option | Description |
|---|---|
| Enabled | Toggle LEDs on/off |
| Monitor | Which screen to sample |
| Strip Height | How tall a strip to sample from the bottom |
| Brightness | Scale LED brightness (50–200%) |
| Vibrancy | Boost colour saturation (50–300%) |
| Start with Windows | Auto-start on login |
| Check for Updates | Manually check for a new release |
| Show Log | Open diagnostic log window |

## Building

Requires [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0).

```powershell
.\publish.ps1
# Output: dist\V2XAmbilight.exe
```

Or push to GitHub — CI builds and publishes to Releases automatically.

## Device Protocol

- **VID:** `0x041E` **PID:** `0x3283`
- COM port auto-detected via WMI (VID/PID match)
- AES-256-GCM challenge-response auth, then binary LED packets
- LED packet: `5A 3A 20 2B 00 01 01` + 7×`[0xFF, B, G, R]` = 35 bytes

## Support

If this saved you some time and you want to buy me a coffee — [ko-fi.com/chaul](https://ko-fi.com/chaul) 🙏

## License

GNU General Public License v3.0 — see [LICENSE](LICENSE).
You are free to use, modify and share this project. Any derivative work must also be open source under GPL v3.
