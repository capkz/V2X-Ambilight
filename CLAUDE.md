# V2X Ambilight

Lightweight Windows tray app that samples the bottom strip of a monitor and drives the 7 LEDs on a Creative Sound Blaster Katana V2X over USB serial.

## Project structure

```
V2XAmbilight.csproj   ← WinForms .NET 8 Windows, System.Management only
Program.cs            ← entry point, single-instance mutex
Settings.cs           ← JSON config at %AppData%\V2XAmbilight\settings.json
KatanaDevice.cs       ← AES-256-GCM auth + binary LED protocol over COM4
ScreenSampler.cs      ← GDI screen capture, 7-zone LockBits color averaging
TrayApp.cs            ← NotifyIcon, menus, capture timer, connect loop
publish.ps1           ← builds self-contained win-x64 exe (needs .NET 8 SDK)
.github/workflows/    ← CI builds V2XAmbilight.exe and commits to repo root
```

## Build

```powershell
# Local (needs .NET 8 SDK)
.\publish.ps1
# Output: dist\V2XAmbilight.exe

# Or just push to GitHub — CI builds and commits V2XAmbilight.exe automatically
```

## Device

- VID: 0x041E, PID: 0x3283 (Creative Sound Blaster Katana V2X)
- COM port: auto-detected via WMI (VID/PID match)
- Protocol: AES-256-GCM challenge-response auth, then binary LED packets
- LED packet: `5A 3A 20 2B 00 01 01` + 7×[0xFF, B, G, R] = 35 bytes

## Key protocol notes

- Send `whoareyou\r\n` FIRST — sending binary bytes before this corrupts the text parser
- After auth: send `SW_MODE1\r\n`, then binary ping `5A 03 00`
- Device stays in binary command mode after Creative app has authed it
- AES key: `[h0, h1, KEY_DATA(28 bytes), pid_lo, pid_hi]`

## Settings (tray menu)

- Monitor: which screen to capture (default: primary)
- Strip Height: 3% / 5% / 10% / 15% of screen height from bottom
- Frame rate: 20 fps (hardcoded, change `FrameRate` in Settings.cs)
- Start with Windows: writes `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run`

## Notes

- No V2XBridge service needed — this app talks directly to COM4
- GDI capture works with borderless windowed games; for fullscreen-exclusive use borderless mode
- Tray icon: gray = no device, orange = connecting, green = running
- Device reconnects automatically every 5 s on failure
