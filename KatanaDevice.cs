using System.IO.Ports;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace V2XAmbilight;

/// <summary>
/// Manages the serial connection to the Katana V2X (VID 0x041E, PID 0x3283),
/// performs AES-256-GCM challenge-response auth, and sends LED color commands.
/// </summary>
public sealed class KatanaDevice : IDisposable
{
    private static readonly byte[] KeyData = [
        0xD3, 0x1A, 0x21, 0x27, 0x9B, 0xE3, 0x46, 0xF0,
        0x99, 0x9D, 0x6E, 0xC4, 0xC3, 0xFE, 0xBE, 0x98,
        0x90, 0x18, 0x69, 0xC1, 0x18, 0xFB, 0xB1, 0x25,
        0x6E, 0x0C, 0xE0, 0x7B,
    ];

    private const int VendorId  = 0x041E;
    private const int ProductId = 0x3283;
    private const int BaudRate  = 115200;

    private SerialPort? _port;
    private readonly Action<string> _log;

    public bool IsConnected => _port?.IsOpen == true;

    public KatanaDevice(Action<string> log) => _log = log;

    // -------------------------------------------------------------------------
    // Connection
    // -------------------------------------------------------------------------

    public async Task ConnectAsync(CancellationToken ct)
    {
        // Close any stale port from a previous attempt
        _port?.Close();
        _port?.Dispose();
        _port = null;

        string portName = FindComPort()
            ?? throw new InvalidOperationException("Katana V2X not found on any COM port.");

        _log($"Opening {portName}");
        _port = new SerialPort(portName, BaudRate) { ReadTimeout = 3000, WriteTimeout = 3000 };
        _port.Open();

        await AuthenticateAsync(ct);
        await SwitchToCommandModeAsync(ct);
        _log($"Ready on {portName}");
    }

    public void Disconnect()
    {
        _port?.Close();
        _port?.Dispose();
        _port = null;
    }

    // -------------------------------------------------------------------------
    // Auth (AES-256-GCM challenge-response)
    // -------------------------------------------------------------------------

    private async Task AuthenticateAsync(CancellationToken ct)
    {
        _port!.DiscardInBuffer();

        // Send whoareyou first — safe in both text and binary mode.
        // Sending binary bytes before this corrupts the text-mode command parser.
        _port.ReadTimeout = 1000;
        WriteLine("whoareyou");
        await Task.Delay(50, ct);

        string? probe = null;
        try   { probe = ReadLine(); }
        catch (TimeoutException) { /* handled below */ }
        finally { _port.ReadTimeout = 3000; }

        if (probe is not null)
        {
            if (probe.Length > 0 && (byte)probe[0] == 0x5A)
            {
                _log("Already in binary command mode");
                _port.DiscardInBuffer();
                return;
            }
            if (probe.StartsWith("Unknown command", StringComparison.Ordinal) ||
                probe.StartsWith("unlock_OK",       StringComparison.Ordinal))
            {
                _log("Already authenticated");
                return;
            }
            if (!probe.StartsWith("whoareyou", StringComparison.Ordinal))
                throw new InvalidOperationException($"Unexpected auth response: {probe}");
        }
        else
        {
            // No text response — try binary ping as fallback
            _port.DiscardInBuffer();
            byte[] ping = [0x5A, 0x03, 0x00];
            _port.Write(ping, 0, ping.Length);
            await Task.Delay(500, ct);

            if (_port.BytesToRead > 0 && _port.ReadByte() == 0x5A)
            {
                _log("Already in binary command mode (silent)");
                _port.DiscardInBuffer();
                return;
            }

            _port.DiscardInBuffer();
            WriteLine("whoareyou");
            await Task.Delay(50, ct);
            probe = ReadLine(); // throws TimeoutException if truly unresponsive
        }

        // Parse challenge: "whoareyou" [h0 h1] [p0 p1] [32-byte nonce]
        byte[] challenge = Encoding.Latin1.GetBytes(probe.TrimEnd('\r', '\n'));
        if (challenge.Length < 45)
            throw new InvalidOperationException("Challenge packet too short");

        byte h0 = challenge[9], h1 = challenge[10];
        byte pid0 = challenge[11], pid1 = challenge[12];
        byte[] nonce = challenge[13..45];

        byte[] key = BuildKey(h0, h1, pid0, pid1);
        byte[] iv  = RandomNumberGenerator.GetBytes(16);

        byte[] ciphertext = new byte[nonce.Length];
        byte[] tag        = new byte[16];
        using var aesGcm  = new AesGcm(key, 16);
        aesGcm.Encrypt(iv[..12], nonce, ciphertext, tag);

        byte[] payload = [.. "unlock"u8.ToArray(), .. iv, .. ciphertext, .. tag];
        string resp = SendBinaryAndReadLine(payload);

        if (!resp.StartsWith("unlock_OK", StringComparison.Ordinal))
            throw new InvalidOperationException($"Auth failed: {resp}");

        _log("Auth OK");
    }

    private static byte[] BuildKey(byte h0, byte h1, byte pid0, byte pid1)
    {
        byte[] key = new byte[32];
        key[0] = h0; key[1] = h1;
        KeyData.CopyTo(key, 2);
        key[30] = pid0; key[31] = pid1;
        return key;
    }

    private async Task SwitchToCommandModeAsync(CancellationToken ct)
    {
        WriteLine("SW_MODE1");
        await Task.Delay(100, ct);
        _port!.DiscardInBuffer();

        byte[] ping = [0x5A, 0x03, 0x00];
        _port.Write(ping, 0, ping.Length);
        await Task.Delay(50, ct);
        _port.DiscardInBuffer();

        SendCommand([0x5A, 0x3A, 0x02, 0x25, 0x01]);       // lighting on
        SendCommand([0x5A, 0x3A, 0x03, 0x37, 0x00, 0x07]); // 7 color slots
        SendCommand([0x5A, 0x3A, 0x03, 0x29, 0x00, 0x01]); // static mode
    }

    // -------------------------------------------------------------------------
    // LED commands
    // -------------------------------------------------------------------------

    /// <summary>21 bytes: 7 × [R, G, B]. Thread-safe to call from any thread.</summary>
    public void SetColors(ReadOnlySpan<byte> rgbColors)
    {
        if (_port is null || !_port.IsOpen) return;

        // Packet: 5A 3A 20 2B 00 01 01 [0xFF B G R] × 7  (35 bytes)
        byte[] packet = new byte[35];
        packet[0] = 0x5A; packet[1] = 0x3A; packet[2] = 0x20;
        packet[3] = 0x2B; packet[4] = 0x00; packet[5] = 0x01; packet[6] = 0x01;

        for (int i = 0; i < 7; i++)
        {
            int src = i * 3, dest = 7 + i * 4;
            byte r = src     < rgbColors.Length ? rgbColors[src]     : (byte)0;
            byte g = src + 1 < rgbColors.Length ? rgbColors[src + 1] : (byte)0;
            byte b = src + 2 < rgbColors.Length ? rgbColors[src + 2] : (byte)0;
            packet[dest] = 0xFF; packet[dest+1] = b; packet[dest+2] = g; packet[dest+3] = r;
        }

        SendCommand(packet);
    }

    private void SendCommand(ReadOnlySpan<byte> data)
    {
        if (_port is null || !_port.IsOpen) return;
        _port.Write(data.ToArray(), 0, data.Length);
    }

    // -------------------------------------------------------------------------
    // Serial helpers
    // -------------------------------------------------------------------------

    private void WriteLine(string text)
    {
        byte[] data = Encoding.ASCII.GetBytes(text + "\r\n");
        _port!.Write(data, 0, data.Length);
    }

    private string ReadLine()
    {
        var sb = new StringBuilder();
        while (true)
        {
            int b = _port!.ReadByte();
            if (b == '\n') break;
            if (b != '\r') sb.Append((char)b);
        }
        return sb.ToString();
    }

    private string SendBinaryAndReadLine(byte[] payload)
    {
        byte[] wire = [.. payload, (byte)'\r', (byte)'\n'];
        _port!.Write(wire, 0, wire.Length);
        return ReadLine();
    }

    // -------------------------------------------------------------------------
    // COM port discovery via WMI
    // -------------------------------------------------------------------------

    private static string? FindComPort()
    {
        string vidStr = $"VID_{VendorId:X4}";
        string pidStr = $"PID_{ProductId:X4}";

        using var searcher = new ManagementObjectSearcher(
            "SELECT Name, DeviceID FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'");

        foreach (ManagementObject obj in searcher.Get())
        {
            string? id   = obj["DeviceID"]?.ToString();
            string? name = obj["Name"]?.ToString();
            if (id is null || name is null) continue;

            if (!id.Contains(vidStr, StringComparison.OrdinalIgnoreCase) ||
                !id.Contains(pidStr, StringComparison.OrdinalIgnoreCase)) continue;

            int s = name.IndexOf("(COM", StringComparison.Ordinal);
            int e = s >= 0 ? name.IndexOf(')', s) : -1;
            if (s >= 0 && e > s) return name[(s + 1)..e];
        }

        return null;
    }

    public void Dispose()
    {
        _port?.Close();
        _port?.Dispose();
    }
}
