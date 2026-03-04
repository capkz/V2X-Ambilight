using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace V2XAmbilight;

/// <summary>
/// Captures the bottom strip of a screen and returns 7 averaged RGB color zones.
/// Uses GDI BitBlt via Graphics.CopyFromScreen — works correctly with borderless
/// windowed games. For fullscreen-exclusive titles, use borderless windowed mode.
/// </summary>
public sealed class ScreenSampler : IDisposable
{
    private Bitmap? _buffer;
    private int _lastWidth, _lastHeight;

    /// <summary>
    /// Returns 21 bytes: 7 × [R, G, B], each averaged from one horizontal zone
    /// of the bottom <paramref name="stripPercent"/>% of <paramref name="screen"/>.
    /// </summary>
    public byte[] Sample(Screen screen, int stripPercent)
    {
        var bounds  = screen.Bounds;
        int stripH  = Math.Max(1, bounds.Height * stripPercent / 100);
        int captureY = bounds.Bottom - stripH;

        // Reuse bitmap if dimensions haven't changed
        if (_buffer == null || _lastWidth != bounds.Width || _lastHeight != stripH)
        {
            _buffer?.Dispose();
            _buffer      = new Bitmap(bounds.Width, stripH, PixelFormat.Format32bppRgb);
            _lastWidth   = bounds.Width;
            _lastHeight  = stripH;
        }

        using var g = Graphics.FromImage(_buffer);
        g.CopyFromScreen(bounds.Left, captureY, 0, 0,
            new Size(bounds.Width, stripH), CopyPixelOperation.SourceCopy);

        return ComputeZoneColors(_buffer, zoneCount: 7);
    }

    private static byte[] ComputeZoneColors(Bitmap bmp, int zoneCount)
    {
        // Lock pixels for fast bulk read
        var area = new Rectangle(0, 0, bmp.Width, bmp.Height);
        var bd   = bmp.LockBits(area, ImageLockMode.ReadOnly, PixelFormat.Format32bppRgb);
        int stride = bd.Stride, w = bmp.Width, h = bmp.Height;

        byte[] pixels = new byte[stride * h];
        Marshal.Copy(bd.Scan0, pixels, 0, pixels.Length);
        bmp.UnlockBits(bd);

        byte[] result  = new byte[zoneCount * 3];
        int zoneW = w / zoneCount;

        for (int z = 0; z < zoneCount; z++)
        {
            int startX = z * zoneW;
            int endX   = z == zoneCount - 1 ? w : startX + zoneW;

            long r = 0, g = 0, b = 0, count = 0;

            // Sample every 4th pixel in each axis — plenty of accuracy at far lower cost
            for (int y = 0; y < h; y += 4)
            for (int x = startX; x < endX; x += 4)
            {
                int off = y * stride + x * 4; // Format32bppRgb: B G R _
                b += pixels[off];
                g += pixels[off + 1];
                r += pixels[off + 2];
                count++;
            }

            if (count > 0)
            {
                result[z * 3]     = (byte)(r / count);
                result[z * 3 + 1] = (byte)(g / count);
                result[z * 3 + 2] = (byte)(b / count);
            }
        }

        return result;
    }

    public void Dispose() => _buffer?.Dispose();
}
