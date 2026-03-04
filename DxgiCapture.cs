using System.Runtime.InteropServices;

namespace V2XAmbilight;

/// <summary>
/// DXGI Desktop Duplication capture — works with fullscreen-exclusive DirectX games.
/// Must be used from a single thread. Call TryInitialize before CaptureZones.
/// </summary>
internal sealed class DxgiCapture : IDisposable
{
    const int WAIT_TIMEOUT  = unchecked((int)0x887A0027);
    const int ACCESS_LOST   = unchecked((int)0x887A0026);
    const int NOT_AVAILABLE = unchecked((int)0x887A0022);
    const int D3D11_SDK_VER = 7;

    ID3D11Device?           _device;
    ID3D11DeviceContext?    _context;
    IDXGIOutputDuplication? _dup;
    ID3D11Texture2D?        _staging;
    int _stagingW, _stagingH;
    byte[]? _lastFrame;
    bool _disposed;

    public string ScreenName { get; private set; } = "";
    public bool IsValid => _dup != null && _device != null;

    public bool TryInitialize(string screenDeviceName)
    {
        Cleanup();
        ScreenName = screenDeviceName;

        int hr = D3D11CreateDevice(0, 1, 0, 0, 0, 0, D3D11_SDK_VER,
            out nint pDev, out _, out nint pCtx);
        if (hr < 0) return false;

        _device  = (ID3D11Device) Marshal.GetObjectForIUnknown(pDev);
        _context = (ID3D11DeviceContext) Marshal.GetObjectForIUnknown(pCtx);
        Marshal.Release(pDev);
        Marshal.Release(pCtx);

        try
        {
            var dxgiDev = (IDXGIDevice)_device;
            hr = dxgiDev.GetAdapter(out IDXGIAdapter pAdapter);
            if (hr < 0) return false;

            IDXGIOutput1? match = null;
            for (uint i = 0; pAdapter.EnumOutputs(i, out IDXGIOutput pOut) == 0; i++)
            {
                pOut.GetDesc(out DXGI_OUTPUT_DESC desc);
                string devName = desc.DeviceName?.TrimEnd('\0') ?? "";
                if (devName == screenDeviceName)
                {
                    match = (IDXGIOutput1)pOut;
                    break;
                }
                Marshal.ReleaseComObject(pOut);
            }
            Marshal.ReleaseComObject(pAdapter);

            if (match == null) return false;

            hr = match.DuplicateOutput(_device, out _dup);
            Marshal.ReleaseComObject(match);
            return hr == 0;
        }
        catch { return false; }
    }

    /// <summary>
    /// Captures the bottom strip and returns 7-zone RGB bytes, or null if DXGI failed
    /// (caller should fall back to GDI). Returns cached last frame on timeout (static desktop).
    /// </summary>
    public byte[]? CaptureZones(int screenW, int screenH, int stripPercent, int zoneCount)
    {
        if (_dup == null || _context == null || _device == null) return null;

        int stripH = Math.Max(1, screenH * stripPercent / 100);

        int hr = _dup.AcquireNextFrame(50, out DXGI_OUTDUPL_FRAME_INFO_STUB _, out nint pRes);

        if (hr == WAIT_TIMEOUT)  return _lastFrame; // desktop unchanged — reuse last result
        if (hr == ACCESS_LOST || hr == NOT_AVAILABLE) { Cleanup(); return null; }
        if (hr < 0) return null;

        try
        {
            var tex = (ID3D11Texture2D)Marshal.GetObjectForIUnknown(pRes);
            Marshal.Release(pRes);

            try
            {
                tex.GetDesc(out D3D11_TEXTURE2D_DESC texDesc);
                uint w = texDesc.Width, h = texDesc.Height;
                uint srcY       = h > (uint)stripH ? h - (uint)stripH : 0;
                uint actualH    = h - srcY;

                if (_staging == null || _stagingW != (int)w || _stagingH != (int)actualH)
                {
                    if (_staging != null) Marshal.ReleaseComObject(_staging);
                    var desc = new D3D11_TEXTURE2D_DESC
                    {
                        Width = w, Height = actualH, MipLevels = 1, ArraySize = 1,
                        Format = texDesc.Format,
                        SampleDesc = new DXGI_SAMPLE_DESC { Count = 1 },
                        Usage = 3,           // D3D11_USAGE_STAGING
                        CPUAccessFlags = 0x20000  // D3D11_CPU_ACCESS_READ
                    };
                    hr = _device.CreateTexture2D(ref desc, 0, out nint pStaging);
                    if (hr < 0) return null;
                    _staging = (ID3D11Texture2D)Marshal.GetObjectForIUnknown(pStaging);
                    Marshal.Release(pStaging);
                    _stagingW = (int)w; _stagingH = (int)actualH;
                }

                var box = new D3D11_BOX { left = 0, top = srcY, front = 0, right = w, bottom = h, back = 1 };
                _context.CopySubresourceRegion(
                    (ID3D11Resource)_staging, 0, 0, 0, 0,
                    (ID3D11Resource)tex, 0, ref box);

                hr = _context.Map((ID3D11Resource)_staging, 0, 1, 0, out D3D11_MAPPED_SUBRESOURCE mapped);
                if (hr < 0) return null;
                try
                {
                    _lastFrame = ComputeZones(mapped.pData, (int)w, (int)actualH, (int)mapped.RowPitch, zoneCount);
                    return _lastFrame;
                }
                finally { _context.Unmap((ID3D11Resource)_staging, 0); }
            }
            finally { Marshal.ReleaseComObject(tex); }
        }
        finally { _dup?.ReleaseFrame(); }
    }

    static byte[] ComputeZones(nint data, int w, int h, int stride, int zones)
    {
        // Copy strip pixels to managed memory, then average per zone
        int len = stride * h;
        byte[] pixels = new byte[len];
        Marshal.Copy(data, pixels, 0, len);

        byte[] result = new byte[zones * 3];
        int zw = w / zones;

        for (int z = 0; z < zones; z++)
        {
            int x0 = z * zw, x1 = z == zones - 1 ? w : x0 + zw;
            long r = 0, g = 0, b = 0, n = 0;

            for (int y = 0; y < h; y += 4)
            for (int x = x0; x < x1; x += 4)
            {
                int off = y * stride + x * 4; // BGRA layout from DXGI
                b += pixels[off];
                g += pixels[off + 1];
                r += pixels[off + 2];
                n++;
            }

            if (n > 0)
            {
                result[z * 3]     = (byte)(r / n);
                result[z * 3 + 1] = (byte)(g / n);
                result[z * 3 + 2] = (byte)(b / n);
            }
        }
        return result;
    }

    void Cleanup()
    {
        _lastFrame = null;
        _stagingW = _stagingH = 0;
        if (_staging != null) { Marshal.ReleaseComObject(_staging);  _staging  = null; }
        if (_dup     != null) { Marshal.ReleaseComObject(_dup);      _dup      = null; }
        if (_context != null) { Marshal.ReleaseComObject(_context);  _context  = null; }
        if (_device  != null) { Marshal.ReleaseComObject(_device);   _device   = null; }
    }

    public void Dispose() { if (!_disposed) { _disposed = true; Cleanup(); } }

    // ─── P/Invoke ─────────────────────────────────────────────────────────────

    [DllImport("d3d11.dll")]
    static extern int D3D11CreateDevice(
        nint pAdapter, int DriverType, nint Software, uint Flags,
        nint pFeatureLevels, uint FeatureLevels, uint SDKVersion,
        out nint ppDevice, out int pFeatureLevel, out nint ppImmediateContext);

    // ─── Structs ──────────────────────────────────────────────────────────────

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct DXGI_OUTPUT_DESC
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;
        public int Left, Top, Right, Bottom; // DesktopCoordinates RECT
        public int AttachedToDesktop, Rotation;
        public nint Monitor;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct D3D11_TEXTURE2D_DESC
    {
        public uint Width, Height, MipLevels, ArraySize;
        public int  Format;         // DXGI_FORMAT
        public DXGI_SAMPLE_DESC SampleDesc;
        public int  Usage;          // D3D11_USAGE
        public uint BindFlags, CPUAccessFlags, MiscFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DXGI_SAMPLE_DESC { public uint Count, Quality; }

    [StructLayout(LayoutKind.Sequential)]
    struct D3D11_MAPPED_SUBRESOURCE { public nint pData; public uint RowPitch, DepthPitch; }

    [StructLayout(LayoutKind.Sequential)]
    struct D3D11_BOX { public uint left, top, front, right, bottom, back; }

    // 48-byte stub for DXGI_OUTDUPL_FRAME_INFO — we only need to discard it
    [StructLayout(LayoutKind.Sequential, Size = 48)]
    struct DXGI_OUTDUPL_FRAME_INFO_STUB { }

    // ─── COM Interfaces ───────────────────────────────────────────────────────
    // Vtable stubs for methods we don't call are declared as void with no params.
    // Declaration order determines vtable slot — that's all that matters for stubs.

    [ComImport, Guid("aec22fb8-76f3-4639-9be0-28eb43a67a2e"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIObject               // [3] SetPrivateData  [4] SetPrivateDataInterface  [5] GetPrivateData  [6] GetParent
    {
        void SetPrivateData(); void SetPrivateDataInterface(); void GetPrivateData(); void GetParent();
    }

    [ComImport, Guid("2411e7e1-12ac-4ccf-bd14-9798e8534dc0"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIAdapter : IDXGIObject  // [7] EnumOutputs  [8] GetDesc  [9] CheckInterfaceSupport
    {
        [PreserveSig] int EnumOutputs(uint Output, out IDXGIOutput ppOutput);
        void GetDesc(); void CheckInterfaceSupport();
    }

    [ComImport, Guid("54ec77fa-1377-44e6-8c32-88fd5f44c84c"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIDevice : IDXGIObject  // [7] GetAdapter  [8-11] stubs
    {
        [PreserveSig] int GetAdapter(out IDXGIAdapter pAdapter);
        void CreateSurface(); void QueryResourceResidency(); void SetGPUThreadPriority(); void GetGPUThreadPriority();
    }

    [ComImport, Guid("ae02eedb-c735-4690-8d52-5a8dc20213aa"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIOutput : IDXGIObject  // [7] GetDesc  [8-18] stubs
    {
        [PreserveSig] int GetDesc(out DXGI_OUTPUT_DESC pDesc);
        void GetDisplayModeList(); void FindClosestMatchingMode(); void WaitForVBlank();
        void TakeOwnership(); void ReleaseOwnership(); void GetGammaControlCapabilities();
        void SetGammaControl(); void GetGammaControl(); void SetDisplaySurface();
        void GetDisplaySurfaceData(); void GetFrameStatistics();
    }

    [ComImport, Guid("00cddea8-939b-4b83-a340-a685226666cc"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIOutput1 : IDXGIOutput  // [19-21] stubs  [22] DuplicateOutput
    {
        void GetDisplayModeList1(); void FindClosestMatchingMode1(); void GetDisplaySurfaceData1();
        [PreserveSig] int DuplicateOutput(ID3D11Device pDevice, out IDXGIOutputDuplication ppOutputDuplication);
    }

    [ComImport, Guid("191cfac3-a341-470d-b26e-a864f428319c"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDXGIOutputDuplication : IDXGIObject  // [7] GetDesc  [8] AcquireNextFrame  [9-13] stubs  [14] ReleaseFrame
    {
        void GetDesc_Dup();
        [PreserveSig] int AcquireNextFrame(uint TimeoutMs, out DXGI_OUTDUPL_FRAME_INFO_STUB pFrameInfo, out nint ppResource);
        void GetFrameDirtyRects(); void GetFrameMoveRects(); void GetFramePointerShape();
        void MapDesktopSurface(); void UnMapDesktopSurface();
        [PreserveSig] int ReleaseFrame();
    }

    [ComImport, Guid("1841e5c8-16b0-489b-bcc8-44cfb0d5deae"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ID3D11DeviceChild  // [3] GetDevice  [4] GetPrivateData  [5] SetPrivateData  [6] SetPrivateDataInterface
    {
        void GetDevice(); void GetPrivateData(); void SetPrivateData(); void SetPrivateDataInterface();
    }

    [ComImport, Guid("dc8e63f3-d12b-4952-b47b-5e45026a862d"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ID3D11Resource : ID3D11DeviceChild  // [7] GetType  [8] SetEvictionPriority  [9] GetEvictionPriority
    {
        void GetType_Res(); void SetEvictionPriority(); void GetEvictionPriority();
    }

    [ComImport, Guid("6f15aaf2-d208-4e89-9ab4-489535d34f9c"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ID3D11Texture2D : ID3D11Resource  // [10] GetDesc
    {
        [PreserveSig] int GetDesc(out D3D11_TEXTURE2D_DESC pDesc);
    }

    [ComImport, Guid("db6f6ddb-ac77-4e88-8253-819df9bbf140"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ID3D11Device  // [3] CreateBuffer ... [5] CreateTexture2D ... [40] GetImmediateContext
    {
        void CreateBuffer(); void CreateTexture1D();
        [PreserveSig] int CreateTexture2D(ref D3D11_TEXTURE2D_DESC pDesc, nint pData, out nint ppTexture);
        void CreateTexture3D(); void CreateShaderResourceView(); void CreateUnorderedAccessView();
        void CreateRenderTargetView(); void CreateDepthStencilView(); void CreateInputLayout();
        void CreateVertexShader(); void CreateGeometryShader(); void CreateGeometryShaderWithStreamOutput();
        void CreatePixelShader(); void CreateHullShader(); void CreateDomainShader();
        void CreateComputeShader(); void CreateClassLinkage(); void CreateBlendState();
        void CreateDepthStencilState(); void CreateRasterizerState(); void CreateSamplerState();
        void CreateQuery(); void CreatePredicate(); void CreateCounter();
        void CreateDeferredContext(); void OpenSharedResource(); void CheckFormatSupport();
        void CheckMultisampleQualityLevels(); void CheckCounterInfo(); void CheckCounter();
        void CheckFeatureSupport(); void GetPrivateData_Dev(); void SetPrivateData_Dev();
        void SetPrivateDataInterface_Dev(); void GetFeatureLevel(); void GetCreationFlags();
        void GetDeviceRemovedReason();
        void GetImmediateContext();  // [40] — unused; context obtained from D3D11CreateDevice
        void SetExceptionMode(); void GetExceptionMode();
    }

    [ComImport, Guid("c0bfa96c-e089-44fb-8eaf-26f8796190da"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ID3D11DeviceContext : ID3D11DeviceChild
    // Vtable (own methods start at [7]):
    //  [7]  VSSetConstantBuffers   [8]  PSSetShaderResources  [9]  PSSetShader
    //  [10] PSSetSamplers          [11] VSSetShader           [12] DrawIndexed
    //  [13] Draw                   [14] Map                   [15] Unmap
    //  [16] PSSetConstantBuffers   [17] IASetInputLayout      [18] IASetVertexBuffers
    //  [19] IASetIndexBuffer       [20] DrawIndexedInstanced  [21] DrawInstanced
    //  [22] GSSetConstantBuffers   [23] GSSetShader           [24] IASetPrimitiveTopology
    //  [25] VSSetShaderResources   [26] VSSetSamplers         [27] Begin
    //  [28] End                    [29] GetData               [30] SetPredication
    //  [31] GSSetShaderResources   [32] GSSetSamplers         [33] OMSetRenderTargets
    //  [34] OMSetRenderTargetsAndUnorderedAccessViews         [35] OMSetBlendState
    //  [36] OMSetDepthStencilState [37] SOSetTargets          [38] DrawAuto
    //  [39] DrawIndexedInstancedIndirect                      [40] DrawInstancedIndirect
    //  [41] Dispatch               [42] DispatchIndirect      [43] RSSetState
    //  [44] RSSetViewports         [45] RSSetScissorRects     [46] CopySubresourceRegion
    //  [47] CopyResource
    {
        void VSSetConstantBuffers(); void PSSetShaderResources(); void PSSetShader();
        void PSSetSamplers(); void VSSetShader(); void DrawIndexed(); void Draw();
        [PreserveSig] int Map(ID3D11Resource res, uint sub, int mapType, uint flags, out D3D11_MAPPED_SUBRESOURCE mapped);
        void Unmap(ID3D11Resource res, uint sub);
        void PSSetConstantBuffers(); void IASetInputLayout(); void IASetVertexBuffers();
        void IASetIndexBuffer(); void DrawIndexedInstanced(); void DrawInstanced();
        void GSSetConstantBuffers(); void GSSetShader(); void IASetPrimitiveTopology();
        void VSSetShaderResources(); void VSSetSamplers(); void Begin(); void End();
        void GetData(); void SetPredication(); void GSSetShaderResources(); void GSSetSamplers();
        void OMSetRenderTargets(); void OMSetRenderTargetsAndUnorderedAccessViews();
        void OMSetBlendState(); void OMSetDepthStencilState(); void SOSetTargets();
        void DrawAuto(); void DrawIndexedInstancedIndirect(); void DrawInstancedIndirect();
        void Dispatch(); void DispatchIndirect(); void RSSetState(); void RSSetViewports();
        void RSSetScissorRects();
        void CopySubresourceRegion(ID3D11Resource dst, uint dstSub, uint dstX, uint dstY, uint dstZ,
                                   ID3D11Resource src, uint srcSub, ref D3D11_BOX box);
        void CopyResource(ID3D11Resource dst, ID3D11Resource src);
    }
}
