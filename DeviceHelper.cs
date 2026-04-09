using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AudioRoute;

public enum DeviceState : uint
{
    Active = 0x00000001,
    Disabled = 0x00000002,
    NotPresent = 0x00000004,
    Unplugged = 0x00000008,
    All = 0x0000000F
}

[StructLayout(LayoutKind.Sequential)]
internal struct PropertyKey
{
    public Guid fmtid;
    public uint pid;

    public static readonly PropertyKey DeviceFriendlyName = new()
    {
        fmtid = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0),
        pid = 14
    };
}

[StructLayout(LayoutKind.Explicit)]
internal struct PropVariant
{
    [FieldOffset(0)] public ushort vt;
    [FieldOffset(8)] public IntPtr pwszVal;

    public string? GetString()
    {
        return vt == 31 && pwszVal != IntPtr.Zero
            ? Marshal.PtrToStringUni(pwszVal)
            : null;
    }
}

[Guid("D666063F-1587-4E43-81F1-B948E807363F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDevice
{
    [PreserveSig]
    int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, out IntPtr ppInterface);

    [PreserveSig]
    int OpenPropertyStore(uint stgmAccess, out IPropertyStore ppProperties);

    [PreserveSig]
    int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

    [PreserveSig]
    int GetState(out DeviceState pdwState);
}

[Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPropertyStore
{
    [PreserveSig]
    int GetCount(out uint cProps);

    [PreserveSig]
    int GetAt(uint iProp, out PropertyKey pkey);

    [PreserveSig]
    int GetValue(ref PropertyKey key, out PropVariant pv);

    [PreserveSig]
    int SetValue(ref PropertyKey key, ref PropVariant propvar);

    [PreserveSig]
    int Commit();
}

[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDeviceEnumerator
{
    [PreserveSig]
    int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState dwStateMask, out IMMDeviceCollection ppDevices);

    [PreserveSig]
    int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);

    [PreserveSig]
    int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
}

[Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDeviceCollection
{
    [PreserveSig]
    int GetCount(out uint pcDevices);

    [PreserveSig]
    int Item(uint nDevice, out IMMDevice ppDevice);
}

[ComImport]
[Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
internal class MMDeviceEnumeratorComObject
{
}

public sealed class AudioDevice
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public EDataFlow Flow { get; init; }
    public bool IsDefault { get; init; }
}

public static class DeviceEnumerator
{
    private static readonly object DeviceCacheSync = new();
    private static readonly Dictionary<EDataFlow, CachedDevices> DeviceCache = new();
    private static readonly TimeSpan DeviceCacheDuration = TimeSpan.FromSeconds(2);

    public static IReadOnlyList<AudioDevice> EnumerateDevices(EDataFlow flow = EDataFlow.eAll)
    {
        var now = DateTimeOffset.UtcNow;

        lock (DeviceCacheSync)
        {
            if (DeviceCache.TryGetValue(flow, out var cachedDevices) &&
                cachedDevices.ExpiresAt > now)
            {
                return cachedDevices.Devices;
            }
        }

        var devices = EnumerateDevicesCore(flow);

        lock (DeviceCacheSync)
            DeviceCache[flow] = new CachedDevices(devices, now + DeviceCacheDuration);

        return devices;
    }

    private static AudioDevice[] EnumerateDevicesCore(EDataFlow flow)
    {
        var devices = new List<AudioDevice>();
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();

        try
        {
            string? defaultRenderDeviceId = null;
            string? defaultCaptureDeviceId = null;

            if (flow is EDataFlow.eRender or EDataFlow.eAll &&
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var defaultRenderDevice) == 0)
            {
                defaultRenderDevice.GetId(out defaultRenderDeviceId);
                Marshal.ReleaseComObject(defaultRenderDevice);
            }

            if (flow is EDataFlow.eCapture or EDataFlow.eAll &&
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia, out var defaultCaptureDevice) == 0)
            {
                defaultCaptureDevice.GetId(out defaultCaptureDeviceId);
                Marshal.ReleaseComObject(defaultCaptureDevice);
            }

            var flows = flow == EDataFlow.eAll
                ? new[] { EDataFlow.eRender, EDataFlow.eCapture }
                : new[] { flow };

            foreach (var currentFlow in flows)
            {
                if (enumerator.EnumAudioEndpoints(currentFlow, DeviceState.Active, out var collection) != 0)
                    continue;

                collection.GetCount(out var count);
                var requiredCapacity = devices.Count + (int)count;
                if (devices.Capacity < requiredCapacity)
                    devices.Capacity = requiredCapacity;

                for (uint i = 0; i < count; i++)
                {
                    if (collection.Item(i, out var device) != 0)
                        continue;

                    device.GetId(out var deviceId);

                    var name = deviceId;
                    if (device.OpenPropertyStore(0, out var propertyStore) == 0)
                    {
                        var key = PropertyKey.DeviceFriendlyName;
                        if (propertyStore.GetValue(ref key, out var variant) == 0)
                            name = variant.GetString() ?? deviceId;

                        Marshal.ReleaseComObject(propertyStore);
                    }

                    var isDefault = currentFlow == EDataFlow.eRender
                        ? deviceId == defaultRenderDeviceId
                        : deviceId == defaultCaptureDeviceId;

                    devices.Add(new AudioDevice
                    {
                        Id = deviceId,
                        Name = name,
                        Flow = currentFlow,
                        IsDefault = isDefault
                    });

                    Marshal.ReleaseComObject(device);
                }

                Marshal.ReleaseComObject(collection);
            }
        }
        finally
        {
            Marshal.ReleaseComObject(enumerator);
        }

        return devices.ToArray();
    }

    private sealed record CachedDevices(IReadOnlyList<AudioDevice> Devices, DateTimeOffset ExpiresAt);
}
