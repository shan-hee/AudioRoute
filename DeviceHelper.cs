using System;
using System.Collections.Generic;
using System.Diagnostics;
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

public sealed class AudioDevice
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public EDataFlow Flow { get; init; }
    public bool IsDefault { get; init; }
}

public static class DeviceEnumerator
{
    private static readonly Guid MMDeviceEnumeratorClassId = new("BCDE0395-E52F-467C-8E3D-C4579291692E");
    private static readonly TimeSpan DeviceCacheDuration = TimeSpan.FromSeconds(2);
    private static readonly ExpiringCache<EDataFlow, IReadOnlyList<AudioDevice>> DeviceCache = new(DeviceCacheDuration, 8);

    public static IReadOnlyList<AudioDevice> EnumerateDevices(EDataFlow flow = EDataFlow.eAll)
    {
        if (DeviceCache.TryGetValue(flow, out var cachedDevices))
            return cachedDevices!;

        var devices = EnumerateDevicesCore(flow);
        DeviceCache.Set(flow, devices);
        return devices;
    }

    public static void InvalidateCache(EDataFlow flow = EDataFlow.eAll)
    {
        if (flow == EDataFlow.eAll)
        {
            DeviceCache.Clear();
            return;
        }

        DeviceCache.Remove(flow);
        DeviceCache.Remove(EDataFlow.eAll);
    }

    private static AudioDevice[] EnumerateDevicesCore(EDataFlow flow)
    {
        var devices = new List<AudioDevice>();
        var enumerator = CreateDeviceEnumerator();

        try
        {
            string? defaultRenderDeviceId = null;
            string? defaultCaptureDeviceId = null;

            if (flow is EDataFlow.eRender or EDataFlow.eAll &&
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var defaultRenderDevice) == 0)
            {
                try
                {
                    defaultRenderDevice.GetId(out defaultRenderDeviceId);
                }
                finally
                {
                    ReleaseComObject(defaultRenderDevice);
                }
            }

            if (flow is EDataFlow.eCapture or EDataFlow.eAll &&
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia, out var defaultCaptureDevice) == 0)
            {
                try
                {
                    defaultCaptureDevice.GetId(out defaultCaptureDeviceId);
                }
                finally
                {
                    ReleaseComObject(defaultCaptureDevice);
                }
            }

            var flows = flow == EDataFlow.eAll
                ? new[] { EDataFlow.eRender, EDataFlow.eCapture }
                : new[] { flow };

            foreach (var currentFlow in flows)
            {
                if (enumerator.EnumAudioEndpoints(currentFlow, DeviceState.Active, out var collection) != 0)
                    continue;

                try
                {
                    collection.GetCount(out var count);
                    var requiredCapacity = devices.Count + (int)count;
                    if (devices.Capacity < requiredCapacity)
                        devices.Capacity = requiredCapacity;

                    for (uint i = 0; i < count; i++)
                    {
                        if (collection.Item(i, out var device) != 0)
                            continue;

                        try
                        {
                            device.GetId(out var deviceId);

                            var name = deviceId;
                            if (device.OpenPropertyStore(0, out var propertyStore) == 0)
                            {
                                try
                                {
                                    var key = PropertyKey.DeviceFriendlyName;
                                    if (propertyStore.GetValue(ref key, out var variant) == 0)
                                    {
                                        try
                                        {
                                            name = variant.GetString() ?? deviceId;
                                        }
                                        finally
                                        {
                                            _ = PropVariantClear(ref variant);
                                        }
                                    }
                                }
                                finally
                                {
                                    ReleaseComObject(propertyStore);
                                }
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
                        }
                        finally
                        {
                            ReleaseComObject(device);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(collection);
                }
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] 枚举音频设备失败: {ex}");
            throw;
        }
        finally
        {
            ReleaseComObject(enumerator);
        }

        return devices.ToArray();
    }

    [DllImport("ole32.dll", ExactSpelling = true)]
    private static extern int PropVariantClear(ref PropVariant propVariant);

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
            Marshal.ReleaseComObject(comObject);
    }

    private static IMMDeviceEnumerator CreateDeviceEnumerator()
    {
        var comType = Type.GetTypeFromCLSID(MMDeviceEnumeratorClassId)
            ?? throw new InvalidOperationException("无法解析 MMDeviceEnumerator COM 类型。");

        var instance = Activator.CreateInstance(comType)
            ?? throw new InvalidOperationException("无法创建 MMDeviceEnumerator COM 实例。");

        return (IMMDeviceEnumerator)instance;
    }
}
