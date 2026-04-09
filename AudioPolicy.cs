using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace AudioRoute;

public enum EDataFlow
{
    eRender = 0,
    eCapture = 1,
    eAll = 2
}

public enum ERole
{
    eConsole = 0,
    eMultimedia = 1,
    eCommunications = 2
}

public static class AudioPolicyManager
{
    private const string AudioPolicyConfigRuntimeClass = "Windows.Media.Internal.AudioPolicyConfig";
    private const string RenderInterfaceId = "{e6327cad-dcec-4949-ae8a-991e976a79d2}";
    private const string CaptureInterfaceId = "{2eef81be-33fa-4800-9670-1cd474972c3f}";
    private const int NewInterfaceBuild = 21390;
    private const int SetPersistedDefaultAudioEndpointVtableIndex = 25;
    private const int GetPersistedDefaultAudioEndpointVtableIndex = 26;
    private const int HResultElementNotFound = unchecked((int)0x80070490);
    private const int HResultChangedMode = unchecked((int)0x80010106);
    private const uint CoinitApartmentThreaded = 0x2;

    private static readonly object SyncRoot = new();
    private static readonly Dictionary<PolicyCacheKey, CachedPolicyValue> DefaultDeviceCache = new();
    private static readonly TimeSpan DefaultDeviceCacheDuration = TimeSpan.FromSeconds(3);

    private static bool s_comInitialized;
    private static IntPtr s_factory;
    private static GetPersistedDefaultAudioEndpointDelegate? s_getPersistedDefaultAudioEndpoint;
    private static SetPersistedDefaultAudioEndpointDelegate? s_setPersistedDefaultAudioEndpoint;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetPersistedDefaultAudioEndpointDelegate(
        IntPtr @this,
        uint processId,
        int flow,
        int role,
        ref IntPtr deviceId);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int SetPersistedDefaultAudioEndpointDelegate(
        IntPtr @this,
        uint processId,
        int flow,
        int role,
        IntPtr deviceId);

    public static string? GetAppDefaultDevice(uint processId, EDataFlow flow = EDataFlow.eRender, ERole role = ERole.eMultimedia)
    {
        EnsureFlowSupported(flow);
        var cacheKey = new PolicyCacheKey(processId, flow, role);
        var now = DateTimeOffset.UtcNow;

        lock (SyncRoot)
        {
            if (DefaultDeviceCache.TryGetValue(cacheKey, out var cachedValue) &&
                cachedValue.ExpiresAt > now)
            {
                return cachedValue.DeviceId;
            }
        }

        EnsureFactory();

        foreach (var currentRole in ExpandRoles(role))
        {
            var packedDeviceId = GetPackedDeviceId(processId, flow, currentRole);
            if (!string.IsNullOrWhiteSpace(packedDeviceId))
            {
                var deviceId = UnpackDeviceId(packedDeviceId, flow);
                StoreDefaultDeviceCache(cacheKey, deviceId, now);
                return deviceId;
            }
        }

        StoreDefaultDeviceCache(cacheKey, null, now);
        return null;
    }

    public static void SetAppDefaultDevice(uint processId, string deviceId, EDataFlow flow = EDataFlow.eRender, ERole role = ERole.eMultimedia)
    {
        EnsureFlowSupported(flow);
        if (string.IsNullOrWhiteSpace(deviceId))
            throw new ArgumentException("Device ID cannot be empty.", nameof(deviceId));

        EnsureFactory();

        var packedDeviceId = PackDeviceId(deviceId, flow);
        foreach (var currentRole in ExpandRoles(role))
            SetPackedDeviceId(processId, flow, currentRole, packedDeviceId);

        InvalidateDefaultDeviceCache(processId, flow);
    }

    public static void ClearAppDefaultDevice(uint processId, EDataFlow flow = EDataFlow.eRender, ERole role = ERole.eMultimedia)
    {
        EnsureFlowSupported(flow);
        EnsureFactory();

        foreach (var currentRole in ExpandRoles(role))
            ClearPackedDeviceId(processId, flow, currentRole);

        InvalidateDefaultDeviceCache(processId, flow);
    }

    public static void Cleanup()
    {
        lock (SyncRoot)
        {
            if (s_factory != IntPtr.Zero)
            {
                Marshal.Release(s_factory);
                s_factory = IntPtr.Zero;
            }

            s_getPersistedDefaultAudioEndpoint = null;
            s_setPersistedDefaultAudioEndpoint = null;
            DefaultDeviceCache.Clear();

            if (s_comInitialized)
            {
                CoUninitialize();
                s_comInitialized = false;
            }
        }
    }

    private static void EnsureFlowSupported(EDataFlow flow)
    {
        if (flow != EDataFlow.eRender && flow != EDataFlow.eCapture)
            throw new NotSupportedException($"Unsupported flow: {flow}");
    }

    private static ERole[] ExpandRoles(ERole preferredRole)
    {
        return preferredRole switch
        {
            ERole.eConsole => [ERole.eConsole],
            ERole.eCommunications => [ERole.eCommunications],
            _ => [ERole.eMultimedia, ERole.eConsole, ERole.eCommunications]
        };
    }

    private static void EnsureFactory()
    {
        if (s_factory != IntPtr.Zero)
            return;

        lock (SyncRoot)
        {
            if (s_factory != IntPtr.Zero)
                return;

            EnsureComInitialized();

            var activationFactoryClassId = ResolveActivationFactoryClassId();
            var runtimeClass = CreateHString(AudioPolicyConfigRuntimeClass);

            try
            {
                var hr = RoGetActivationFactory(runtimeClass, ref activationFactoryClassId, out s_factory);
                ThrowIfFailed(hr, "RoGetActivationFactory");

                s_getPersistedDefaultAudioEndpoint =
                    GetComMethod<GetPersistedDefaultAudioEndpointDelegate>(s_factory, GetPersistedDefaultAudioEndpointVtableIndex);
                s_setPersistedDefaultAudioEndpoint =
                    GetComMethod<SetPersistedDefaultAudioEndpointDelegate>(s_factory, SetPersistedDefaultAudioEndpointVtableIndex);
            }
            finally
            {
                DeleteHString(runtimeClass);
            }
        }
    }

    private static string? GetPackedDeviceId(uint processId, EDataFlow flow, ERole role)
    {
        if (s_getPersistedDefaultAudioEndpoint == null)
            throw new InvalidOperationException("Audio policy factory is not initialized.");

        var packedDeviceIdHString = IntPtr.Zero;
        var hr = s_getPersistedDefaultAudioEndpoint(s_factory, processId, (int)flow, (int)role, ref packedDeviceIdHString);
        if (hr == HResultElementNotFound)
            return null;

        ThrowIfFailed(hr, $"GetPersistedDefaultAudioEndpoint(pid={processId}, flow={flow}, role={role})");

        if (packedDeviceIdHString == IntPtr.Zero)
            return null;

        try
        {
            return HStringToString(packedDeviceIdHString);
        }
        finally
        {
            DeleteHString(packedDeviceIdHString);
        }
    }

    private static void SetPackedDeviceId(uint processId, EDataFlow flow, ERole role, string packedDeviceId)
    {
        if (s_setPersistedDefaultAudioEndpoint == null)
            throw new InvalidOperationException("Audio policy factory is not initialized.");

        var packedDeviceIdHString = CreateHString(packedDeviceId);
        try
        {
            var hr = s_setPersistedDefaultAudioEndpoint(s_factory, processId, (int)flow, (int)role, packedDeviceIdHString);
            ThrowIfFailed(hr, $"SetPersistedDefaultAudioEndpoint(pid={processId}, flow={flow}, role={role})");
        }
        finally
        {
            DeleteHString(packedDeviceIdHString);
        }
    }

    private static void ClearPackedDeviceId(uint processId, EDataFlow flow, ERole role)
    {
        if (s_setPersistedDefaultAudioEndpoint == null)
            throw new InvalidOperationException("Audio policy factory is not initialized.");

        var hr = s_setPersistedDefaultAudioEndpoint(s_factory, processId, (int)flow, (int)role, IntPtr.Zero);
        ThrowIfFailed(hr, $"ClearPersistedDefaultAudioEndpoint(pid={processId}, flow={flow}, role={role})");
    }

    private static string PackDeviceId(string deviceId, EDataFlow flow)
    {
        var interfaceId = flow == EDataFlow.eCapture ? CaptureInterfaceId : RenderInterfaceId;
        return $@"\\?\SWD#MMDEVAPI#{deviceId}#{interfaceId}";
    }

    private static string UnpackDeviceId(string packedDeviceId, EDataFlow flow)
    {
        var interfaceId = flow == EDataFlow.eCapture ? CaptureInterfaceId : RenderInterfaceId;
        var prefix = @"\\?\SWD#MMDEVAPI#";

        if (!packedDeviceId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            return packedDeviceId;

        if (!packedDeviceId.EndsWith(interfaceId, StringComparison.OrdinalIgnoreCase))
            return packedDeviceId;

        var payloadLength = packedDeviceId.Length - prefix.Length - interfaceId.Length - 1;
        if (payloadLength <= 0)
            return packedDeviceId;

        return packedDeviceId.Substring(prefix.Length, payloadLength);
    }

    private static void StoreDefaultDeviceCache(PolicyCacheKey cacheKey, string? deviceId, DateTimeOffset now)
    {
        lock (SyncRoot)
        {
            DefaultDeviceCache[cacheKey] = new CachedPolicyValue(deviceId, now + DefaultDeviceCacheDuration);

            if (DefaultDeviceCache.Count <= 256)
                return;

            var expiredKeys = DefaultDeviceCache
                .Where(pair => pair.Value.ExpiresAt <= now)
                .Select(pair => pair.Key)
                .ToArray();

            foreach (var key in expiredKeys)
                DefaultDeviceCache.Remove(key);

            if (DefaultDeviceCache.Count <= 256)
                return;

            using var enumerator = DefaultDeviceCache.Keys.GetEnumerator();
            if (enumerator.MoveNext())
                DefaultDeviceCache.Remove(enumerator.Current);
        }
    }

    private static void InvalidateDefaultDeviceCache(uint processId, EDataFlow flow)
    {
        lock (SyncRoot)
        {
            var keysToRemove = DefaultDeviceCache.Keys
                .Where(key => key.ProcessId == processId && key.Flow == flow)
                .ToArray();

            foreach (var key in keysToRemove)
                DefaultDeviceCache.Remove(key);
        }
    }

    private static Guid ResolveActivationFactoryClassId()
    {
        var buildNumber = GetWindowsBuildNumber();
        var interfaceId = buildNumber >= NewInterfaceBuild
            ? "ab3d4648-e242-459f-b02f-541c70306324"
            : "2a59116d-6c4f-45e0-a74f-707e3fef9258";

        return new Guid(interfaceId);
    }

    private static int GetWindowsBuildNumber()
    {
        var versionInfo = new OSVersionInfo
        {
            OSVersionInfoSize = Marshal.SizeOf<OSVersionInfo>()
        };

        var status = RtlGetVersion(ref versionInfo);
        if (status != 0)
            throw new InvalidOperationException($"RtlGetVersion failed with NTSTATUS 0x{status:X8}.");

        return versionInfo.BuildNumber;
    }

    private static T GetComMethod<T>(IntPtr comObject, int vtableIndex) where T : Delegate
    {
        var vtable = Marshal.ReadIntPtr(comObject);
        var methodPtr = Marshal.ReadIntPtr(vtable, vtableIndex * IntPtr.Size);
        return Marshal.GetDelegateForFunctionPointer<T>(methodPtr);
    }

    private static void EnsureComInitialized()
    {
        if (s_comInitialized)
            return;

        var hr = CoInitializeEx(IntPtr.Zero, CoinitApartmentThreaded);
        if (hr == HResultChangedMode)
            return;

        ThrowIfFailed(hr, "CoInitializeEx");
        s_comInitialized = true;
    }

    private static IntPtr CreateHString(string value)
    {
        var hr = WindowsCreateString(value, value.Length, out var hstring);
        ThrowIfFailed(hr, "WindowsCreateString");
        return hstring;
    }

    private static void DeleteHString(IntPtr hstring)
    {
        if (hstring != IntPtr.Zero)
            WindowsDeleteString(hstring);
    }

    private static string HStringToString(IntPtr hstring)
    {
        var rawBuffer = WindowsGetStringRawBuffer(hstring, out var length);
        return length == 0 ? string.Empty : Marshal.PtrToStringUni(rawBuffer, (int)length) ?? string.Empty;
    }

    private static void ThrowIfFailed(int hr, string operation)
    {
        if (hr < 0)
            throw new InvalidOperationException($"{operation} failed with HRESULT 0x{hr & 0xFFFFFFFF:X8}.");
    }

    [DllImport("combase.dll", ExactSpelling = true)]
    private static extern int RoGetActivationFactory(IntPtr activatableClassId, ref Guid iid, out IntPtr factory);

    [DllImport("combase.dll", ExactSpelling = true)]
    private static extern int WindowsCreateString(
        [MarshalAs(UnmanagedType.LPWStr)] string sourceString,
        int length,
        out IntPtr hstring);

    [DllImport("combase.dll", ExactSpelling = true)]
    private static extern int WindowsDeleteString(IntPtr hstring);

    [DllImport("combase.dll", ExactSpelling = true)]
    private static extern IntPtr WindowsGetStringRawBuffer(IntPtr hstring, out uint length);

    [DllImport("ntdll.dll", ExactSpelling = true)]
    private static extern int RtlGetVersion(ref OSVersionInfo versionInfo);

    [DllImport("ole32.dll", ExactSpelling = true)]
    private static extern int CoInitializeEx(IntPtr reserved, uint coInit);

    [DllImport("ole32.dll", ExactSpelling = true)]
    private static extern void CoUninitialize();

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct OSVersionInfo
    {
        public int OSVersionInfoSize;
        public int MajorVersion;
        public int MinorVersion;
        public int BuildNumber;
        public int PlatformId;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string CsdVersion;
    }

    private readonly record struct PolicyCacheKey(uint ProcessId, EDataFlow Flow, ERole Role);

    private sealed record CachedPolicyValue(string? DeviceId, DateTimeOffset ExpiresAt);
}
