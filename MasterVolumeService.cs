using System;
using NAudio.CoreAudioApi;

namespace AudioRoute;

public readonly record struct MasterVolumeState(
    int Percentage,
    bool IsMuted,
    TrayVolumeIconKind IconKind,
    string? DeviceId,
    string? DeviceName);

public enum TrayVolumeIconKind
{
    NoDevice,
    Muted,
    ZeroBars,
    Low,
    Medium,
    High
}

public static class MasterVolumeService
{
    public static MasterVolumeState? TryGetMasterVolumeState()
    {
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            using var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

            return CreateState(
                device.AudioEndpointVolume.MasterVolumeLevelScalar,
                device.AudioEndpointVolume.Mute,
                device.ID,
                device.FriendlyName);
        }
        catch
        {
            return null;
        }
    }

    public static MasterVolumeState? TryAdjustMasterVolume(int deltaPercentage)
    {
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            using var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var endpointVolume = device.AudioEndpointVolume;
            var currentPercentage = Math.Clamp(
                (int)Math.Round(endpointVolume.MasterVolumeLevelScalar * 100),
                0,
                100);
            var nextPercentage = Math.Clamp(currentPercentage + deltaPercentage, 0, 100);
            endpointVolume.MasterVolumeLevelScalar = nextPercentage / 100f;

            return CreateState(
                endpointVolume.MasterVolumeLevelScalar,
                endpointVolume.Mute,
                device.ID,
                device.FriendlyName);
        }
        catch
        {
            return null;
        }
    }

    public static MasterVolumeState? TryToggleMasterMute()
    {
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            using var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var endpointVolume = device.AudioEndpointVolume;
            endpointVolume.Mute = !endpointVolume.Mute;

            return CreateState(
                endpointVolume.MasterVolumeLevelScalar,
                endpointVolume.Mute,
                device.ID,
                device.FriendlyName);
        }
        catch
        {
            return null;
        }
    }

    public static MasterVolumeState CreateState(
        float volumeScalar,
        bool isMuted,
        string? deviceId = null,
        string? deviceName = null)
    {
        var percentage = Math.Clamp((int)Math.Round(volumeScalar * 100), 0, 100);

        return new MasterVolumeState(
            percentage,
            isMuted,
            ResolveIconKind(percentage, isMuted),
            deviceId,
            deviceName);
    }

    private static TrayVolumeIconKind ResolveIconKind(int percentage, bool isMuted)
    {
        if (isMuted)
            return TrayVolumeIconKind.Muted;

        if (percentage == 0)
            return TrayVolumeIconKind.ZeroBars;

        if (percentage < 34)
            return TrayVolumeIconKind.Low;

        if (percentage < 67)
            return TrayVolumeIconKind.Medium;

        return TrayVolumeIconKind.High;
    }
}
