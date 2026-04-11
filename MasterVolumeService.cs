using System;
using NAudio.CoreAudioApi;

namespace AudioRoute;

public readonly record struct MasterVolumeState(int Percentage, bool IsMuted, TrayVolumeIconKind IconKind);

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

            return CreateState(device.AudioEndpointVolume.MasterVolumeLevelScalar, device.AudioEndpointVolume.Mute);
        }
        catch
        {
            return null;
        }
    }

    public static MasterVolumeState CreateState(float volumeScalar, bool isMuted)
    {
        var percentage = Math.Clamp((int)Math.Round(volumeScalar * 100), 0, 100);

        return new MasterVolumeState(
            percentage,
            isMuted,
            ResolveIconKind(percentage, isMuted));
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
