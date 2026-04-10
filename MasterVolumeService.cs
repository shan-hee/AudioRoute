using System;
using NAudio.CoreAudioApi;

namespace AudioRoute;

public readonly record struct MasterVolumeState(int Percentage, bool IsMuted, TrayVolumeIconKind IconKind);

public enum TrayVolumeIconKind
{
    Muted,
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
        var effectiveMuted = isMuted || percentage == 0;

        return new MasterVolumeState(
            percentage,
            effectiveMuted,
            ResolveIconKind(percentage, effectiveMuted));
    }

    private static TrayVolumeIconKind ResolveIconKind(int percentage, bool isMuted)
    {
        if (isMuted || percentage == 0)
            return TrayVolumeIconKind.Muted;

        if (percentage < 34)
            return TrayVolumeIconKind.Low;

        if (percentage < 67)
            return TrayVolumeIconKind.Medium;

        return TrayVolumeIconKind.High;
    }
}
