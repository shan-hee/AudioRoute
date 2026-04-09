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

            var percentage = Math.Clamp((int)Math.Round(device.AudioEndpointVolume.MasterVolumeLevelScalar * 100), 0, 100);
            var isMuted = device.AudioEndpointVolume.Mute || percentage == 0;

            return new MasterVolumeState(
                percentage,
                isMuted,
                ResolveIconKind(percentage, isMuted));
        }
        catch
        {
            return null;
        }
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
