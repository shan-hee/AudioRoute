using System;

namespace AudioRoute;

public sealed class MixerSessionInfo
{
    public required string SessionKey { get; init; }
    public required string DisplayName { get; init; }
    public required string ActualDeviceSummary { get; init; }
    public required string BoundDeviceSummary { get; init; }
    public string ProcessName { get; init; } = string.Empty;
    public string? ExecutablePath { get; init; }
    public string? BoundDeviceId { get; init; }
    public EDataFlow Flow { get; init; }
    public int ProcessId { get; init; }
    public float Volume { get; init; }
    public bool IsMuted { get; init; }
    public bool IsSystemSession { get; init; }
    public bool CanChangeDevice => ProcessId > 0;
}

public sealed class MixerDeviceChangedEventArgs : EventArgs
{
    public MixerDeviceChangedEventArgs(MixerSessionInfo session, string? deviceId)
    {
        Session = session;
        DeviceId = deviceId;
    }

    public MixerSessionInfo Session { get; }

    public string? DeviceId { get; }
}

public sealed class MixerVolumeChangedEventArgs : EventArgs
{
    public MixerVolumeChangedEventArgs(MixerSessionInfo session, float volume)
    {
        Session = session;
        Volume = volume;
    }

    public MixerSessionInfo Session { get; }

    public float Volume { get; }
}

public sealed class MixerInteractionStateChangedEventArgs : EventArgs
{
    public MixerInteractionStateChangedEventArgs(MixerSessionInfo session, bool isInteracting)
    {
        Session = session;
        IsInteracting = isInteracting;
    }

    public MixerSessionInfo Session { get; }

    public bool IsInteracting { get; }
}
