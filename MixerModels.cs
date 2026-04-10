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
    public string? RoutingUnavailableReason { get; init; }
    public EDataFlow Flow { get; init; }
    public int ProcessId { get; init; }
    public float Volume { get; init; }
    public bool IsMuted { get; init; }
    public bool IsSystemSession { get; init; }
    public bool IsRoutingSupported { get; init; } = true;
    public bool CanChangeDevice => ProcessId > 0 && IsRoutingSupported;
}

public sealed class MixerAppSessionInfo
{
    public required string SessionKey { get; init; }
    public required MixerSessionInfo PrimarySession { get; init; }
    public MixerSessionInfo? OutputSession { get; init; }
    public MixerSessionInfo? InputSession { get; init; }

    public string DisplayName => PrimarySession.DisplayName;

    public string ProcessName => PrimarySession.ProcessName;

    public string? ExecutablePath => PrimarySession.ExecutablePath;

    public int ProcessId => PrimarySession.ProcessId;

    public bool IsSystemSession => PrimarySession.IsSystemSession;

    public bool HasOutput => OutputSession is not null;

    public bool HasInput => InputSession is not null;

    public MixerSessionInfo? GetSession(EDataFlow flow)
    {
        return flow switch
        {
            EDataFlow.eRender => OutputSession,
            EDataFlow.eCapture => InputSession,
            _ => null
        };
    }

    public EDataFlow GetAvailableFlow(EDataFlow preferredFlow)
    {
        if (preferredFlow == EDataFlow.eRender && HasOutput)
            return EDataFlow.eRender;

        if (preferredFlow == EDataFlow.eCapture && HasInput)
            return EDataFlow.eCapture;

        return HasOutput ? EDataFlow.eRender : EDataFlow.eCapture;
    }
}

public sealed class MixerDeviceChangedEventArgs : EventArgs
{
    public MixerDeviceChangedEventArgs(MixerSessionInfo session, string? deviceId, string selectedDeviceSummary)
    {
        Session = session;
        DeviceId = deviceId;
        SelectedDeviceSummary = selectedDeviceSummary;
    }

    public MixerSessionInfo Session { get; }

    public string? DeviceId { get; }

    public string SelectedDeviceSummary { get; }
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
