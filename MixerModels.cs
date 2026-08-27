using System;

using System.Collections.Generic;
using System.Linq;

namespace AudioRoute;

public sealed record MixerSessionInfo
{
    public required string SessionKey { get; init; }
    public required string DisplayName { get; init; }
    public required string ActualDeviceSummary { get; init; }
    public required string BoundDeviceSummary { get; init; }
    public string ProcessName { get; init; } = string.Empty;
    public string? ExecutablePath { get; init; }
    public string? BoundDeviceId { get; init; }
    public string? VolumeDeviceId { get; init; }
    public string? RoutingUnavailableReason { get; init; }
    public IReadOnlyList<MixerDeviceSessionState> DeviceStates { get; init; } = Array.Empty<MixerDeviceSessionState>();
    public EDataFlow Flow { get; init; }
    public int ProcessId { get; init; }
    public float Volume { get; init; }
    public bool IsMuted { get; init; }
    public bool IsSystemSession { get; init; }
    public bool IsRoutingSupported { get; init; } = true;
    public bool CanChangeDevice => ProcessId > 0 && IsRoutingSupported;

    public bool Equals(MixerSessionInfo? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return string.Equals(SessionKey, other.SessionKey, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(DisplayName, other.DisplayName, StringComparison.Ordinal) &&
            string.Equals(ActualDeviceSummary, other.ActualDeviceSummary, StringComparison.Ordinal) &&
            string.Equals(BoundDeviceSummary, other.BoundDeviceSummary, StringComparison.Ordinal) &&
            string.Equals(ProcessName, other.ProcessName, StringComparison.Ordinal) &&
            string.Equals(ExecutablePath, other.ExecutablePath, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(BoundDeviceId, other.BoundDeviceId, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(VolumeDeviceId, other.VolumeDeviceId, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(RoutingUnavailableReason, other.RoutingUnavailableReason, StringComparison.Ordinal) &&
            DeviceStatesEqual(DeviceStates, other.DeviceStates) &&
            Flow == other.Flow &&
            ProcessId == other.ProcessId &&
            Math.Abs(Volume - other.Volume) < 0.001f &&
            IsMuted == other.IsMuted &&
            IsSystemSession == other.IsSystemSession &&
            IsRoutingSupported == other.IsRoutingSupported;
    }

    public override int GetHashCode()
    {
        return StringComparer.OrdinalIgnoreCase.GetHashCode(SessionKey ?? string.Empty);
    }

    public MixerSessionInfo SelectVolumeDevice(string? deviceId)
    {
        if (string.IsNullOrWhiteSpace(deviceId))
            return this;

        var deviceState = DeviceStates.FirstOrDefault(state =>
            string.Equals(state.DeviceId, deviceId, StringComparison.OrdinalIgnoreCase));

        return deviceState is null
            ? this with { VolumeDeviceId = deviceId }
            : this with
            {
                VolumeDeviceId = deviceState.DeviceId,
                Volume = deviceState.Volume,
                IsMuted = deviceState.IsMuted
            };
    }

    public MixerSessionInfo WithDeviceVolume(string deviceId, string deviceName, float volume, bool isMuted)
    {
        if (string.IsNullOrWhiteSpace(deviceId))
            return this with { Volume = volume, IsMuted = isMuted };

        var updatedStates = new List<MixerDeviceSessionState>(DeviceStates.Count + 1);
        var found = false;

        foreach (var state in DeviceStates)
        {
            if (!string.Equals(state.DeviceId, deviceId, StringComparison.OrdinalIgnoreCase))
            {
                updatedStates.Add(state);
                continue;
            }

            updatedStates.Add(state with { DeviceName = deviceName, Volume = volume, IsMuted = isMuted });
            found = true;
        }

        if (!found)
            updatedStates.Add(new MixerDeviceSessionState(deviceId, deviceName, volume, isMuted));

        var updated = this with { DeviceStates = updatedStates };
        return string.Equals(VolumeDeviceId, deviceId, StringComparison.OrdinalIgnoreCase)
            ? updated with { Volume = volume, IsMuted = isMuted }
            : updated;
    }

    private static bool DeviceStatesEqual(
        IReadOnlyList<MixerDeviceSessionState> left,
        IReadOnlyList<MixerDeviceSessionState> right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left.Count != right.Count)
            return false;

        for (var index = 0; index < left.Count; index++)
        {
            if (left[index] != right[index])
                return false;
        }

        return true;
    }
}

public sealed record MixerDeviceSessionState(
    string DeviceId,
    string DeviceName,
    float Volume,
    bool IsMuted);

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

public sealed class MixerSessionStateChangedEventArgs : EventArgs
{
    public MixerSessionStateChangedEventArgs(MixerSessionInfo session, float volume, bool isMuted)
    {
        Session = session;
        Volume = volume;
        IsMuted = isMuted;
    }

    public MixerSessionInfo Session { get; }

    public float Volume { get; }

    public bool IsMuted { get; }
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
