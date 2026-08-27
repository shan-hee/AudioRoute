using System;
using System.Collections.Generic;

namespace AudioRoute;

internal enum TrayDockEdge
{
    Bottom,
    Left,
    Right,
    Top
}

public sealed partial class MainWindow
{
    [Flags]
    private enum RefreshSessionScope
    {
        None = 0,
        Render = 1,
        Capture = 2,
        All = Render | Capture
    }

    private sealed record PanelSnapshot(IReadOnlyList<AudioDevice> Devices, IReadOnlyList<MixerAppSessionInfo> Sessions);
    private sealed record PendingSessionStateCommit(MixerSessionStateChangedEventArgs Change, SessionCardControl? SourceCard);
    private sealed record PendingRoutedSessionStateRestore(
        string SessionKey,
        EDataFlow Flow,
        string DeviceId,
        float Volume,
        bool IsMuted);
    private sealed record SessionStateCommitFailure(PendingSessionStateCommit Commit, Exception Exception);
    private sealed record SessionStateCommitBatchResult(
        IReadOnlyList<PendingSessionStateCommit> Succeeded,
        IReadOnlyList<PendingSessionStateCommit> Deferred,
        IReadOnlyList<SessionStateCommitFailure> Failures);
}
