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
    private sealed record PendingSessionVolumeCommit(MixerVolumeChangedEventArgs Change, SessionCardControl? SourceCard);
    private sealed record SessionVolumeCommitFailure(PendingSessionVolumeCommit Commit, Exception Exception);
    private sealed record SessionVolumeCommitBatchResult(
        IReadOnlyList<PendingSessionVolumeCommit> Succeeded,
        IReadOnlyList<SessionVolumeCommitFailure> Failures);
}
