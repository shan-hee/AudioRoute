using Xunit;

namespace AudioRoute.Tests;

public sealed class MixerSessionInfoTests
{
    [Fact]
    public void SelectVolumeDevice_RestoresRememberedVolumeAndMuteState()
    {
        var session = CreateSession(
            volumeDeviceId: "speaker-a",
            volume: 0.25f,
            isMuted: false,
            new MixerDeviceSessionState("speaker-a", "扬声器 A", 0.25f, false),
            new MixerDeviceSessionState("speaker-b", "扬声器 B", 0.72f, true));

        var selected = session.SelectVolumeDevice("speaker-b");

        Assert.Equal("speaker-b", selected.VolumeDeviceId);
        Assert.Equal(0.72f, selected.Volume);
        Assert.True(selected.IsMuted);
    }

    [Fact]
    public void SelectVolumeDevice_UnknownDeviceChangesSelectionWithoutInventingState()
    {
        var session = CreateSession(
            volumeDeviceId: "speaker-a",
            volume: 0.25f,
            isMuted: false,
            new MixerDeviceSessionState("speaker-a", "扬声器 A", 0.25f, false));

        var selected = session.SelectVolumeDevice("speaker-new");

        Assert.Equal("speaker-new", selected.VolumeDeviceId);
        Assert.Equal(0.25f, selected.Volume);
        Assert.False(selected.IsMuted);
        Assert.Single(selected.DeviceStates);
    }

    [Fact]
    public void WithDeviceVolume_UpdatesOnlyTheRequestedDevice()
    {
        var session = CreateSession(
            volumeDeviceId: "speaker-b",
            volume: 0.72f,
            isMuted: false,
            new MixerDeviceSessionState("speaker-a", "扬声器 A", 0.25f, false),
            new MixerDeviceSessionState("speaker-b", "扬声器 B", 0.72f, false));

        var updated = session.WithDeviceVolume("SPEAKER-A", "扬声器 A", 0.41f, true);

        var stateA = Assert.Single(updated.DeviceStates, state => state.DeviceId == "speaker-a");
        var stateB = Assert.Single(updated.DeviceStates, state => state.DeviceId == "speaker-b");
        Assert.Equal(0.41f, stateA.Volume);
        Assert.True(stateA.IsMuted);
        Assert.Equal(0.72f, stateB.Volume);
        Assert.False(stateB.IsMuted);
        Assert.Equal(0.72f, updated.Volume);
        Assert.False(updated.IsMuted);
    }

    [Fact]
    public void WithDeviceVolume_DelayedOldDeviceEventDoesNotOverwriteVisibleDevice()
    {
        var session = CreateSession(
            volumeDeviceId: "speaker-b",
            volume: 0.72f,
            isMuted: true,
            new MixerDeviceSessionState("speaker-a", "扬声器 A", 0.25f, false),
            new MixerDeviceSessionState("speaker-b", "扬声器 B", 0.72f, true));

        var updated = session.WithDeviceVolume("speaker-a", "扬声器 A", 0.12f, false);

        Assert.Equal("speaker-b", updated.VolumeDeviceId);
        Assert.Equal(0.72f, updated.Volume);
        Assert.True(updated.IsMuted);
        Assert.Equal(0.12f, Assert.Single(updated.DeviceStates, state => state.DeviceId == "speaker-a").Volume);
    }

    [Fact]
    public void WithDeviceVolume_PreservesNonZeroVolumeWhileMuted()
    {
        var session = CreateSession(
            volumeDeviceId: "speaker-a",
            volume: 0.63f,
            isMuted: false,
            new MixerDeviceSessionState("speaker-a", "扬声器 A", 0.63f, false));

        var muted = session.WithDeviceVolume("speaker-a", "扬声器 A", 0.63f, true);
        var unmuted = muted.WithDeviceVolume("speaker-a", "扬声器 A", 0.63f, false);

        Assert.Equal(0.63f, muted.Volume);
        Assert.True(muted.IsMuted);
        Assert.Equal(0.63f, unmuted.Volume);
        Assert.False(unmuted.IsMuted);
    }

    private static MixerSessionInfo CreateSession(
        string volumeDeviceId,
        float volume,
        bool isMuted,
        params MixerDeviceSessionState[] states)
    {
        return new MixerSessionInfo
        {
            SessionKey = "pid:42",
            DisplayName = "测试应用",
            ActualDeviceSummary = "测试设备",
            BoundDeviceSummary = "测试设备",
            ProcessName = "test",
            VolumeDeviceId = volumeDeviceId,
            Flow = EDataFlow.eRender,
            ProcessId = 42,
            Volume = volume,
            IsMuted = isMuted,
            DeviceStates = states
        };
    }
}
