using Xunit;

namespace AudioRoute.Tests;

public sealed class MasterVolumeStateTests
{
    [Fact]
    public void CreateState_MapsCurrentSystemVolumeAndDeviceMetadata()
    {
        var state = MasterVolumeService.CreateState(
            0.31f,
            isMuted: false,
            deviceId: "device-a",
            deviceName: "测试扬声器");

        Assert.Equal(31, state.Percentage);
        Assert.False(state.IsMuted);
        Assert.Equal(TrayVolumeIconKind.Low, state.IconKind);
        Assert.Equal("device-a", state.DeviceId);
        Assert.Equal("测试扬声器", state.DeviceName);
    }

    [Fact]
    public void CreateState_MuteOverridesVolumeIconKind()
    {
        var state = MasterVolumeService.CreateState(0.78f, isMuted: true);

        Assert.Equal(78, state.Percentage);
        Assert.True(state.IsMuted);
        Assert.Equal(TrayVolumeIconKind.Muted, state.IconKind);
    }

    [Theory]
    [InlineData(0.00f, TrayVolumeIconKind.ZeroBars)]
    [InlineData(0.01f, TrayVolumeIconKind.Low)]
    [InlineData(0.33f, TrayVolumeIconKind.Low)]
    [InlineData(0.34f, TrayVolumeIconKind.Medium)]
    [InlineData(0.66f, TrayVolumeIconKind.Medium)]
    [InlineData(0.67f, TrayVolumeIconKind.High)]
    [InlineData(1.00f, TrayVolumeIconKind.High)]
    public void CreateState_UsesWindowsVolumeIconThresholds(
        float scalar,
        TrayVolumeIconKind expectedIconKind)
    {
        var state = MasterVolumeService.CreateState(scalar, isMuted: false);

        Assert.Equal(expectedIconKind, state.IconKind);
    }
}
