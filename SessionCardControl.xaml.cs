using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

namespace AudioRoute;

public sealed partial class SessionCardControl : UserControl
{
    private readonly DispatcherQueueTimer volumeCommitTimer;
    private MixerAppSessionInfo appSession;
    private IReadOnlyList<AudioDevice> devices;
    private EDataFlow volumeFlow;
    private bool updatingUi;
    private bool isVolumeInteracting;
    private bool isRoutePickerInteracting;
    private bool isInteracting;
    private bool hasPendingVolumeCommit;
    private float lastCommittedVolume;
    private bool isMuted;
    private string? loadedIconPath;
    private int iconLoadVersion;

    public SessionCardControl(MixerAppSessionInfo appSession, IReadOnlyList<AudioDevice> devices)
    {
        this.appSession = appSession;
        this.devices = devices;
        volumeFlow = appSession.GetAvailableFlow(EDataFlow.eRender);
        InitializeComponent();

        volumeCommitTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        volumeCommitTimer.Interval = TimeSpan.FromMilliseconds(160);
        volumeCommitTimer.Tick += VolumeCommitTimer_Tick;

        ApplySession();
    }

    public event EventHandler<MixerDeviceChangedEventArgs>? DeviceChanged;

    public event EventHandler<MixerVolumeChangedEventArgs>? VolumeChanged;

    public event EventHandler<MixerInteractionStateChangedEventArgs>? InteractionStateChanged;

    public void UpdateSession(MixerAppSessionInfo updatedSession, IReadOnlyList<AudioDevice> updatedDevices)
    {
        appSession = updatedSession;
        devices = updatedDevices;
        volumeFlow = appSession.GetAvailableFlow(volumeFlow);
        ApplySession();
    }

    private MixerSessionInfo? VolumeSession => appSession.GetSession(volumeFlow);

    private void ApplySession()
    {
        TitleTextBlock.Text = appSession.DisplayName;
        ProcessTextBlock.Text = appSession.IsSystemSession
            ? "系统音频会话"
            : $"{appSession.ProcessName} | PID {appSession.ProcessId}";
        SessionFallbackIcon.Glyph = appSession.IsSystemSession ? "\uE7F5" : "\uE77B";

        ToolTipService.SetToolTip(TitleTextBlock, appSession.DisplayName);

        UpdateRoutePicker(OutputRouteButton, OutputRouteTextBlock, appSession.OutputSession, EDataFlow.eRender);
        UpdateRoutePicker(InputRouteButton, InputRouteTextBlock, appSession.InputSession, EDataFlow.eCapture);
        ApplyVolumeSession();
        _ = UpdateSessionIconAsync();
    }

    private void UpdateRoutePicker(Button button, TextBlock textBlock, MixerSessionInfo? session, EDataFlow flow)
    {
        var label = GetFlowLabel(flow);
        var summary = session switch
        {
            null => label,
            _ when session.CanChangeDevice => $"{label} · {session.BoundDeviceSummary}",
            _ => $"{label} · {session.ActualDeviceSummary}"
        };

        button.IsEnabled = session?.CanChangeDevice == true;
        button.Opacity = session is null ? 0.45d : 1d;
        textBlock.Text = summary;
        textBlock.Opacity = button.IsEnabled ? 1d : 0.58d;
        ToolTipService.SetToolTip(button, summary);
    }

    private void ApplyVolumeSession()
    {
        var volumeSession = VolumeSession;
        if (volumeSession is null)
        {
            VolumeSlider.IsEnabled = false;
            MuteButton.IsEnabled = false;
            ToolTipService.SetToolTip(VolumeSlider, null);
            return;
        }

        isMuted = volumeSession.IsMuted;
        VolumeSlider.IsEnabled = true;
        MuteButton.IsEnabled = true;
        UpdateMuteIcon();

        if (!isVolumeInteracting)
        {
            updatingUi = true;
            VolumeSlider.Value = Math.Clamp((int)Math.Round(volumeSession.Volume * 100), 0, 100);
            updatingUi = false;
            lastCommittedVolume = (float)(VolumeSlider.Value / 100d);
            hasPendingVolumeCommit = false;
        }

        UpdateVolumeText();
    }

    private void UpdateMuteIcon()
    {
        if (isMuted || VolumeSlider.Value == 0)
        {
            MuteIcon.Glyph = "\uE74F";
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 150, 150, 150));
        }
        else if (VolumeSlider.Value < 34)
        {
            MuteIcon.Glyph = "\uE993";
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
        else if (VolumeSlider.Value < 67)
        {
            MuteIcon.Glyph = "\uE994";
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
        else
        {
            MuteIcon.Glyph = "\uE767";
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
    }

    private void MuteButton_Click(object sender, RoutedEventArgs e)
    {
        if (VolumeSession is null)
            return;

        if (isMuted)
        {
            isMuted = false;
            if (VolumeSlider.Value == 0)
                VolumeSlider.Value = 100;
        }
        else
        {
            isMuted = true;
            updatingUi = true;
            VolumeSlider.Value = 0;
            updatingUi = false;
            hasPendingVolumeCommit = true;
            CommitVolumeChange();
        }

        UpdateMuteIcon();
    }

    private void OutputRouteButton_Click(object sender, RoutedEventArgs e)
    {
        ShowRoutePicker(OutputRouteButton, appSession.OutputSession, EDataFlow.eRender);
    }

    private void InputRouteButton_Click(object sender, RoutedEventArgs e)
    {
        ShowRoutePicker(InputRouteButton, appSession.InputSession, EDataFlow.eCapture);
    }

    private void ShowRoutePicker(FrameworkElement anchor, MixerSessionInfo? session, EDataFlow flow)
    {
        if (session is null || !session.CanChangeDevice)
            return;

        var flyout = BuildRouteFlyout(session, flow);
        flyout.Opened += RouteFlyout_Opened;
        flyout.Closed += RouteFlyout_Closed;
        flyout.ShowAt(anchor);
    }

    private MenuFlyout BuildRouteFlyout(MixerSessionInfo session, EDataFlow flow)
    {
        var flyout = new MenuFlyout();
        flyout.Items.Add(CreateRouteMenuItem(session, null, "跟随系统默认", string.IsNullOrWhiteSpace(session.BoundDeviceId)));

        foreach (var device in devices.Where(device => device.Flow == flow))
        {
            flyout.Items.Add(CreateRouteMenuItem(
                session,
                device.Id,
                device.Name,
                string.Equals(device.Id, session.BoundDeviceId, StringComparison.OrdinalIgnoreCase)));
        }

        if (!string.IsNullOrWhiteSpace(session.BoundDeviceId) &&
            devices.Where(device => device.Flow == flow)
                .All(device => !string.Equals(device.Id, session.BoundDeviceId, StringComparison.OrdinalIgnoreCase)))
        {
            flyout.Items.Add(CreateRouteMenuItem(session, session.BoundDeviceId, session.BoundDeviceSummary, true));
        }

        return flyout;
    }

    private MenuFlyoutItem CreateRouteMenuItem(MixerSessionInfo session, string? deviceId, string title, bool selected)
    {
        var item = new MenuFlyoutItem
        {
            Text = title
        };

        if (selected)
            item.Icon = new FontIcon { Glyph = "\uE73E" };

        item.Click += (_, _) => DeviceChanged?.Invoke(this, new MixerDeviceChangedEventArgs(session, deviceId));
        return item;
    }

    private void RouteFlyout_Opened(object sender, object e)
    {
        isRoutePickerInteracting = true;
        UpdateInteractionState();
    }

    private void RouteFlyout_Closed(object sender, object e)
    {
        isRoutePickerInteracting = false;
        UpdateInteractionState();
    }

    private void VolumeSlider_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        isVolumeInteracting = true;
        UpdateInteractionState();
    }

    private void VolumeSlider_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        isVolumeInteracting = false;
        UpdateInteractionState();
        CommitVolumeChange();
    }

    private void VolumeSlider_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        isVolumeInteracting = false;
        UpdateInteractionState();
        CommitVolumeChange();
    }

    private void VolumeSlider_ValueChanged(object sender, Microsoft.UI.Xaml.Controls.Primitives.RangeBaseValueChangedEventArgs e)
    {
        UpdateVolumeText();
        UpdateMuteIcon();

        if (updatingUi || VolumeSession is null)
            return;

        hasPendingVolumeCommit = true;
        volumeCommitTimer.Stop();
        volumeCommitTimer.Start();
    }

    private void UpdateVolumeText()
    {
        var value = Math.Clamp((int)Math.Round(VolumeSlider.Value), 0, 100);
        ToolTipService.SetToolTip(VolumeSlider, $"{value}%");
    }

    private async System.Threading.Tasks.Task UpdateSessionIconAsync()
    {
        var executablePath = appSession.ExecutablePath;
        var requestVersion = ++iconLoadVersion;

        if (appSession.IsSystemSession || string.IsNullOrWhiteSpace(executablePath))
        {
            loadedIconPath = null;
            ShowFallbackIcon(appSession.IsSystemSession, clearImageSource: true);
            return;
        }

        if (string.Equals(loadedIconPath, executablePath, StringComparison.OrdinalIgnoreCase) &&
            SessionIconImage.Source is not null)
        {
            ShowLoadedIcon();
            return;
        }

        ShowFallbackIcon(appSession.IsSystemSession, clearImageSource: true);

        var iconSource = await AppIconService.TryLoadIconAsync(executablePath);
        if (requestVersion != iconLoadVersion)
            return;

        if (iconSource is null)
        {
            loadedIconPath = null;
            ShowFallbackIcon(appSession.IsSystemSession, clearImageSource: true);
            return;
        }

        loadedIconPath = executablePath;
        SessionIconImage.Source = iconSource;
        ShowLoadedIcon();
    }

    private void ShowFallbackIcon(bool isSystemSession, bool clearImageSource)
    {
        SessionFallbackIcon.Glyph = isSystemSession ? "\uE7F5" : "\uE77B";
        SessionFallbackIcon.Visibility = Visibility.Visible;
        SessionIconImage.Visibility = Visibility.Collapsed;
        if (clearImageSource)
            SessionIconImage.Source = null;
    }

    private void ShowLoadedIcon()
    {
        SessionFallbackIcon.Visibility = Visibility.Collapsed;
        SessionIconImage.Visibility = Visibility.Visible;
    }

    private void VolumeCommitTimer_Tick(object? sender, object e)
    {
        CommitVolumeChange();
    }

    private void CommitVolumeChange()
    {
        volumeCommitTimer.Stop();

        if (!hasPendingVolumeCommit)
            return;

        var volumeSession = VolumeSession;
        if (volumeSession is null)
        {
            hasPendingVolumeCommit = false;
            return;
        }

        var volume = (float)(VolumeSlider.Value / 100d);
        if (Math.Abs(volume - lastCommittedVolume) < 0.005f)
        {
            hasPendingVolumeCommit = false;
            return;
        }

        hasPendingVolumeCommit = false;
        lastCommittedVolume = volume;
        VolumeChanged?.Invoke(this, new MixerVolumeChangedEventArgs(volumeSession, volume));
    }

    private void UpdateInteractionState()
    {
        var nextState = isVolumeInteracting || isRoutePickerInteracting;
        if (nextState == isInteracting)
            return;

        isInteracting = nextState;
        InteractionStateChanged?.Invoke(this, new MixerInteractionStateChangedEventArgs(VolumeSession ?? appSession.PrimarySession, isInteracting));
    }

    private static string GetFlowLabel(EDataFlow flow)
    {
        return flow == EDataFlow.eRender ? "输出" : "输入";
    }
}
