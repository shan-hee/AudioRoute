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
    private MixerSessionInfo session;
    private bool updatingUi;
    private bool isVolumeInteracting;
    private bool isComboInteracting;
    private bool isInteracting;
    private bool hasPendingVolumeCommit;
    private float lastCommittedVolume;
    private bool isMuted;

    public SessionCardControl(MixerSessionInfo session, IReadOnlyList<AudioDevice> devices)
    {
        this.session = session;
        InitializeComponent();

        volumeCommitTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        volumeCommitTimer.Interval = TimeSpan.FromMilliseconds(160);
        volumeCommitTimer.Tick += VolumeCommitTimer_Tick;

        ApplySession(devices);
    }

    public event EventHandler<MixerDeviceChangedEventArgs>? DeviceChanged;

    public event EventHandler<MixerVolumeChangedEventArgs>? VolumeChanged;

    public event EventHandler<MixerInteractionStateChangedEventArgs>? InteractionStateChanged;

    public void UpdateSession(MixerSessionInfo updatedSession, IReadOnlyList<AudioDevice> devices)
    {
        session = updatedSession;
        ApplySession(devices);
    }

    private void ApplySession(IReadOnlyList<AudioDevice> devices)
    {
        TitleTextBlock.Text = session.DisplayName;
        ProcessTextBlock.Text = session.IsSystemSession
            ? "System audio session"
            : $"{session.ProcessName} | PID {session.ProcessId}";
        SessionIcon.Glyph = session.IsSystemSession ? "\uE7F5" : "\uE77B";

        ToolTipService.SetToolTip(TitleTextBlock, session.DisplayName);

        // Device info - compact single line
        var deviceInfo = session.CanChangeDevice
            ? session.BoundDeviceSummary
            : session.ActualDeviceSummary;
        DeviceInfoTextBlock.Text = deviceInfo;
        ToolTipService.SetToolTip(DeviceInfoTextBlock, $"Current: {session.ActualDeviceSummary}\nTarget: {(session.CanChangeDevice ? session.BoundDeviceSummary : "N/A")}");

        // Device panel visibility
        DevicePanel.Visibility = session.CanChangeDevice ? Visibility.Visible : Visibility.Collapsed;

        DeviceComboBox.IsEnabled = session.CanChangeDevice;
        PopulateDeviceChoices(devices);

        // Mute state
        isMuted = session.IsMuted;
        UpdateMuteIcon();

        if (!isVolumeInteracting)
        {
            updatingUi = true;
            VolumeSlider.Value = Math.Clamp((int)Math.Round(session.Volume * 100), 0, 100);
            updatingUi = false;
            lastCommittedVolume = (float)(VolumeSlider.Value / 100d);
        }

        UpdateVolumeText();
    }

    private void UpdateMuteIcon()
    {
        if (isMuted || VolumeSlider.Value == 0)
        {
            MuteIcon.Glyph = "\uE74F"; // Muted
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 150, 150, 150));
        }
        else if (VolumeSlider.Value < 34)
        {
            MuteIcon.Glyph = "\uE993"; // Low volume
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
        else if (VolumeSlider.Value < 67)
        {
            MuteIcon.Glyph = "\uE994"; // Medium volume
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
        else
        {
            MuteIcon.Glyph = "\uE767"; // Full volume
            MuteIcon.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
                Windows.UI.Color.FromArgb(255, 255, 255, 255));
        }
    }

    private void MuteButton_Click(object sender, RoutedEventArgs e)
    {
        if (isMuted)
        {
            // Unmute: restore volume to previous value or 100%
            isMuted = false;
            if (VolumeSlider.Value == 0)
            {
                VolumeSlider.Value = 100;
            }
        }
        else
        {
            // Mute: set volume to 0
            isMuted = true;
            updatingUi = true;
            VolumeSlider.Value = 0;
            updatingUi = false;
            hasPendingVolumeCommit = true;
            CommitVolumeChange();
        }
        UpdateMuteIcon();
    }

    private void PopulateDeviceChoices(IReadOnlyList<AudioDevice> devices)
    {
        updatingUi = true;
        try
        {
            DeviceComboBox.Items.Clear();

            if (session.CanChangeDevice)
            {
                DeviceComboBox.Items.Add(new DeviceChoice(null, "Follow system default"));

                foreach (var device in devices.Where(device => device.Flow == session.Flow))
                {
                    var title = device.IsDefault ? $"{device.Name} [default]" : device.Name;
                    DeviceComboBox.Items.Add(new DeviceChoice(device.Id, title));
                }

                if (!string.IsNullOrWhiteSpace(session.BoundDeviceId) &&
                    DeviceComboBox.Items.OfType<DeviceChoice>().All(item => !string.Equals(item.DeviceId, session.BoundDeviceId, StringComparison.OrdinalIgnoreCase)))
                {
                    DeviceComboBox.Items.Add(new DeviceChoice(session.BoundDeviceId, session.BoundDeviceSummary));
                }

                var selected = DeviceComboBox.Items
                    .OfType<DeviceChoice>()
                    .FirstOrDefault(item => string.Equals(item.DeviceId, session.BoundDeviceId, StringComparison.OrdinalIgnoreCase))
                    ?? DeviceComboBox.Items.OfType<DeviceChoice>().FirstOrDefault(item => item.DeviceId == null);

                DeviceComboBox.SelectedItem = selected;
            }
            else
            {
                DeviceComboBox.Items.Add(new DeviceChoice(null, "System session"));
                DeviceComboBox.SelectedIndex = 0;
            }
        }
        finally
        {
            updatingUi = false;
        }
    }

    private void DeviceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (updatingUi || DeviceComboBox.SelectedItem is not DeviceChoice choice)
            return;

        DeviceChanged?.Invoke(this, new MixerDeviceChangedEventArgs(session, choice.DeviceId));
    }

    private void DeviceComboBox_DropDownOpened(object sender, object e)
    {
        isComboInteracting = true;
        UpdateInteractionState();
    }

    private void DeviceComboBox_DropDownClosed(object sender, object e)
    {
        isComboInteracting = false;
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

        if (updatingUi)
            return;

        hasPendingVolumeCommit = true;
        volumeCommitTimer.Stop();
        volumeCommitTimer.Start();
    }

    private void UpdateVolumeText()
    {
        var value = Math.Clamp((int)Math.Round(VolumeSlider.Value), 0, 100);
        VolumeValueTextBlock.Text = $"{value}%";
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

        var volume = (float)(VolumeSlider.Value / 100d);
        if (Math.Abs(volume - lastCommittedVolume) < 0.005f)
        {
            hasPendingVolumeCommit = false;
            return;
        }

        hasPendingVolumeCommit = false;
        lastCommittedVolume = volume;
        VolumeChanged?.Invoke(this, new MixerVolumeChangedEventArgs(session, volume));
    }

    private void UpdateInteractionState()
    {
        var nextState = isVolumeInteracting || isComboInteracting;
        if (nextState == isInteracting)
            return;

        isInteracting = nextState;
        InteractionStateChanged?.Invoke(this, new MixerInteractionStateChangedEventArgs(session, isInteracting));
    }

    private sealed class DeviceChoice
    {
        public DeviceChoice(string? deviceId, string title)
        {
            DeviceId = deviceId;
            Title = title;
        }

        public string? DeviceId { get; }

        public string Title { get; }

        public override string ToString() => Title;
    }
}
