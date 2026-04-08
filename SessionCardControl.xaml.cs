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
        ActualDeviceTextBlock.Text = session.ActualDeviceSummary;
        BoundDeviceTextBlock.Text = session.CanChangeDevice
            ? session.BoundDeviceSummary
            : "System sessions cannot be rebound yet";
        RouteLabelTextBlock.Text = session.Flow == EDataFlow.eRender ? "Output device" : "Input device";
        VolumeLabelTextBlock.Text = session.Flow == EDataFlow.eRender ? "App volume" : "Input volume";
        SessionIcon.Glyph = session.IsSystemSession ? "\uE7F5" : "\uE77B";

        ToolTipService.SetToolTip(TitleTextBlock, session.DisplayName);
        ToolTipService.SetToolTip(ActualDeviceTextBlock, session.ActualDeviceSummary);
        ToolTipService.SetToolTip(BoundDeviceTextBlock, BoundDeviceTextBlock.Text);

        DeviceComboBox.IsEnabled = session.CanChangeDevice;
        PopulateDeviceChoices(devices);

        if (!isVolumeInteracting)
        {
            updatingUi = true;
            VolumeSlider.Value = Math.Clamp((int)Math.Round(session.Volume * 100), 0, 100);
            updatingUi = false;
            lastCommittedVolume = (float)(VolumeSlider.Value / 100d);
        }

        UpdateVolumeText();
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
