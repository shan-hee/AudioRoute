using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.UI.Composition;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Hosting;
using Microsoft.UI.Xaml.Input;
using WinRT.Interop;
using Windows.Graphics;

namespace AudioRoute;

public sealed partial class MainWindow : Window
{
    private const int PanelWidth = 472;
    private const int PanelHeight = 706;
    private const int ScreenMargin = 18;

    private readonly DispatcherQueueTimer refreshTimer;
    private readonly Dictionary<string, SessionCardControl> sessionCards = new(StringComparer.OrdinalIgnoreCase);
    private AppWindow? appWindow;
    private bool allowDeactivateClose;
    private bool isConfigured;
    private bool isClosing;
    private bool isRefreshing;
    private bool hasDeferredRefresh;
    private int interactionDepth;

    public MainWindow()
    {
        InitializeComponent();

        Activated += OnWindowActivated;
        WindowRoot.Loaded += OnWindowLoaded;
        WindowRoot.KeyDown += OnWindowRootKeyDown;

        refreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        refreshTimer.Interval = TimeSpan.FromSeconds(8);
        refreshTimer.Tick += (_, _) =>
        {
            if (!isClosing)
                RefreshData();
        };
        refreshTimer.Start();

        UpdateFlowButtons();
    }

    private EDataFlow CurrentFlow { get; set; } = EDataFlow.eRender;

    private async void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        if (isConfigured)
            return;

        isConfigured = true;

        ConfigureWindow();
        RefreshData();
        await ShowPanelAsync();
        allowDeactivateClose = true;
        DismissButton.Focus(FocusState.Programmatic);
    }

    private async void OnWindowActivated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState != WindowActivationState.Deactivated)
            return;

        if (!allowDeactivateClose || isClosing)
            return;

        await HidePanelAndExitAsync();
    }

    private async void DismissButton_Click(object sender, RoutedEventArgs e)
    {
        await HidePanelAndExitAsync();
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshData();
    }

    private void OutputButton_Click(object sender, RoutedEventArgs e)
    {
        SwitchFlow(EDataFlow.eRender);
    }

    private void InputButton_Click(object sender, RoutedEventArgs e)
    {
        SwitchFlow(EDataFlow.eCapture);
    }

    private async void OnWindowRootKeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key != Windows.System.VirtualKey.Escape)
            return;

        e.Handled = true;
        await HidePanelAndExitAsync();
    }

    private void RefreshData()
    {
        if (isClosing)
            return;

        if (isRefreshing)
        {
            hasDeferredRefresh = true;
            return;
        }

        if (interactionDepth > 0)
        {
            hasDeferredRefresh = true;
            return;
        }

        hasDeferredRefresh = false;
        isRefreshing = true;
        try
        {
            RefreshDataCore();
        }
        finally
        {
            isRefreshing = false;

            if (hasDeferredRefresh && interactionDepth == 0 && !isClosing)
            {
                hasDeferredRefresh = false;
                RefreshData();
            }
        }
    }

    private void RefreshDataCore()
    {
        var devices = DeviceEnumerator.EnumerateDevices(CurrentFlow);
        var defaultDevice = devices.Find(device => device.IsDefault);
        var sessions = AudioSessionService.GetActiveSessions(CurrentFlow);

        SummaryTextBlock.Text = CurrentFlow == EDataFlow.eRender
            ? $"Default output: {defaultDevice?.Name ?? "No device detected"}"
            : $"Default input: {defaultDevice?.Name ?? "No device detected"}";

        MetaTextBlock.Text = sessions.Count == 0
            ? (CurrentFlow == EDataFlow.eRender
                ? "Open an app that is actively playing audio and it will appear here."
                : "Open an app that is actively recording audio and it will appear here.")
            : $"{sessions.Count} active sessions. Routes update in place without restarting the panel.";

        if (sessions.Count == 0)
        {
            RemoveInactiveCards(Array.Empty<string>());
            ShowPlaceholder(
                CurrentFlow == EDataFlow.eRender
                    ? "No active output sessions right now."
                    : "No active input sessions right now.");
            return;
        }

        var desiredCards = new List<SessionCardControl>(sessions.Count);
        var activeKeys = new List<string>(sessions.Count);

        foreach (var session in sessions)
        {
            activeKeys.Add(session.SessionKey);

            if (!sessionCards.TryGetValue(session.SessionKey, out var card))
            {
                card = new SessionCardControl(session, devices);
                card.DeviceChanged += OnDeviceChanged;
                card.VolumeChanged += OnVolumeChanged;
                card.InteractionStateChanged += OnInteractionStateChanged;
                sessionCards.Add(session.SessionKey, card);
            }
            else
            {
                card.UpdateSession(session, devices);
            }

            desiredCards.Add(card);
        }

        RemoveInactiveCards(activeKeys);
        ReplaceSessionHostChildren(desiredCards);
    }

    private void SwitchFlow(EDataFlow flow)
    {
        if (CurrentFlow == flow)
            return;

        CurrentFlow = flow;
        UpdateFlowButtons();
        RefreshData();
    }

    private void UpdateFlowButtons()
    {
        ApplyFlowButton(OutputButton, CurrentFlow == EDataFlow.eRender);
        ApplyFlowButton(InputButton, CurrentFlow == EDataFlow.eCapture);
    }

    private static void ApplyFlowButton(Button button, bool selected)
    {
        button.Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(
            selected ? Windows.UI.Color.FromArgb(255, 49, 115, 194) : Windows.UI.Color.FromArgb(40, 255, 255, 255));
        button.BorderBrush = new Microsoft.UI.Xaml.Media.SolidColorBrush(
            selected ? Windows.UI.Color.FromArgb(255, 78, 156, 255) : Windows.UI.Color.FromArgb(24, 255, 255, 255));
        button.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
            selected ? Windows.UI.Color.FromArgb(255, 255, 255, 255) : Windows.UI.Color.FromArgb(220, 230, 230, 230));
    }

    private FrameworkElement BuildPlaceholder(string message)
    {
        return new Border
        {
            Padding = new Thickness(18),
            CornerRadius = new CornerRadius(14),
            Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(255, 42, 42, 42)),
            BorderBrush = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(32, 255, 255, 255)),
            BorderThickness = new Thickness(1),
            Child = new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.WrapWholeWords,
                Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(255, 200, 200, 200))
            }
        };
    }

    private void OnDeviceChanged(object? sender, MixerDeviceChangedEventArgs e)
    {
        try
        {
            if (e.Session.ProcessId <= 0)
                return;

            if (string.IsNullOrWhiteSpace(e.DeviceId))
                AudioPolicyManager.ClearAppDefaultDevice((uint)e.Session.ProcessId, e.Session.Flow);
            else
                AudioPolicyManager.SetAppDefaultDevice((uint)e.Session.ProcessId, e.DeviceId, e.Session.Flow);

            hasDeferredRefresh = true;
            RefreshData();
        }
        catch (Exception ex)
        {
            ShowError($"Failed to change device: {ex.Message}");
        }
    }

    private void OnVolumeChanged(object? sender, MixerVolumeChangedEventArgs e)
    {
        try
        {
            AudioSessionService.SetSessionVolume(e.Session.SessionKey, e.Session.Flow, e.Volume);
        }
        catch (Exception ex)
        {
            ShowError($"Failed to change volume: {ex.Message}");
        }
    }

    private void OnInteractionStateChanged(object? sender, MixerInteractionStateChangedEventArgs e)
    {
        interactionDepth += e.IsInteracting ? 1 : -1;
        if (interactionDepth < 0)
            interactionDepth = 0;

        if (interactionDepth == 0 && hasDeferredRefresh)
            RefreshData();
    }

    private async Task ShowPanelAsync()
    {
        var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
        visual.Opacity = 0f;
        visual.Offset = new Vector3(0f, 26f, 0f);
        await AnimatePanelAsync(visual, new Vector3(0f, 26f, 0f), Vector3.Zero, 0f, 1f);
    }

    private async Task HidePanelAndExitAsync()
    {
        if (isClosing)
            return;

        isClosing = true;
        allowDeactivateClose = false;
        refreshTimer.Stop();

        var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
        await AnimatePanelAsync(visual, Vector3.Zero, new Vector3(0f, 24f, 0f), 1f, 0f);
        Close();
    }

    private static Task AnimatePanelAsync(Visual visual, Vector3 fromOffset, Vector3 toOffset, float fromOpacity, float toOpacity)
    {
        var compositor = visual.Compositor;
        var batch = compositor.CreateScopedBatch(CompositionBatchTypes.Animation);
        var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        var easing = compositor.CreateCubicBezierEasingFunction(
            new Vector2(0.18f, 0.9f),
            new Vector2(0.24f, 1f));

        var offsetAnimation = compositor.CreateVector3KeyFrameAnimation();
        offsetAnimation.InsertKeyFrame(0f, fromOffset);
        offsetAnimation.InsertKeyFrame(1f, toOffset, easing);
        offsetAnimation.Duration = TimeSpan.FromMilliseconds(240);

        var opacityAnimation = compositor.CreateScalarKeyFrameAnimation();
        opacityAnimation.InsertKeyFrame(0f, fromOpacity);
        opacityAnimation.InsertKeyFrame(1f, toOpacity, easing);
        opacityAnimation.Duration = TimeSpan.FromMilliseconds(220);

        visual.StartAnimation(nameof(visual.Offset), offsetAnimation);
        visual.StartAnimation(nameof(visual.Opacity), opacityAnimation);

        batch.Completed += (_, _) => completed.TrySetResult();
        batch.End();

        return completed.Task;
    }

    private void ConfigureWindow()
    {
        var hwnd = WindowNative.GetWindowHandle(this);
        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
        appWindow = AppWindow.GetFromWindowId(windowId);

        if (appWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.SetBorderAndTitleBar(false, false);
            presenter.IsResizable = false;
            presenter.IsMaximizable = false;
            presenter.IsMinimizable = false;
        }

        NativeMethods.ApplyToolWindowStyle(hwnd);
        appWindow.Resize(new SizeInt32(PanelWidth, PanelHeight));

        var displayArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary);
        var x = displayArea.WorkArea.X + displayArea.WorkArea.Width - PanelWidth - ScreenMargin;
        var y = displayArea.WorkArea.Y + displayArea.WorkArea.Height - PanelHeight - ScreenMargin;
        appWindow.Move(new PointInt32(x, y));

        NativeMethods.SetTopMost(hwnd);
    }

    private void ShowError(string message)
    {
        var dialog = new ContentDialog
        {
            Title = "AudioRoute",
            Content = message,
            CloseButtonText = "OK",
            XamlRoot = WindowRoot.XamlRoot
        };

        _ = dialog.ShowAsync();
    }

    private void ShowPlaceholder(string message)
    {
        FrameworkElement placeholder;

        if (SessionHost.Children.Count == 1 && SessionHost.Children[0] is FrameworkElement existingPlaceholder && existingPlaceholder.Tag as string == "placeholder")
        {
            placeholder = existingPlaceholder;
            if (placeholder is Border border && border.Child is TextBlock textBlock)
                textBlock.Text = message;
        }
        else
        {
            placeholder = BuildPlaceholder(message);
            placeholder.Tag = "placeholder";
        }

        SessionHost.Children.Clear();
        SessionHost.Children.Add(placeholder);
    }

    private void RemoveInactiveCards(IReadOnlyCollection<string> activeKeys)
    {
        var activeLookup = new HashSet<string>(activeKeys, StringComparer.OrdinalIgnoreCase);
        var staleKeys = sessionCards.Keys.Where(key => !activeLookup.Contains(key)).ToList();

        foreach (var key in staleKeys)
        {
            var card = sessionCards[key];
            card.DeviceChanged -= OnDeviceChanged;
            card.VolumeChanged -= OnVolumeChanged;
            card.InteractionStateChanged -= OnInteractionStateChanged;
            sessionCards.Remove(key);
        }
    }

    private void ReplaceSessionHostChildren(IReadOnlyList<SessionCardControl> desiredCards)
    {
        var needsUpdate = SessionHost.Children.Count != desiredCards.Count;

        if (!needsUpdate)
        {
            for (var index = 0; index < desiredCards.Count; index++)
            {
                if (!ReferenceEquals(SessionHost.Children[index], desiredCards[index]))
                {
                    needsUpdate = true;
                    break;
                }
            }
        }

        if (!needsUpdate)
            return;

        SessionHost.Children.Clear();
        foreach (var card in desiredCards)
            SessionHost.Children.Add(card);
    }

    private static class NativeMethods
    {
        private const int GwlExStyle = -20;
        private const int WsExAppWindow = 0x00040000;
        private const int WsExToolWindow = 0x00000080;
        private const uint SwpNomove = 0x0002;
        private const uint SwpNosize = 0x0001;
        private const uint SwpNoactivate = 0x0010;
        private static readonly IntPtr HwndTopMost = new(-1);

        public static void ApplyToolWindowStyle(IntPtr hwnd)
        {
            var exStyle = GetWindowLongPtr(hwnd, GwlExStyle).ToInt64();
            exStyle |= WsExToolWindow;
            exStyle &= ~WsExAppWindow;
            _ = SetWindowLongPtr(hwnd, GwlExStyle, new IntPtr(exStyle));
        }

        public static void SetTopMost(IntPtr hwnd)
        {
            _ = SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNoactivate);
        }

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
    }
}
