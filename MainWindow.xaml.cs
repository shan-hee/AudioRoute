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
    private const int PanelWidth = 400;
    private const int PanelHeight = 660;
    private const int ScreenMargin = 18;
    private const uint TrayIconId = 1;
    private const uint TrayCallbackMessage = NativeMethods.WmApp + 1;
    private const uint HomeTrayMenuItemId = 1001;
    private const uint ExitTrayMenuItemId = 1002;

    private readonly DispatcherQueueTimer deactivateHideTimer;
    private readonly DispatcherQueueTimer refreshTimer;
    private readonly Dictionary<string, SessionCardControl> sessionCards = new(StringComparer.OrdinalIgnoreCase);
    private AppWindow? appWindow;
    private IntPtr hwnd;
    private bool allowDeactivateHide;
    private bool isConfigured;
    private bool isExitRequested;
    private bool isPanelVisible;
    private bool isRefreshing;
    private bool isVisibilityTransitioning;
    private bool hasDeferredRefresh;
    private bool isTrayIconCreated;
    private DateTimeOffset lastPrimaryTrayInvokeAt;
    private int interactionDepth;
    private WndProcDelegate? windowProcDelegate;
    private IntPtr originalWindowProc;

    public MainWindow()
    {
        InitializeComponent();

        Activated += OnWindowActivated;
        Closed += OnWindowClosed;
        WindowRoot.Loaded += OnWindowLoaded;
        WindowRoot.KeyDown += OnWindowRootKeyDown;

        deactivateHideTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        deactivateHideTimer.Interval = TimeSpan.FromMilliseconds(140);
        deactivateHideTimer.IsRepeating = false;
        deactivateHideTimer.Tick += async (_, _) =>
        {
            deactivateHideTimer.Stop();

            if (allowDeactivateHide && !isExitRequested && isPanelVisible && !isVisibilityTransitioning)
                await HideToTrayAsync();
        };

        refreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        refreshTimer.Interval = TimeSpan.FromSeconds(8);
        refreshTimer.Tick += (_, _) =>
        {
            if (!isExitRequested && isPanelVisible)
                RefreshData();
        };

        UpdateFlowButtons();
    }

    private EDataFlow CurrentFlow { get; set; } = EDataFlow.eRender;
    public bool IsPanelVisible => isPanelVisible;

    private async void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        if (isConfigured)
            return;

        isConfigured = true;

        ConfigureWindow();
        InitializeTrayIcon();
        RefreshData();
        await ShowPanelAsync(focusDismissButton: true);
    }

    private void OnWindowActivated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState != WindowActivationState.Deactivated)
            return;

        if (!allowDeactivateHide || isExitRequested || !isPanelVisible || isVisibilityTransitioning)
            return;

        ScheduleDeactivateHide();
    }

    private async void DismissButton_Click(object sender, RoutedEventArgs e)
    {
        await HideToTrayAsync();
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
        await HideToTrayAsync();
    }

    private void RefreshData()
    {
        if (isExitRequested)
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

            if (hasDeferredRefresh && interactionDepth == 0 && !isExitRequested)
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
            selected ? Windows.UI.Color.FromArgb(255, 56, 56, 56) : Windows.UI.Color.FromArgb(10, 255, 255, 255));
        button.BorderBrush = new Microsoft.UI.Xaml.Media.SolidColorBrush(
            selected ? Windows.UI.Color.FromArgb(255, 96, 205, 255) : Windows.UI.Color.FromArgb(18, 255, 255, 255));
        button.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(
            selected ? Windows.UI.Color.FromArgb(255, 96, 205, 255) : Windows.UI.Color.FromArgb(180, 200, 200, 200));
    }

    private FrameworkElement BuildPlaceholder(string message)
    {
        return new Border
        {
            Padding = new Thickness(16, 14, 16, 14),
            CornerRadius = new CornerRadius(8),
            Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(10, 255, 255, 255)),
            Child = new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.WrapWholeWords,
                FontSize = 12,
                Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Windows.UI.Color.FromArgb(255, 160, 160, 160))
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

    public Task ShowOrBringToFrontAsync()
    {
        CancelDeactivateHide();

        if (!isConfigured || isExitRequested || isVisibilityTransitioning)
            return Task.CompletedTask;

        if (isPanelVisible)
        {
            BringPanelToFront();
            return Task.CompletedTask;
        }

        return ShowPanelAsync(focusDismissButton: true);
    }

    public void PrepareForExit()
    {
        if (isExitRequested)
            return;

        isExitRequested = true;
        allowDeactivateHide = false;
        deactivateHideTimer.Stop();
        refreshTimer.Stop();
        DisposeTrayIcon();
    }

    private async Task ShowPanelAsync(bool focusDismissButton = false)
    {
        if (!isConfigured || isExitRequested || isVisibilityTransitioning || isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;

        try
        {
            RefreshData();
            refreshTimer.Start();
            UpdateWindowPosition();

            var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
            visual.Opacity = 0f;
            visual.Offset = new Vector3(0f, 26f, 0f);

            NativeMethods.ShowWindow(hwnd, NativeMethods.SwShow);
            Activate();
            NativeMethods.SetTopMost(hwnd);

            await AnimatePanelAsync(visual, new Vector3(0f, 26f, 0f), Vector3.Zero, 0f, 1f);
            isPanelVisible = true;
            allowDeactivateHide = true;
        }
        finally
        {
            isVisibilityTransitioning = false;
            allowDeactivateHide = isPanelVisible && !isExitRequested;
        }

        if (focusDismissButton)
            DismissButton.Focus(FocusState.Programmatic);
    }

    private async Task HideToTrayAsync()
    {
        if (!isConfigured || isExitRequested || isVisibilityTransitioning || !isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;
        refreshTimer.Stop();

        try
        {
            var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
            await AnimatePanelAsync(visual, Vector3.Zero, new Vector3(0f, 24f, 0f), 1f, 0f);
            NativeMethods.ShowWindow(hwnd, NativeMethods.SwHide);
            isPanelVisible = false;
        }
        finally
        {
            isVisibilityTransitioning = false;
        }
    }

    private void BringPanelToFront()
    {
        if (!isConfigured || isExitRequested)
            return;

        UpdateWindowPosition();
        NativeMethods.ShowWindow(hwnd, NativeMethods.SwShow);
        Activate();
        NativeMethods.SetForegroundWindow(hwnd);
        NativeMethods.SetTopMost(hwnd);
        allowDeactivateHide = true;
        DismissButton.Focus(FocusState.Programmatic);
    }

    private void ScheduleDeactivateHide()
    {
        deactivateHideTimer.Stop();
        deactivateHideTimer.Start();
    }

    private void CancelDeactivateHide()
    {
        if (deactivateHideTimer.IsRunning)
            deactivateHideTimer.Stop();
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
        hwnd = WindowNative.GetWindowHandle(this);
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
        AttachWindowProc();
        appWindow.Resize(new SizeInt32(PanelWidth, PanelHeight));
        UpdateWindowPosition();
        NativeMethods.SetTopMost(hwnd);
    }

    private void InitializeTrayIcon()
    {
        if (isTrayIconCreated)
            return;

        var iconHandle = NativeMethods.LoadApplicationIcon();
        var notifyIconData = CreateTrayIconData(iconHandle);

        if (!NativeMethods.ShellNotifyIcon(NativeMethods.NimAdd, ref notifyIconData))
            return;

        notifyIconData.uVersionOrTimeout = NativeMethods.NotifyIconVersion4;
        _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimSetVersion, ref notifyIconData);
        isTrayIconCreated = true;
    }

    private void DisposeTrayIcon()
    {
        if (isTrayIconCreated)
        {
            var notifyIconData = CreateTrayIconData(IntPtr.Zero);
            _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimDelete, ref notifyIconData);
            isTrayIconCreated = false;
        }

        DetachWindowProc();
    }

    private void AttachWindowProc()
    {
        if (originalWindowProc != IntPtr.Zero)
            return;

        windowProcDelegate = WindowProc;
        var windowProcPointer = Marshal.GetFunctionPointerForDelegate(windowProcDelegate);
        originalWindowProc = NativeMethods.SetWindowProc(hwnd, windowProcPointer);
    }

    private void DetachWindowProc()
    {
        if (originalWindowProc == IntPtr.Zero || hwnd == IntPtr.Zero)
            return;

        _ = NativeMethods.SetWindowProc(hwnd, originalWindowProc);
        originalWindowProc = IntPtr.Zero;
        windowProcDelegate = null;
    }

    private NativeMethods.NotifyIconData CreateTrayIconData(IntPtr iconHandle)
    {
        return new NativeMethods.NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.NotifyIconData>(),
            hWnd = hwnd,
            uID = TrayIconId,
            uFlags = NativeMethods.NifMessage | NativeMethods.NifIcon | NativeMethods.NifTip | NativeMethods.NifShowTip,
            uCallbackMessage = TrayCallbackMessage,
            hIcon = iconHandle,
            szTip = "AudioRoute",
            szInfo = string.Empty,
            szInfoTitle = string.Empty,
            uVersionOrTimeout = 0
        };
    }

    private IntPtr WindowProc(IntPtr windowHandle, uint message, IntPtr wParam, IntPtr lParam)
    {
        if (message == NativeMethods.WmTaskbarCreated)
        {
            if (!isExitRequested)
            {
                isTrayIconCreated = false;
                InitializeTrayIcon();
            }

            return IntPtr.Zero;
        }

        if (message == TrayCallbackMessage)
        {
            HandleTrayMessage(NativeMethods.GetLowWord(lParam), wParam);
            return IntPtr.Zero;
        }

        if (message == NativeMethods.WmCommand)
        {
            var commandId = NativeMethods.GetLowWord(wParam);
            if (commandId == HomeTrayMenuItemId)
            {
                _ = ShowOrBringToFrontAsync();
                return IntPtr.Zero;
            }

            if (commandId == ExitTrayMenuItemId)
            {
                RequestExit();
                return IntPtr.Zero;
            }
        }

        return originalWindowProc != IntPtr.Zero
            ? NativeMethods.CallWindowProc(originalWindowProc, windowHandle, message, wParam, lParam)
            : NativeMethods.DefWindowProc(windowHandle, message, wParam, lParam);
    }

    private void HandleTrayMessage(uint trayMessage, IntPtr invokePointData)
    {
        if (isExitRequested)
            return;

        if (trayMessage == NativeMethods.NinSelect || trayMessage == NativeMethods.WmLButtonUp)
        {
            if (DateTimeOffset.UtcNow - lastPrimaryTrayInvokeAt < TimeSpan.FromMilliseconds(200))
                return;

            lastPrimaryTrayInvokeAt = DateTimeOffset.UtcNow;
            CancelDeactivateHide();
            _ = ShowOrBringToFrontAsync();

            return;
        }

        if (trayMessage == NativeMethods.WmRButtonUp || trayMessage == NativeMethods.WmContextMenu)
        {
            CancelDeactivateHide();
            ShowTrayContextMenu(GetTrayMenuAnchorPoint(invokePointData));
        }
    }

    private Point GetTrayMenuAnchorPoint(IntPtr invokePointData)
    {
        var callbackPoint = NativeMethods.GetPointFromLParam(invokePointData);
        if (callbackPoint.X != 0 || callbackPoint.Y != 0)
            return callbackPoint;

        if (TryGetTrayIconRect(out var trayIconRect))
        {
            return new Point
            {
                X = trayIconRect.Left,
                Y = trayIconRect.Bottom
            };
        }

        return NativeMethods.GetCursorPos(out var cursorPosition)
            ? cursorPosition
            : default;
    }

    private bool TryGetTrayIconRect(out Rect trayIconRect)
    {
        var identifier = new NativeMethods.NotifyIconIdentifier
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.NotifyIconIdentifier>(),
            hWnd = hwnd,
            uID = TrayIconId
        };

        return NativeMethods.ShellNotifyIconGetRect(ref identifier, out trayIconRect) == 0;
    }

    private void ShowTrayContextMenu(Point anchorPoint)
    {
        var menuHandle = NativeMethods.CreatePopupMenu();
        if (menuHandle == IntPtr.Zero)
            return;

        try
        {
            if (anchorPoint.X == 0 && anchorPoint.Y == 0)
            {
                if (!NativeMethods.GetCursorPos(out anchorPoint))
                    return;
            }

            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfString, HomeTrayMenuItemId, "\u4E3B\u9875");
            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfString, ExitTrayMenuItemId, "\u9000\u51FA");

            NativeMethods.SetForegroundWindow(hwnd);
            _ = NativeMethods.TrackPopupMenu(
                menuHandle,
                NativeMethods.TpmBottomAlign | NativeMethods.TpmLeftAlign | NativeMethods.TpmRightButton,
                anchorPoint.X,
                anchorPoint.Y,
                0,
                hwnd,
                IntPtr.Zero);

            _ = NativeMethods.PostMessage(hwnd, NativeMethods.WmNull, IntPtr.Zero, IntPtr.Zero);
        }
        finally
        {
            _ = NativeMethods.DestroyMenu(menuHandle);
        }
    }

    private void RequestExit()
    {
        PrepareForExit();
        Close();
        App.Current.Exit();
    }

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        PrepareForExit();
    }

    private void UpdateWindowPosition()
    {
        if (appWindow is null)
            return;

        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
        var displayArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary);
        var x = displayArea.WorkArea.X + displayArea.WorkArea.Width - PanelWidth - ScreenMargin;
        var y = displayArea.WorkArea.Y + displayArea.WorkArea.Height - PanelHeight - ScreenMargin;
        appWindow.Move(new PointInt32(x, y));
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
        private const int GwlWndProc = -4;
        private const int IdiApplication = 32512;
        private const int WsExAppWindow = 0x00040000;
        private const int WsExToolWindow = 0x00000080;
        public const int SwHide = 0;
        public const int SwShow = 5;
        public const uint WmApp = 0x8000;
        public const uint WmCommand = 0x0111;
        public const uint WmContextMenu = 0x007B;
        public const uint WmLButtonUp = 0x0202;
        public const uint WmNull = 0x0000;
        public const uint WmRButtonUp = 0x0205;
        public const uint NinSelect = 0x0400;
        public const uint NimAdd = 0x00000000;
        public const uint NimDelete = 0x00000002;
        public const uint NimSetVersion = 0x00000004;
        public const uint NifMessage = 0x00000001;
        public const uint NifIcon = 0x00000002;
        public const uint NifTip = 0x00000004;
        public const uint NifShowTip = 0x00000080;
        public const uint NotifyIconVersion4 = 4;
        public const uint MfString = 0x00000000;
        public static readonly uint WmTaskbarCreated = RegisterWindowMessage("TaskbarCreated");
        private const uint SwpNomove = 0x0002;
        private const uint SwpNosize = 0x0001;
        private const uint SwpNoactivate = 0x0010;
        public const uint TpmBottomAlign = 0x0020;
        public const uint TpmLeftAlign = 0x0000;
        public const uint TpmRightButton = 0x0002;
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

        public static void ShowWindow(IntPtr hwnd, int command)
        {
            _ = ShowWindowNative(hwnd, command);
        }

        public static IntPtr SetWindowProc(IntPtr hwnd, IntPtr windowProc)
        {
            return SetWindowLongPtr(hwnd, GwlWndProc, windowProc);
        }

        public static bool ShellNotifyIcon(uint message, ref NotifyIconData notifyIconData)
        {
            return Shell_NotifyIcon(message, ref notifyIconData);
        }

        public static IntPtr LoadApplicationIcon()
        {
            return LoadIcon(IntPtr.Zero, new IntPtr(IdiApplication));
        }

        public static uint GetLowWord(IntPtr value)
        {
            return (uint)(value.ToInt64() & 0xFFFF);
        }

        public static Point GetPointFromLParam(IntPtr value)
        {
            var raw = value.ToInt64();
            return new Point
            {
                X = unchecked((short)(raw & 0xFFFF)),
                Y = unchecked((short)((raw >> 16) & 0xFFFF))
            };
        }

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        private static extern bool ShowWindowNative(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", EntryPoint = "CallWindowProcW", SetLastError = true)]
        public static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "DefWindowProcW", SetLastError = true)]
        public static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool AppendMenu(IntPtr hMenu, uint uFlags, uint uIDNewItem, string lpNewItem);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyMenu(IntPtr hMenu);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetCursorPos(out Point cursorPosition);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, IntPtr prcRect);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadIcon(IntPtr hInstance, IntPtr lpIconName);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool Shell_NotifyIcon(uint dwMessage, ref NotifyIconData lpData);

        [DllImport("shell32.dll", EntryPoint = "Shell_NotifyIconGetRect", SetLastError = true)]
        public static extern int ShellNotifyIconGetRect(ref NotifyIconIdentifier identifier, out Rect iconLocation);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint RegisterWindowMessage(string lpString);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct NotifyIconData
        {
            public uint cbSize;
            public IntPtr hWnd;
            public uint uID;
            public uint uFlags;
            public uint uCallbackMessage;
            public IntPtr hIcon;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szTip;

            public uint dwState;
            public uint dwStateMask;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string szInfo;

            public uint uVersionOrTimeout;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string szInfoTitle;

            public uint dwInfoFlags;
            public Guid guidItem;
            public IntPtr hBalloonIcon;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct NotifyIconIdentifier
        {
            public uint cbSize;
            public IntPtr hWnd;
            public uint uID;
            public Guid guidItem;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}
