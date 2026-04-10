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
using WinRT.Interop;
using Windows.Graphics;

namespace AudioRoute;

public sealed partial class MainWindow : Window
{
    private const int PanelWidth = 400;
    private const int PanelHeight = 460;
    private const int ScreenMargin = 18;
    private static readonly TimeSpan DeactivateHideDelay = TimeSpan.FromMilliseconds(140);
    private static readonly TimeSpan ForegroundMonitorInterval = TimeSpan.FromMilliseconds(650);
    private static readonly TimeSpan HideSuppressionDuration = TimeSpan.FromMilliseconds(320);
    private static readonly TimeSpan TrayReopenSuppressionDuration = TimeSpan.FromMilliseconds(300);
    private static readonly TimeSpan TrayIconRefreshInterval = TimeSpan.FromSeconds(1);
    private static readonly TimeSpan VisibleRefreshInterval = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan PanelOpenOffsetAnimationDuration = TimeSpan.FromMilliseconds(240);
    private static readonly TimeSpan PanelOpenOpacityAnimationDuration = TimeSpan.FromMilliseconds(220);
    private static readonly TimeSpan PanelCloseOffsetAnimationDuration = TimeSpan.FromMilliseconds(120);
    private static readonly TimeSpan PanelCloseOpacityAnimationDuration = TimeSpan.FromMilliseconds(100);
    private const uint TrayIconId = 1;
    private const uint TrayCallbackMessage = NativeMethods.WmApp + 1;
    private const uint HomeTrayMenuItemId = 1001;
    private const uint ExitTrayMenuItemId = 1002;
    private const uint StartupTrayMenuItemId = 1003;

    private readonly DispatcherQueueTimer deactivateHideTimer;
    private readonly DispatcherQueueTimer foregroundMonitorTimer;
    private readonly DispatcherQueueTimer refreshTimer;
    private readonly DispatcherQueueTimer trayIconTimer;
    private readonly StaThreadDispatcher refreshDispatcher;
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
    private DateTimeOffset suppressPrimaryTrayInvokeUntil;
    private MasterVolumeState? lastTrayVolumeState;
    private PanelSnapshot? lastSnapshot;
    private DateTimeOffset suppressHideUntil;
    private int interactionDepth;
    private int refreshGeneration;
    private WndProcDelegate? windowProcDelegate;
    private IntPtr originalWindowProc;

    public MainWindow()
    {
        InitializeComponent();
        refreshDispatcher = new StaThreadDispatcher("AudioRoute.Refresh");
        ConfigureWindow();

        Activated += OnWindowActivated;
        Closed += OnWindowClosed;

        deactivateHideTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        deactivateHideTimer.Interval = DeactivateHideDelay;
        deactivateHideTimer.IsRepeating = false;
        deactivateHideTimer.Tick += async (_, _) =>
        {
            deactivateHideTimer.Stop();

            if (ShouldHideForLostForeground())
                await HideToTrayAsync();
        };

        foregroundMonitorTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        foregroundMonitorTimer.Interval = ForegroundMonitorInterval;
        foregroundMonitorTimer.IsRepeating = true;
        foregroundMonitorTimer.Tick += async (_, _) =>
        {
            if (ShouldHideForLostForeground())
                await HideToTrayAsync();
        };

        refreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        refreshTimer.Interval = VisibleRefreshInterval;
        refreshTimer.Tick += (_, _) =>
        {
            if (!isExitRequested && isPanelVisible)
                RefreshData();
        };

        trayIconTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        trayIconTimer.Interval = TrayIconRefreshInterval;
        trayIconTimer.IsRepeating = true;
        trayIconTimer.Tick += (_, _) =>
        {
            if (!isExitRequested && isTrayIconCreated)
                UpdateTrayIcon();
        };

        EnsureConfigured();
    }

    public bool IsPanelVisible => isPanelVisible;

    private void OnWindowActivated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState != WindowActivationState.Deactivated)
            return;

        HandlePanelDeactivated();
    }

    private void EnsureConfigured()
    {
        if (isConfigured || isExitRequested)
            return;

        ConfigureWindow();
        if (hwnd == IntPtr.Zero)
            return;

        isConfigured = true;
        InitializeTrayIcon();
        RefreshData();
    }

    private void HandlePanelDeactivated()
    {
        if (!allowDeactivateHide || isExitRequested || !isPanelVisible || isVisibilityTransitioning || interactionDepth > 0)
            return;

        ScheduleDeactivateHide();
    }

    private void RefreshData()
    {
        if (isExitRequested)
            return;

        EnsureConfigured();

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
        var refreshVersion = ++refreshGeneration;
        _ = RefreshDataAsync(refreshVersion);
    }

    private async Task RefreshDataAsync(int refreshVersion)
    {
        try
        {
            var snapshot = await refreshDispatcher.InvokeAsync(BuildPanelSnapshot);
            if (isExitRequested || refreshVersion != refreshGeneration)
                return;

            lastSnapshot = snapshot;
            ApplyPanelSnapshot(snapshot);
        }
        catch when (isExitRequested)
        {
        }
        catch (Exception ex)
        {
            if (refreshVersion == refreshGeneration && isPanelVisible)
                ShowPlaceholder($"刷新音频会话失败: {ex.Message}");
        }
        finally
        {
            if (refreshVersion == refreshGeneration)
                isRefreshing = false;

            if (hasDeferredRefresh && interactionDepth == 0 && !isExitRequested)
            {
                hasDeferredRefresh = false;
                RefreshData();
            }
        }
    }

    private static IReadOnlyDictionary<string, AudioDevice> CreateDeviceMap(IReadOnlyList<AudioDevice> devices, EDataFlow flow)
    {
        var deviceMap = new Dictionary<string, AudioDevice>(devices.Count, StringComparer.OrdinalIgnoreCase);

        foreach (var device in devices)
        {
            if (device.Flow == flow)
                deviceMap[device.Id] = device;
        }

        return deviceMap;
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

    private static PanelSnapshot BuildPanelSnapshot()
    {
        var devices = DeviceEnumerator.EnumerateDevices(EDataFlow.eAll);
        var outputDeviceMap = CreateDeviceMap(devices, EDataFlow.eRender);
        var inputDeviceMap = CreateDeviceMap(devices, EDataFlow.eCapture);
        var outputSessions = AudioSessionService.GetActiveSessions(EDataFlow.eRender, outputDeviceMap);
        var inputSessions = AudioSessionService.GetActiveSessions(EDataFlow.eCapture, inputDeviceMap);
        var sessions = MergeSessions(outputSessions, inputSessions);

        return new PanelSnapshot(devices, sessions);
    }

    private void ApplyPanelSnapshot(PanelSnapshot snapshot)
    {
        if (snapshot.Sessions.Count == 0)
        {
            RemoveInactiveCards(Array.Empty<string>());
            ShowPlaceholder("当前没有活跃的音频会话。");
            return;
        }

        var desiredCards = new List<SessionCardControl>(snapshot.Sessions.Count);
        var activeKeys = new List<string>(snapshot.Sessions.Count);

        foreach (var session in snapshot.Sessions)
        {
            activeKeys.Add(session.SessionKey);

            if (!sessionCards.TryGetValue(session.SessionKey, out var card))
            {
                card = new SessionCardControl(session, snapshot.Devices);
                card.DeviceChanged += OnDeviceChanged;
                card.VolumeChanged += OnVolumeChanged;
                card.InteractionStateChanged += OnInteractionStateChanged;
                sessionCards.Add(session.SessionKey, card);
            }
            else
            {
                card.UpdateSession(session, snapshot.Devices);
            }

            desiredCards.Add(card);
        }

        RemoveInactiveCards(activeKeys);
        ReplaceSessionHostChildren(desiredCards);
    }

    private void ShowCachedSnapshotOrLoading()
    {
        if (lastSnapshot is not null)
        {
            ApplyPanelSnapshot(lastSnapshot);
            return;
        }

        ShowPlaceholder("正在加载音频会话...");
    }

    private static IReadOnlyList<MixerAppSessionInfo> MergeSessions(
        IReadOnlyList<MixerSessionInfo> outputSessions,
        IReadOnlyList<MixerSessionInfo> inputSessions)
    {
        var groupedSessions = new Dictionary<string, (MixerSessionInfo? Output, MixerSessionInfo? Input)>(StringComparer.OrdinalIgnoreCase);

        foreach (var session in outputSessions)
        {
            groupedSessions.TryGetValue(session.SessionKey, out var pair);
            pair.Output = session;
            groupedSessions[session.SessionKey] = pair;
        }

        foreach (var session in inputSessions)
        {
            groupedSessions.TryGetValue(session.SessionKey, out var pair);
            pair.Input = session;
            groupedSessions[session.SessionKey] = pair;
        }

        return groupedSessions
            .Values
            .Select(pair =>
            {
                var primarySession = pair.Output ?? pair.Input
                    ?? throw new InvalidOperationException("会话聚合结果不能为空。");

                return new MixerAppSessionInfo
                {
                    SessionKey = primarySession.SessionKey,
                    PrimarySession = primarySession,
                    OutputSession = pair.Output,
                    InputSession = pair.Input
                };
            })
            .OrderBy(session => session.IsSystemSession ? 0 : 1)
            .ThenBy(session => session.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
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
            ShowError($"更改设备失败: {ex.Message}");
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
            ShowError($"调整音量失败: {ex.Message}");
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

        if (!isConfigured && hwnd == IntPtr.Zero)
            Activate();

        EnsureConfigured();

        if (!isConfigured || isExitRequested || isVisibilityTransitioning)
            return Task.CompletedTask;

        if (isPanelVisible)
        {
            BringPanelToFront();
            return Task.CompletedTask;
        }

        return ShowPanelAsync();
    }

    private Task TogglePanelVisibilityAsync()
    {
        CancelDeactivateHide();

        if (!isConfigured && hwnd == IntPtr.Zero)
            Activate();

        EnsureConfigured();

        if (!isConfigured || isExitRequested || isVisibilityTransitioning)
            return Task.CompletedTask;

        return isPanelVisible
            ? HideToTrayAsync()
            : ShowPanelAsync();
    }

    public void PrepareForExit()
    {
        if (isExitRequested)
            return;

        isExitRequested = true;
        allowDeactivateHide = false;
        deactivateHideTimer.Stop();
        foregroundMonitorTimer.Stop();
        refreshTimer.Stop();
        trayIconTimer.Stop();
        refreshDispatcher.Dispose();
        DisposeTrayIcon();
        TrayVolumeIconService.Dispose();
        AudioPolicyManager.Cleanup();
    }

    private async Task ShowPanelAsync()
    {
        if (!isConfigured || isExitRequested || isVisibilityTransitioning || isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;

        try
        {
            ShowCachedSnapshotOrLoading();
            refreshTimer.Start();
            UpdateWindowPosition();

            WindowRoot.UpdateLayout();
            PanelSurface.UpdateLayout();

            var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
            visual.Opacity = 0f;
            visual.Offset = new Vector3(0f, 26f, 0f);

            NativeMethods.ShowWindow(hwnd, NativeMethods.SwShow);
            Activate();
            NativeMethods.SetForegroundWindow(hwnd);
            NativeMethods.SetTopMost(hwnd);
            NativeMethods.SetWindowCloaked(hwnd, false);

            await AnimatePanelAsync(visual, new Vector3(0f, 26f, 0f), Vector3.Zero, 0f, 1f);
            NativeMethods.SetForegroundWindow(hwnd);
            isPanelVisible = true;
            suppressHideUntil = DateTimeOffset.UtcNow + HideSuppressionDuration;
            allowDeactivateHide = true;
            foregroundMonitorTimer.Start();
            RefreshData();
        }
        finally
        {
            isVisibilityTransitioning = false;
            allowDeactivateHide = isPanelVisible && !isExitRequested;
        }
    }

    private async Task HideToTrayAsync()
    {
        if (!isConfigured || isExitRequested || isVisibilityTransitioning || !isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;
        foregroundMonitorTimer.Stop();
        refreshTimer.Stop();

        try
        {
            var visual = ElementCompositionPreview.GetElementVisual(PanelSurface);
            await AnimatePanelAsync(
                visual,
                Vector3.Zero,
                new Vector3(0f, 24f, 0f),
                1f,
                0f,
                PanelCloseOffsetAnimationDuration,
                PanelCloseOpacityAnimationDuration);
            NativeMethods.SetWindowCloaked(hwnd, true);
            NativeMethods.ShowWindow(hwnd, NativeMethods.SwHide);
            isPanelVisible = false;
            suppressPrimaryTrayInvokeUntil = DateTimeOffset.UtcNow + TrayReopenSuppressionDuration;
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
        NativeMethods.SetWindowCloaked(hwnd, false);
        suppressHideUntil = DateTimeOffset.UtcNow + HideSuppressionDuration;
        allowDeactivateHide = true;
        if (!foregroundMonitorTimer.IsRunning)
            foregroundMonitorTimer.Start();
    }

    private void ScheduleDeactivateHide()
    {
        deactivateHideTimer.Stop();
        deactivateHideTimer.Start();
    }

    private bool ShouldHideForLostForeground()
    {
        if (!allowDeactivateHide || isExitRequested || !isPanelVisible || isVisibilityTransitioning || interactionDepth > 0)
            return false;

        if (DateTimeOffset.UtcNow < suppressHideUntil)
            return false;

        var foregroundWindow = NativeMethods.GetForegroundWindow();
        if (foregroundWindow == IntPtr.Zero)
            return true;

        return NativeMethods.GetWindowProcessId(foregroundWindow) != Environment.ProcessId;
    }

    private void CancelDeactivateHide()
    {
        if (deactivateHideTimer.IsRunning)
            deactivateHideTimer.Stop();
    }

    private static Task AnimatePanelAsync(
        Visual visual,
        Vector3 fromOffset,
        Vector3 toOffset,
        float fromOpacity,
        float toOpacity,
        TimeSpan? offsetDuration = null,
        TimeSpan? opacityDuration = null)
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
        offsetAnimation.Duration = offsetDuration ?? PanelOpenOffsetAnimationDuration;

        var opacityAnimation = compositor.CreateScalarKeyFrameAnimation();
        opacityAnimation.InsertKeyFrame(0f, fromOpacity);
        opacityAnimation.InsertKeyFrame(1f, toOpacity, easing);
        opacityAnimation.Duration = opacityDuration ?? PanelOpenOpacityAnimationDuration;

        visual.StartAnimation(nameof(visual.Offset), offsetAnimation);
        visual.StartAnimation(nameof(visual.Opacity), opacityAnimation);

        batch.Completed += (_, _) => completed.TrySetResult();
        batch.End();

        return completed.Task;
    }

    private void ConfigureWindow()
    {
        if (hwnd == IntPtr.Zero)
            hwnd = WindowNative.GetWindowHandle(this);

        if (hwnd == IntPtr.Zero)
            return;

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
        NativeMethods.SetWindowCloaked(hwnd, true);
    }

    private void InitializeTrayIcon()
    {
        if (hwnd == IntPtr.Zero)
            return;

        if (isTrayIconCreated)
        {
            UpdateTrayIcon(force: true);
            return;
        }

        lastTrayVolumeState = MasterVolumeService.TryGetMasterVolumeState();
        var iconHandle = TrayVolumeIconService.GetIconHandle(lastTrayVolumeState?.IconKind ?? TrayVolumeIconKind.High);
        var notifyIconData = CreateTrayIconData(iconHandle, BuildTrayToolTip(lastTrayVolumeState));

        if (!NativeMethods.ShellNotifyIcon(NativeMethods.NimAdd, ref notifyIconData))
            return;

        notifyIconData.uVersionOrTimeout = NativeMethods.NotifyIconVersion4;
        _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimSetVersion, ref notifyIconData);
        isTrayIconCreated = true;
        trayIconTimer.Start();
        UpdateTrayIcon(force: true);
    }

    private void DisposeTrayIcon()
    {
        if (isTrayIconCreated)
        {
            var notifyIconData = CreateTrayIconData();
            _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimDelete, ref notifyIconData);
            isTrayIconCreated = false;
            lastTrayVolumeState = null;
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

    private NativeMethods.NotifyIconData CreateTrayIconData()
    {
        return CreateTrayIconData(IntPtr.Zero, string.Empty);
    }

    private NativeMethods.NotifyIconData CreateTrayIconData(IntPtr iconHandle, string toolTip)
    {
        var flags = NativeMethods.NifMessage;

        if (iconHandle != IntPtr.Zero)
            flags |= NativeMethods.NifIcon;

        if (!string.IsNullOrWhiteSpace(toolTip))
            flags |= NativeMethods.NifTip | NativeMethods.NifShowTip;

        return new NativeMethods.NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.NotifyIconData>(),
            hWnd = hwnd,
            uID = TrayIconId,
            uFlags = flags,
            uCallbackMessage = TrayCallbackMessage,
            hIcon = iconHandle,
            szTip = toolTip,
            szInfo = string.Empty,
            szInfoTitle = string.Empty,
            uVersionOrTimeout = 0
        };
    }

    private void UpdateTrayIcon(bool force = false)
    {
        if (!isTrayIconCreated)
            return;

        var currentState = MasterVolumeService.TryGetMasterVolumeState();
        if (!force && EqualityComparer<MasterVolumeState?>.Default.Equals(currentState, lastTrayVolumeState))
            return;

        lastTrayVolumeState = currentState;

        var iconHandle = TrayVolumeIconService.GetIconHandle(currentState?.IconKind ?? TrayVolumeIconKind.High);
        var notifyIconData = CreateTrayIconData(iconHandle, BuildTrayToolTip(currentState));
        _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimModify, ref notifyIconData);
    }

    private static string BuildTrayToolTip(MasterVolumeState? volumeState)
    {
        if (volumeState is null)
            return "AudioRoute";

        return volumeState.Value.IsMuted
            ? $"AudioRoute 主音量 {volumeState.Value.Percentage}% 已静音"
            : $"AudioRoute 主音量 {volumeState.Value.Percentage}%";
    }

    private IntPtr WindowProc(IntPtr windowHandle, uint message, IntPtr wParam, IntPtr lParam)
    {
        if (message == NativeMethods.WmTaskbarCreated)
        {
            if (!isExitRequested)
            {
                isTrayIconCreated = false;
                lastTrayVolumeState = null;
                InitializeTrayIcon();
            }

            return IntPtr.Zero;
        }

        if (message == TrayCallbackMessage)
        {
            HandleTrayMessage(NativeMethods.GetLowWord(lParam), wParam);
            return IntPtr.Zero;
        }

        if (message == NativeMethods.WmActivate &&
            NativeMethods.GetLowWord(wParam) == NativeMethods.WaInactive)
        {
            HandlePanelDeactivated();
        }

        if (message == NativeMethods.WmActivateApp &&
            wParam == IntPtr.Zero)
        {
            HandlePanelDeactivated();
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

            if (commandId == StartupTrayMenuItemId)
            {
                ToggleStartupRegistration();
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
            if (!isPanelVisible && DateTimeOffset.UtcNow < suppressPrimaryTrayInvokeUntil)
                return;

            if (DateTimeOffset.UtcNow - lastPrimaryTrayInvokeAt < TimeSpan.FromMilliseconds(200))
                return;

            lastPrimaryTrayInvokeAt = DateTimeOffset.UtcNow;
            CancelDeactivateHide();
            _ = TogglePanelVisibilityAsync();

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

            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfString, HomeTrayMenuItemId, "主页");
            _ = NativeMethods.AppendMenu(
                menuHandle,
                NativeMethods.MfString | (StartupManager.IsEnabled() ? NativeMethods.MfChecked : 0),
                StartupTrayMenuItemId,
                "开机自启");
            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfSeparator, 0, string.Empty);
            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfString, ExitTrayMenuItemId, "退出");

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
            RestoreTrayIconFocus();
        }
        finally
        {
            _ = NativeMethods.DestroyMenu(menuHandle);
        }
    }

    private void ToggleStartupRegistration()
    {
        try
        {
            var enabled = StartupManager.IsEnabled();
            StartupManager.SetEnabled(!enabled);
        }
        catch (Exception ex)
        {
            ShowError($"切换开机自启失败: {ex.Message}");
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

        if (TryGetTrayAnchoredPanelPosition(out var panelPosition))
        {
            appWindow.Move(panelPosition);
            return;
        }

        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
        var displayArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary);
        var workArea = displayArea.WorkArea;
        var x = workArea.X + workArea.Width - PanelWidth - ScreenMargin;
        var y = workArea.Y + workArea.Height - PanelHeight - ScreenMargin;
        appWindow.Move(new PointInt32(x, y));
    }

    private bool TryGetTrayAnchoredPanelPosition(out PointInt32 panelPosition)
    {
        panelPosition = default;

        if (!TryGetTrayIconRect(out var trayIconRect))
            return false;

        var trayCenter = new PointInt32(
            trayIconRect.Left + ((trayIconRect.Right - trayIconRect.Left) / 2),
            trayIconRect.Top + ((trayIconRect.Bottom - trayIconRect.Top) / 2));
        var displayArea = DisplayArea.GetFromPoint(trayCenter, DisplayAreaFallback.Nearest);
        if (displayArea is null)
            return false;

        panelPosition = CalculatePanelPosition(displayArea.WorkArea, trayIconRect);
        return true;
    }

    private void RestoreTrayIconFocus()
    {
        if (!isTrayIconCreated)
            return;

        var notifyIconData = new NativeMethods.NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.NotifyIconData>(),
            hWnd = hwnd,
            uID = TrayIconId
        };

        _ = NativeMethods.ShellNotifyIcon(NativeMethods.NimSetFocus, ref notifyIconData);
    }

    private static PointInt32 CalculatePanelPosition(RectInt32 workArea, Rect trayIconRect)
    {
        var minX = workArea.X + ScreenMargin;
        var minY = workArea.Y + ScreenMargin;
        var maxX = Math.Max(minX, workArea.X + workArea.Width - PanelWidth - ScreenMargin);
        var maxY = Math.Max(minY, workArea.Y + workArea.Height - PanelHeight - ScreenMargin);

        return DetectTrayDockEdge(workArea, trayIconRect) switch
        {
            TrayDockEdge.Left => new PointInt32(minX, maxY),
            TrayDockEdge.Top => new PointInt32(maxX, minY),
            TrayDockEdge.Right => new PointInt32(maxX, maxY),
            _ => new PointInt32(maxX, maxY)
        };
    }

    private static TrayDockEdge DetectTrayDockEdge(RectInt32 workArea, Rect trayIconRect)
    {
        var workAreaLeft = workArea.X;
        var workAreaTop = workArea.Y;
        var workAreaRight = workArea.X + workArea.Width;
        var workAreaBottom = workArea.Y + workArea.Height;

        var leftOverflow = Math.Max(0, workAreaLeft - trayIconRect.Left);
        var topOverflow = Math.Max(0, workAreaTop - trayIconRect.Top);
        var rightOverflow = Math.Max(0, trayIconRect.Right - workAreaRight);
        var bottomOverflow = Math.Max(0, trayIconRect.Bottom - workAreaBottom);

        if (bottomOverflow > 0 || rightOverflow > 0 || leftOverflow > 0 || topOverflow > 0)
        {
            if (bottomOverflow >= rightOverflow && bottomOverflow >= leftOverflow && bottomOverflow >= topOverflow)
                return TrayDockEdge.Bottom;

            if (rightOverflow >= leftOverflow && rightOverflow >= topOverflow)
                return TrayDockEdge.Right;

            if (leftOverflow >= topOverflow)
                return TrayDockEdge.Left;

            return TrayDockEdge.Top;
        }

        var leftDistance = Math.Abs(trayIconRect.Left - workAreaLeft);
        var topDistance = Math.Abs(trayIconRect.Top - workAreaTop);
        var rightDistance = Math.Abs(workAreaRight - trayIconRect.Right);
        var bottomDistance = Math.Abs(workAreaBottom - trayIconRect.Bottom);
        var nearestDistance = Math.Min(Math.Min(leftDistance, rightDistance), Math.Min(topDistance, bottomDistance));

        if (nearestDistance == bottomDistance)
            return TrayDockEdge.Bottom;

        if (nearestDistance == rightDistance)
            return TrayDockEdge.Right;

        if (nearestDistance == leftDistance)
            return TrayDockEdge.Left;

        return TrayDockEdge.Top;
    }

    private void ShowError(string message)
    {
        var dialog = new ContentDialog
        {
            Title = "AudioRoute",
            Content = message,
            CloseButtonText = "确定",
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
        private const int GwlStyle = -16;
        private const int GwlWndProc = -4;
        private const int WsBorder = 0x00800000;
        private const int WsCaption = 0x00C00000;
        private const int WsDlgFrame = 0x00400000;
        private const int WsExAppWindow = 0x00040000;
        private const int WsExToolWindow = 0x00000080;
        private const int WsThickFrame = 0x00040000;
        private const uint DwmaCloak = 13;
        public const int SwHide = 0;
        public const int SwShow = 5;
        public const uint WaInactive = 0;
        public const uint WmApp = 0x8000;
        public const uint WmActivate = 0x0006;
        public const uint WmActivateApp = 0x001C;
        public const uint WmCommand = 0x0111;
        public const uint WmContextMenu = 0x007B;
        public const uint WmLButtonUp = 0x0202;
        public const uint WmNull = 0x0000;
        public const uint WmRButtonUp = 0x0205;
        public const uint NinSelect = 0x0400;
        public const uint NimAdd = 0x00000000;
        public const uint NimModify = 0x00000001;
        public const uint NimDelete = 0x00000002;
        public const uint NimSetFocus = 0x00000003;
        public const uint NimSetVersion = 0x00000004;
        public const uint NifMessage = 0x00000001;
        public const uint NifIcon = 0x00000002;
        public const uint NifTip = 0x00000004;
        public const uint NifShowTip = 0x00000080;
        public const uint NotifyIconVersion4 = 4;
        public const uint MfString = 0x00000000;
        public const uint MfChecked = 0x00000008;
        public const uint MfSeparator = 0x00000800;
        public static readonly uint WmTaskbarCreated = RegisterWindowMessage("TaskbarCreated");
        private const uint SwpFrameChanged = 0x0020;
        private const uint SwpNomove = 0x0002;
        private const uint SwpNozorder = 0x0004;
        private const uint SwpNosize = 0x0001;
        private const uint SwpNoactivate = 0x0010;
        public const uint TpmBottomAlign = 0x0020;
        public const uint TpmLeftAlign = 0x0000;
        public const uint TpmRightButton = 0x0002;
        private static readonly IntPtr HwndTopMost = new(-1);

        public static void ApplyToolWindowStyle(IntPtr hwnd)
        {
            var style = GetWindowLongPtr(hwnd, GwlStyle).ToInt64();
            style &= ~(WsCaption | WsThickFrame | WsBorder | WsDlgFrame);
            _ = SetWindowLongPtr(hwnd, GwlStyle, new IntPtr(style));

            var exStyle = GetWindowLongPtr(hwnd, GwlExStyle).ToInt64();
            exStyle |= WsExToolWindow;
            exStyle &= ~WsExAppWindow;
            _ = SetWindowLongPtr(hwnd, GwlExStyle, new IntPtr(exStyle));

            _ = SetWindowPos(
                hwnd,
                IntPtr.Zero,
                0,
                0,
                0,
                0,
                SwpNomove | SwpNosize | SwpNozorder | SwpNoactivate | SwpFrameChanged);
        }

        public static void SetTopMost(IntPtr hwnd)
        {
            _ = SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNoactivate);
        }

        public static void ShowWindow(IntPtr hwnd, int command)
        {
            _ = ShowWindowNative(hwnd, command);
        }

        public static void SetWindowCloaked(IntPtr hwnd, bool cloaked)
        {
            var value = cloaked ? 1 : 0;
            _ = DwmSetWindowAttribute(hwnd, DwmaCloak, ref value, Marshal.SizeOf<int>());
        }

        public static IntPtr SetWindowProc(IntPtr hwnd, IntPtr windowProc)
        {
            return SetWindowLongPtr(hwnd, GwlWndProc, windowProc);
        }

        public static bool ShellNotifyIcon(uint message, ref NotifyIconData notifyIconData)
        {
            return Shell_NotifyIcon(message, ref notifyIconData);
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

        [DllImport("dwmapi.dll", SetLastError = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, uint dwAttribute, ref int pvAttribute, int cbAttribute);

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
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, IntPtr prcRect);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool Shell_NotifyIcon(uint dwMessage, ref NotifyIconData lpData);

        [DllImport("shell32.dll", EntryPoint = "Shell_NotifyIconGetRect", SetLastError = true)]
        public static extern int ShellNotifyIconGetRect(ref NotifyIconIdentifier identifier, out Rect iconLocation);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        public static uint GetWindowProcessId(IntPtr hWnd)
        {
            _ = GetWindowThreadProcessId(hWnd, out var processId);
            return processId;
        }

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

    private enum TrayDockEdge
    {
        Bottom,
        Left,
        Right,
        Top
    }

    private sealed record PanelSnapshot(IReadOnlyList<AudioDevice> Devices, IReadOnlyList<MixerAppSessionInfo> Sessions);

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}
