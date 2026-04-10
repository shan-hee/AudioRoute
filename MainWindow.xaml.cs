using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    private static readonly TimeSpan ForegroundMonitorFallbackInterval = TimeSpan.FromMilliseconds(650);
    private static readonly TimeSpan HideSuppressionDuration = TimeSpan.FromMilliseconds(320);
    private static readonly TimeSpan TrayReopenSuppressionDuration = TimeSpan.FromMilliseconds(300);
    private static readonly TimeSpan SnapshotHealthCheckInterval = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan SnapshotRefreshOnOpenAge = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan SnapshotRefreshWhileVisibleAge = TimeSpan.FromMinutes(3);
    private static readonly TimeSpan AudioChangeRefreshDebounceInterval = TimeSpan.FromMilliseconds(350);
    private static readonly TimeSpan InitialSnapshotWarmupWait = TimeSpan.FromMilliseconds(160);
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
    private readonly DispatcherQueueTimer audioChangeRefreshTimer;
    private readonly StaThreadDispatcher refreshDispatcher;
    private readonly AudioChangeMonitor audioChangeMonitor;
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
    private bool deferredRefreshCanReuseDevices;
    private RefreshSessionScope deferredRefreshScope = RefreshSessionScope.None;
    private bool isTrayIconCreated;
    private bool isSnapshotStale = true;
    private DateTimeOffset lastPrimaryTrayInvokeAt;
    private DateTimeOffset suppressPrimaryTrayInvokeUntil;
    private MasterVolumeState? lastTrayVolumeState;
    private PanelSnapshot? lastSnapshot;
    private DateTimeOffset lastSnapshotUpdatedAt;
    private DateTimeOffset suppressHideUntil;
    private int interactionDepth;
    private int refreshGeneration;
    private bool scheduledRefreshCanReuseDevices;
    private RefreshSessionScope scheduledRefreshScope = RefreshSessionScope.None;
    private Task? latestRefreshTask;
    private WinEventDelegate? foregroundEventDelegate;
    private IntPtr foregroundEventHook;
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
        foregroundMonitorTimer.Interval = ForegroundMonitorFallbackInterval;
        foregroundMonitorTimer.IsRepeating = true;
        foregroundMonitorTimer.Tick += async (_, _) =>
        {
            if (ShouldHideForLostForeground())
                await HideToTrayAsync();
        };

        refreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        refreshTimer.Interval = SnapshotHealthCheckInterval;
        refreshTimer.Tick += (_, _) =>
        {
            if (!isExitRequested && isPanelVisible)
                RefreshDataIfNeeded(SnapshotRefreshWhileVisibleAge);
        };

        audioChangeRefreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        audioChangeRefreshTimer.Interval = AudioChangeRefreshDebounceInterval;
        audioChangeRefreshTimer.IsRepeating = false;
        audioChangeRefreshTimer.Tick += (_, _) =>
        {
            audioChangeRefreshTimer.Stop();
            var canReuseDevices = scheduledRefreshCanReuseDevices;
            var scope = scheduledRefreshScope;
            scheduledRefreshCanReuseDevices = false;
            scheduledRefreshScope = RefreshSessionScope.None;

            if (!isExitRequested && (isPanelVisible || lastSnapshot is not null))
                RefreshData(canReuseDevices, scope);
        };

        audioChangeMonitor = new AudioChangeMonitor();
        audioChangeMonitor.Changed += OnAudioEnvironmentChanged;
        audioChangeMonitor.SessionStructureChanged += OnObservedSessionStructureChanged;
        audioChangeMonitor.SessionVolumeChanged += OnObservedSessionVolumeChanged;
        audioChangeMonitor.MasterVolumeChanged += OnObservedMasterVolumeChanged;
        audioChangeMonitor.SessionDisplayNameChanged += OnObservedSessionDisplayNameChanged;

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
        AttachForegroundEventHook();
        InitializeTrayIcon();
        RefreshData();
    }

    private void HandlePanelDeactivated()
    {
        if (!allowDeactivateHide || isExitRequested || !isPanelVisible || isVisibilityTransitioning || interactionDepth > 0)
            return;

        ScheduleDeactivateHide();
    }

    private void OnAudioEnvironmentChanged(object? sender, EventArgs e)
    {
        if (isExitRequested)
            return;

        DispatcherQueue.TryEnqueue(HandleAudioEnvironmentChanged);
    }

    private void OnObservedSessionStructureChanged(object? sender, ObservedSessionStructureChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        DispatcherQueue.TryEnqueue(() => HandleObservedSessionStructureChanged(e));
    }

    private void OnObservedSessionVolumeChanged(object? sender, ObservedSessionVolumeChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        DispatcherQueue.TryEnqueue(() => HandleObservedSessionVolumeChanged(e));
    }

    private void OnObservedMasterVolumeChanged(object? sender, ObservedMasterVolumeChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        DispatcherQueue.TryEnqueue(() => HandleObservedMasterVolumeChanged(e));
    }

    private void OnObservedSessionDisplayNameChanged(object? sender, ObservedSessionDisplayNameChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        DispatcherQueue.TryEnqueue(() => HandleObservedSessionDisplayNameChanged(e));
    }

    private void HandleAudioEnvironmentChanged()
    {
        if (isExitRequested)
            return;

        isSnapshotStale = true;

        if (!isPanelVisible && lastSnapshot is null)
            return;

        ScheduleAudioRefresh(canReuseDevices: false, scope: RefreshSessionScope.All);
    }

    private void HandleObservedSessionStructureChanged(ObservedSessionStructureChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        isSnapshotStale = true;

        if (!isPanelVisible && lastSnapshot is null)
            return;

        ScheduleAudioRefresh(
            canReuseDevices: lastSnapshot is not null,
            scope: GetRefreshScope(e.Flow));
    }

    private void HandleObservedSessionVolumeChanged(ObservedSessionVolumeChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        UpdateCachedSessionVolume(e.SessionKey, e.Flow, e.Volume, e.IsMuted);

        if (!isPanelVisible)
            return;

        if (sessionCards.TryGetValue(e.SessionKey, out var card))
            card.ApplyObservedVolume(e.Flow, e.Volume, e.IsMuted);
    }

    private void HandleObservedMasterVolumeChanged(ObservedMasterVolumeChangedEventArgs e)
    {
        if (isExitRequested || !isTrayIconCreated)
            return;

        UpdateTrayIcon(e.State);
    }

    private void HandleObservedSessionDisplayNameChanged(ObservedSessionDisplayNameChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        if (string.IsNullOrWhiteSpace(e.DisplayName))
        {
            isSnapshotStale = true;

            if (!isPanelVisible && lastSnapshot is null)
                return;

            ScheduleAudioRefresh(
                canReuseDevices: lastSnapshot is not null,
                scope: GetRefreshScope(e.Flow));
            return;
        }

        if (!TryUpdateCachedSessionDisplayName(e.SessionKey, e.Flow, e.DisplayName, out var updatedAppSession))
            return;

        if (!isPanelVisible || lastSnapshot is null || updatedAppSession is null)
            return;

        if (sessionCards.TryGetValue(e.SessionKey, out var card))
            card.UpdateSession(updatedAppSession, lastSnapshot.Devices);

        ReorderSessionCardsFromSnapshot();
    }

    private void ScheduleAudioRefresh(bool canReuseDevices, RefreshSessionScope scope = RefreshSessionScope.All)
    {
        scope &= RefreshSessionScope.All;
        if (scope == RefreshSessionScope.None)
            return;

        if (!audioChangeRefreshTimer.IsRunning)
        {
            scheduledRefreshCanReuseDevices = canReuseDevices;
            scheduledRefreshScope = scope;
        }
        else
        {
            scheduledRefreshCanReuseDevices &= canReuseDevices;
            scheduledRefreshScope |= scope;
        }

        audioChangeRefreshTimer.Stop();
        audioChangeRefreshTimer.Start();
    }

    private void RefreshData(bool canReuseDevices = false, RefreshSessionScope scope = RefreshSessionScope.All)
    {
        if (isExitRequested)
            return;

        scope &= RefreshSessionScope.All;
        if (scope == RefreshSessionScope.None)
            return;

        EnsureConfigured();

        if (isRefreshing)
        {
            deferredRefreshCanReuseDevices = hasDeferredRefresh
                ? deferredRefreshCanReuseDevices && canReuseDevices
                : canReuseDevices;
            deferredRefreshScope = hasDeferredRefresh
                ? deferredRefreshScope | scope
                : scope;
            hasDeferredRefresh = true;
            return;
        }

        if (interactionDepth > 0)
        {
            deferredRefreshCanReuseDevices = hasDeferredRefresh
                ? deferredRefreshCanReuseDevices && canReuseDevices
                : canReuseDevices;
            deferredRefreshScope = hasDeferredRefresh
                ? deferredRefreshScope | scope
                : scope;
            hasDeferredRefresh = true;
            return;
        }

        hasDeferredRefresh = false;
        deferredRefreshCanReuseDevices = false;
        deferredRefreshScope = RefreshSessionScope.None;
        isRefreshing = true;
        var refreshVersion = ++refreshGeneration;
        var cachedSnapshot = lastSnapshot;
        var cachedDevices = canReuseDevices ? cachedSnapshot?.Devices : null;
        var refreshTask = RefreshDataAsync(refreshVersion, cachedSnapshot, cachedDevices, scope);
        latestRefreshTask = refreshTask;
        _ = refreshTask;
    }

    private async Task RefreshDataAsync(
        int refreshVersion,
        PanelSnapshot? cachedSnapshot,
        IReadOnlyList<AudioDevice>? cachedDevices,
        RefreshSessionScope scope)
    {
        try
        {
            var snapshot = await refreshDispatcher.InvokeAsync(() => BuildPanelSnapshot(cachedSnapshot, cachedDevices, scope));
            if (isExitRequested || refreshVersion != refreshGeneration)
                return;

            var previousSnapshot = lastSnapshot;
            lastSnapshot = snapshot;
            lastSnapshotUpdatedAt = DateTimeOffset.UtcNow;
            isSnapshotStale = false;

            if (isPanelVisible)
                ApplyPanelSnapshot(snapshot, previousSnapshot);
        }
        catch when (isExitRequested)
        {
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] 刷新音频会话失败: {ex}");
            if (refreshVersion == refreshGeneration && isPanelVisible)
                ShowPlaceholder($"刷新音频会话失败: {ex.Message}");
        }
        finally
        {
            if (refreshVersion == refreshGeneration)
                isRefreshing = false;

            if (hasDeferredRefresh && interactionDepth == 0 && !isExitRequested)
            {
                var canReuseDevices = deferredRefreshCanReuseDevices;
                var deferredScope = deferredRefreshScope;
                hasDeferredRefresh = false;
                deferredRefreshCanReuseDevices = false;
                deferredRefreshScope = RefreshSessionScope.None;
                RefreshData(canReuseDevices, deferredScope);
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

    private static PanelSnapshot BuildPanelSnapshot(
        PanelSnapshot? cachedSnapshot,
        IReadOnlyList<AudioDevice>? devices = null,
        RefreshSessionScope scope = RefreshSessionScope.All)
    {
        scope &= RefreshSessionScope.All;
        if (scope == RefreshSessionScope.None || cachedSnapshot is null)
            scope = RefreshSessionScope.All;

        devices ??= DeviceEnumerator.EnumerateDevices(EDataFlow.eAll);
        var outputSessions = (scope & RefreshSessionScope.Render) != 0
            ? AudioSessionService.GetActiveSessions(EDataFlow.eRender, CreateDeviceMap(devices, EDataFlow.eRender))
            : GetSessionsForFlow(cachedSnapshot!, EDataFlow.eRender);
        var inputSessions = (scope & RefreshSessionScope.Capture) != 0
            ? AudioSessionService.GetActiveSessions(EDataFlow.eCapture, CreateDeviceMap(devices, EDataFlow.eCapture))
            : GetSessionsForFlow(cachedSnapshot!, EDataFlow.eCapture);
        var sessions = MergeSessions(outputSessions, inputSessions);

        return new PanelSnapshot(devices, sessions);
    }

    private static IReadOnlyList<MixerSessionInfo> GetSessionsForFlow(PanelSnapshot snapshot, EDataFlow flow)
    {
        var sessions = new List<MixerSessionInfo>(snapshot.Sessions.Count);

        foreach (var appSession in snapshot.Sessions)
        {
            var session = appSession.GetSession(flow);
            if (session is not null)
                sessions.Add(session);
        }

        return sessions;
    }

    private static RefreshSessionScope GetRefreshScope(EDataFlow flow)
    {
        return flow switch
        {
            EDataFlow.eRender => RefreshSessionScope.Render,
            EDataFlow.eCapture => RefreshSessionScope.Capture,
            _ => RefreshSessionScope.All
        };
    }

    private void ApplyPanelSnapshot(PanelSnapshot snapshot, PanelSnapshot? previousSnapshot = null)
    {
        if (snapshot.Sessions.Count == 0)
        {
            RemoveInactiveCards(Array.Empty<string>());
            ShowPlaceholder("当前没有活跃的音频会话。");
            return;
        }

        var previousSessionMap = previousSnapshot?.Sessions.ToDictionary(session => session.SessionKey, StringComparer.OrdinalIgnoreCase);
        var devicesChanged = !AreDevicesEquivalent(previousSnapshot?.Devices, snapshot.Devices);
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
                var needsUpdate = devicesChanged ||
                    previousSessionMap is null ||
                    !previousSessionMap.TryGetValue(session.SessionKey, out var previousSession) ||
                    !AreAppSessionsEquivalent(previousSession, session);

                if (needsUpdate)
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
            ApplyPanelSnapshot(lastSnapshot, lastSnapshot);
            return;
        }

        ShowPlaceholder("正在加载音频会话...");
    }

    private async Task WaitForInitialSnapshotWarmupAsync()
    {
        if (lastSnapshot is not null || !isRefreshing)
            return;

        var refreshTask = latestRefreshTask;
        if (refreshTask is null || refreshTask.IsCompleted)
            return;

        try
        {
            await Task.WhenAny(refreshTask, Task.Delay(InitialSnapshotWarmupWait));
        }
        catch when (isExitRequested)
        {
        }
    }

    private static IReadOnlyList<MixerAppSessionInfo> OrderSessions(IEnumerable<MixerAppSessionInfo> sessions)
    {
        return sessions
            .OrderBy(session => session.IsSystemSession ? 0 : 1)
            .ThenBy(session => session.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    private void RefreshDataIfNeeded(TimeSpan maxSnapshotAge)
    {
        if (lastSnapshot is null || isSnapshotStale)
        {
            RefreshData();
            return;
        }

        if (DateTimeOffset.UtcNow - lastSnapshotUpdatedAt >= maxSnapshotAge)
            RefreshData();
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

        return OrderSessions(groupedSessions
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
            }));
    }

    private void OnDeviceChanged(object? sender, MixerDeviceChangedEventArgs e)
    {
        try
        {
            if (e.Session.ProcessId <= 0)
                return;

            if (!e.Session.IsRoutingSupported)
            {
                ShowError(e.Session.RoutingUnavailableReason ?? "当前系统不支持应用级音频路由。");
                return;
            }

            if (string.IsNullOrWhiteSpace(e.DeviceId))
                AudioPolicyManager.ClearAppDefaultDevice((uint)e.Session.ProcessId, e.Session.Flow);
            else
                AudioPolicyManager.SetAppDefaultDevice((uint)e.Session.ProcessId, e.DeviceId, e.Session.Flow);

            UpdateCachedSessionRoute(e.Session.SessionKey, e.Session.Flow, e.DeviceId, e.SelectedDeviceSummary);
            isSnapshotStale = true;

            if (isPanelVisible)
                ScheduleAudioRefresh(canReuseDevices: true, scope: GetRefreshScope(e.Session.Flow));
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] 更改设备失败: {ex}");
            hasDeferredRefresh = true;
            deferredRefreshCanReuseDevices = false;
            RefreshData();
            ShowError($"更改设备失败: {ex.Message}");
        }
    }

    private void OnVolumeChanged(object? sender, MixerVolumeChangedEventArgs e)
    {
        try
        {
            if (!audioChangeMonitor.TrySetSessionVolume(e.Session.SessionKey, e.Session.Flow, e.Volume))
                AudioSessionService.SetSessionVolume(e.Session.SessionKey, e.Session.Flow, e.Volume);

            UpdateCachedSessionVolume(e.Session.SessionKey, e.Session.Flow, e.Volume, isMuted: e.Volume <= 0.005f);
        }
        catch (Exception ex)
        {
            if (sender is SessionCardControl card)
                card.NotifyVolumeCommitFailed();

            Trace.WriteLine($"[AudioRoute] 调整音量失败: {ex}");
            hasDeferredRefresh = true;
            RefreshData();
            ShowError($"调整音量失败: {ex.Message}");
        }
    }

    private void OnInteractionStateChanged(object? sender, MixerInteractionStateChangedEventArgs e)
    {
        interactionDepth += e.IsInteracting ? 1 : -1;
        if (interactionDepth < 0)
            interactionDepth = 0;

        if (interactionDepth == 0 && hasDeferredRefresh)
        {
            var canReuseDevices = deferredRefreshCanReuseDevices;
            var scope = deferredRefreshScope;
            hasDeferredRefresh = false;
            deferredRefreshCanReuseDevices = false;
            deferredRefreshScope = RefreshSessionScope.None;
            ScheduleAudioRefresh(canReuseDevices, scope);
        }
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
        audioChangeRefreshTimer.Stop();
        DetachForegroundEventHook();
        audioChangeMonitor.Changed -= OnAudioEnvironmentChanged;
        audioChangeMonitor.SessionStructureChanged -= OnObservedSessionStructureChanged;
        audioChangeMonitor.SessionVolumeChanged -= OnObservedSessionVolumeChanged;
        audioChangeMonitor.MasterVolumeChanged -= OnObservedMasterVolumeChanged;
        audioChangeMonitor.SessionDisplayNameChanged -= OnObservedSessionDisplayNameChanged;
        audioChangeMonitor.Dispose();
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
            await WaitForInitialSnapshotWarmupAsync();
            if (isExitRequested || isPanelVisible)
                return;

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
            StartForegroundMonitorFallbackIfNeeded();
            RefreshDataIfNeeded(SnapshotRefreshOnOpenAge);
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
        StartForegroundMonitorFallbackIfNeeded();
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

    private void AttachForegroundEventHook()
    {
        if (foregroundEventHook != IntPtr.Zero)
            return;

        foregroundEventDelegate = ForegroundEventProc;
        foregroundEventHook = NativeMethods.SetWinEventHook(
            NativeMethods.EventSystemForeground,
            NativeMethods.EventSystemForeground,
            IntPtr.Zero,
            foregroundEventDelegate,
            0,
            0,
            NativeMethods.WineventOutofcontext | NativeMethods.WineventSkipOwnProcess);

        if (foregroundEventHook == IntPtr.Zero)
            Trace.WriteLine("[AudioRoute] Failed to attach foreground event hook, fallback timer will be used.");
    }

    private void DetachForegroundEventHook()
    {
        if (foregroundEventHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(foregroundEventHook);
            foregroundEventHook = IntPtr.Zero;
        }

        foregroundEventDelegate = null;
    }

    private void StartForegroundMonitorFallbackIfNeeded()
    {
        if (foregroundEventHook != IntPtr.Zero)
            return;

        if (!foregroundMonitorTimer.IsRunning)
            foregroundMonitorTimer.Start();
    }

    private void ForegroundEventProc(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr eventWindow,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime)
    {
        if (eventType != NativeMethods.EventSystemForeground ||
            idObject != NativeMethods.ObjIdWindow ||
            isExitRequested ||
            !isPanelVisible ||
            !allowDeactivateHide)
        {
            return;
        }

        DispatcherQueue.TryEnqueue(HandlePanelDeactivated);
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
        UpdateTrayIcon(lastTrayVolumeState, force: true);
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
        UpdateTrayIcon(MasterVolumeService.TryGetMasterVolumeState(), force);
    }

    private void UpdateTrayIcon(MasterVolumeState? currentState, bool force = false)
    {
        if (!isTrayIconCreated)
            return;

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

    private void UpdateCachedSessionVolume(string sessionKey, EDataFlow flow, float volume, bool isMuted)
    {
        if (lastSnapshot is null)
            return;

        var changed = false;
        var updatedSessions = new List<MixerAppSessionInfo>(lastSnapshot.Sessions.Count);

        foreach (var appSession in lastSnapshot.Sessions)
        {
            if (!string.Equals(appSession.SessionKey, sessionKey, StringComparison.OrdinalIgnoreCase))
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var session = appSession.GetSession(flow);
            if (session is null)
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var updatedSession = CloneSessionWithVolume(session, volume, isMuted);
            var outputSession = flow == EDataFlow.eRender ? updatedSession : appSession.OutputSession;
            var inputSession = flow == EDataFlow.eCapture ? updatedSession : appSession.InputSession;

            updatedSessions.Add(new MixerAppSessionInfo
            {
                SessionKey = appSession.SessionKey,
                PrimarySession = outputSession ?? inputSession ?? updatedSession,
                OutputSession = outputSession,
                InputSession = inputSession
            });

            changed = true;
        }

        if (changed)
        {
            lastSnapshot = lastSnapshot with { Sessions = updatedSessions };
            lastSnapshotUpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    private void UpdateCachedSessionRoute(string sessionKey, EDataFlow flow, string? boundDeviceId, string boundDeviceSummary)
    {
        if (lastSnapshot is null)
            return;

        var changed = false;
        var updatedSessions = new List<MixerAppSessionInfo>(lastSnapshot.Sessions.Count);

        foreach (var appSession in lastSnapshot.Sessions)
        {
            if (!string.Equals(appSession.SessionKey, sessionKey, StringComparison.OrdinalIgnoreCase))
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var session = appSession.GetSession(flow);
            if (session is null)
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var updatedSession = CloneSessionWithRoute(session, boundDeviceId, boundDeviceSummary);
            var outputSession = flow == EDataFlow.eRender ? updatedSession : appSession.OutputSession;
            var inputSession = flow == EDataFlow.eCapture ? updatedSession : appSession.InputSession;

            updatedSessions.Add(new MixerAppSessionInfo
            {
                SessionKey = appSession.SessionKey,
                PrimarySession = outputSession ?? inputSession ?? updatedSession,
                OutputSession = outputSession,
                InputSession = inputSession
            });

            changed = true;
        }

        if (changed)
        {
            lastSnapshot = lastSnapshot with { Sessions = updatedSessions };
            lastSnapshotUpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    private bool TryUpdateCachedSessionDisplayName(
        string sessionKey,
        EDataFlow flow,
        string displayName,
        out MixerAppSessionInfo? updatedAppSession)
    {
        updatedAppSession = null;

        if (lastSnapshot is null)
            return false;

        var changed = false;
        var updatedSessions = new List<MixerAppSessionInfo>(lastSnapshot.Sessions.Count);

        foreach (var appSession in lastSnapshot.Sessions)
        {
            if (!string.Equals(appSession.SessionKey, sessionKey, StringComparison.OrdinalIgnoreCase))
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var session = appSession.GetSession(flow);
            if (session is null || string.Equals(session.DisplayName, displayName, StringComparison.Ordinal))
            {
                updatedSessions.Add(appSession);
                continue;
            }

            var updatedSession = CloneSessionWithDisplayName(session, displayName);
            var outputSession = flow == EDataFlow.eRender ? updatedSession : appSession.OutputSession;
            var inputSession = flow == EDataFlow.eCapture ? updatedSession : appSession.InputSession;

            updatedAppSession = new MixerAppSessionInfo
            {
                SessionKey = appSession.SessionKey,
                PrimarySession = outputSession ?? inputSession ?? updatedSession,
                OutputSession = outputSession,
                InputSession = inputSession
            };

            updatedSessions.Add(updatedAppSession);
            changed = true;
        }

        if (!changed)
            return false;

        lastSnapshot = lastSnapshot with { Sessions = OrderSessions(updatedSessions) };
        lastSnapshotUpdatedAt = DateTimeOffset.UtcNow;
        return true;
    }

    private static MixerSessionInfo CloneSessionWithVolume(MixerSessionInfo session, float volume, bool isMuted)
    {
        return new MixerSessionInfo
        {
            SessionKey = session.SessionKey,
            DisplayName = session.DisplayName,
            ActualDeviceSummary = session.ActualDeviceSummary,
            BoundDeviceSummary = session.BoundDeviceSummary,
            ProcessName = session.ProcessName,
            ExecutablePath = session.ExecutablePath,
            BoundDeviceId = session.BoundDeviceId,
            RoutingUnavailableReason = session.RoutingUnavailableReason,
            Flow = session.Flow,
            ProcessId = session.ProcessId,
            Volume = volume,
            IsMuted = isMuted,
            IsSystemSession = session.IsSystemSession,
            IsRoutingSupported = session.IsRoutingSupported
        };
    }

    private static MixerSessionInfo CloneSessionWithRoute(MixerSessionInfo session, string? boundDeviceId, string boundDeviceSummary)
    {
        return new MixerSessionInfo
        {
            SessionKey = session.SessionKey,
            DisplayName = session.DisplayName,
            ActualDeviceSummary = session.ActualDeviceSummary,
            BoundDeviceSummary = boundDeviceSummary,
            ProcessName = session.ProcessName,
            ExecutablePath = session.ExecutablePath,
            BoundDeviceId = boundDeviceId,
            RoutingUnavailableReason = session.RoutingUnavailableReason,
            Flow = session.Flow,
            ProcessId = session.ProcessId,
            Volume = session.Volume,
            IsMuted = session.IsMuted,
            IsSystemSession = session.IsSystemSession,
            IsRoutingSupported = session.IsRoutingSupported
        };
    }

    private static MixerSessionInfo CloneSessionWithDisplayName(MixerSessionInfo session, string displayName)
    {
        return new MixerSessionInfo
        {
            SessionKey = session.SessionKey,
            DisplayName = displayName,
            ActualDeviceSummary = session.ActualDeviceSummary,
            BoundDeviceSummary = session.BoundDeviceSummary,
            ProcessName = session.ProcessName,
            ExecutablePath = session.ExecutablePath,
            BoundDeviceId = session.BoundDeviceId,
            RoutingUnavailableReason = session.RoutingUnavailableReason,
            Flow = session.Flow,
            ProcessId = session.ProcessId,
            Volume = session.Volume,
            IsMuted = session.IsMuted,
            IsSystemSession = session.IsSystemSession,
            IsRoutingSupported = session.IsRoutingSupported
        };
    }

    private static bool AreDevicesEquivalent(IReadOnlyList<AudioDevice>? left, IReadOnlyList<AudioDevice>? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left is null || right is null || left.Count != right.Count)
            return false;

        for (var index = 0; index < left.Count; index++)
        {
            var leftDevice = left[index];
            var rightDevice = right[index];
            if (!string.Equals(leftDevice.Id, rightDevice.Id, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(leftDevice.Name, rightDevice.Name, StringComparison.Ordinal) ||
                leftDevice.Flow != rightDevice.Flow ||
                leftDevice.IsDefault != rightDevice.IsDefault)
            {
                return false;
            }
        }

        return true;
    }

    private static bool AreAppSessionsEquivalent(MixerAppSessionInfo left, MixerAppSessionInfo right)
    {
        return string.Equals(left.SessionKey, right.SessionKey, StringComparison.OrdinalIgnoreCase) &&
            AreSessionEquivalent(left.OutputSession, right.OutputSession) &&
            AreSessionEquivalent(left.InputSession, right.InputSession);
    }

    private static bool AreSessionEquivalent(MixerSessionInfo? left, MixerSessionInfo? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left is null || right is null)
            return false;

        return string.Equals(left.SessionKey, right.SessionKey, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(left.DisplayName, right.DisplayName, StringComparison.Ordinal) &&
            string.Equals(left.ActualDeviceSummary, right.ActualDeviceSummary, StringComparison.Ordinal) &&
            string.Equals(left.BoundDeviceSummary, right.BoundDeviceSummary, StringComparison.Ordinal) &&
            string.Equals(left.ProcessName, right.ProcessName, StringComparison.Ordinal) &&
            string.Equals(left.ExecutablePath, right.ExecutablePath, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(left.BoundDeviceId, right.BoundDeviceId, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(left.RoutingUnavailableReason, right.RoutingUnavailableReason, StringComparison.Ordinal) &&
            left.Flow == right.Flow &&
            left.ProcessId == right.ProcessId &&
            Math.Abs(left.Volume - right.Volume) < 0.001f &&
            left.IsMuted == right.IsMuted &&
            left.IsSystemSession == right.IsSystemSession &&
            left.IsRoutingSupported == right.IsRoutingSupported;
    }

    private void ReorderSessionCardsFromSnapshot()
    {
        if (lastSnapshot is null)
            return;

        var desiredCards = new List<SessionCardControl>(lastSnapshot.Sessions.Count);
        foreach (var session in lastSnapshot.Sessions)
        {
            if (sessionCards.TryGetValue(session.SessionKey, out var card))
                desiredCards.Add(card);
        }

        ReplaceSessionHostChildren(desiredCards);
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
        if (SessionHost.Children.Count == 0)
        {
            foreach (var card in desiredCards)
                SessionHost.Children.Add(card);

            return;
        }

        if (SessionHost.Children.OfType<UIElement>().Any(child => child is not SessionCardControl))
            SessionHost.Children.Clear();

        var desiredLookup = new HashSet<SessionCardControl>(desiredCards);

        for (var index = SessionHost.Children.Count - 1; index >= 0; index--)
        {
            if (SessionHost.Children[index] is not SessionCardControl existingCard || !desiredLookup.Contains(existingCard))
                SessionHost.Children.RemoveAt(index);
        }

        for (var index = 0; index < desiredCards.Count; index++)
        {
            var desiredCard = desiredCards[index];
            if (index < SessionHost.Children.Count && ReferenceEquals(SessionHost.Children[index], desiredCard))
                continue;

            var existingIndex = FindSessionHostChildIndex(desiredCard);
            if (existingIndex >= 0)
            {
                SessionHost.Children.RemoveAt(existingIndex);
                SessionHost.Children.Insert(index, desiredCard);
            }
            else
            {
                SessionHost.Children.Insert(index, desiredCard);
            }
        }

        while (SessionHost.Children.Count > desiredCards.Count)
            SessionHost.Children.RemoveAt(SessionHost.Children.Count - 1);
    }

    private int FindSessionHostChildIndex(SessionCardControl card)
    {
        for (var index = 0; index < SessionHost.Children.Count; index++)
        {
            if (ReferenceEquals(SessionHost.Children[index], card))
                return index;
        }

        return -1;
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
        public const uint EventSystemForeground = 0x0003;
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
        public const uint WineventOutofcontext = 0x0000;
        public const uint WineventSkipOwnProcess = 0x0002;
        public const int ObjIdWindow = 0;
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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWinEventHook(
            uint eventMin,
            uint eventMax,
            IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc,
            uint idProcess,
            uint idThread,
            uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

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

    [Flags]
    private enum RefreshSessionScope
    {
        None = 0,
        Render = 1,
        Capture = 2,
        All = Render | Capture
    }

    private sealed record PanelSnapshot(IReadOnlyList<AudioDevice> Devices, IReadOnlyList<MixerAppSessionInfo> Sessions);

    private delegate void WinEventDelegate(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime);

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}
