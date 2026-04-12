using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT.Interop;
using Windows.Graphics;

namespace AudioRoute;

public sealed partial class MainWindow : Window
{
    private const int PanelWidth = 400;
    private const int PanelHeight = 460;
    private static readonly TimeSpan TrayReopenSuppressionDuration = TimeSpan.FromMilliseconds(300);
    private static readonly TimeSpan SnapshotHealthCheckInterval = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan SnapshotRefreshOnOpenAge = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan SnapshotRefreshWhileVisibleAge = TimeSpan.FromMinutes(3);
    private static readonly TimeSpan AudioChangeRefreshDebounceInterval = TimeSpan.FromMilliseconds(350);
    private static readonly TimeSpan InitialSnapshotWarmupWait = TimeSpan.FromMilliseconds(160);
    private static readonly TimeSpan MasterVolumePollInterval = TimeSpan.FromMilliseconds(400);
    private static readonly TimeSpan MasterVolumePollQuietWindow = TimeSpan.FromMilliseconds(900);
    private static readonly Guid TrayIconGuid = new("05F26E47-7F95-4C1C-B0E6-4D748C6CFB6F");
    private const uint TrayIconId = 1;
    private const uint TrayCallbackMessage = NativeMethods.WmApp + 1;

    private readonly DispatcherQueueTimer refreshTimer;
    private readonly DispatcherQueueTimer audioChangeRefreshTimer;
    private readonly DispatcherQueueTimer masterVolumePollTimer;
    private readonly StaThreadDispatcher refreshDispatcher;
    private readonly StaThreadDispatcher trayDispatcher;
    private readonly AudioChangeMonitor audioChangeMonitor;
    private readonly ShellNotifyIconHost trayIconHost;
    private readonly TrayIconManager trayIconManager;
    private PanelController panelController = null!;
    private readonly Dictionary<string, SessionCardControl> sessionCards = new(StringComparer.OrdinalIgnoreCase);
    private readonly object masterVolumeUpdateSync = new();
    private readonly object sessionVolumeUpdateSync = new();
    private readonly object sessionVolumeCommitSync = new();
    private readonly Dictionary<string, ObservedSessionVolumeChangedEventArgs> pendingObservedSessionVolumeChanges = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, PendingSessionVolumeCommit> pendingSessionVolumeCommits = new(StringComparer.OrdinalIgnoreCase);
    private AppWindow? appWindow;
    private IntPtr hwnd;
    private bool isConfigured;
    private bool isExitRequested;
    private bool isRefreshing;
    private bool hasDeferredRefresh;
    private bool deferredRefreshCanReuseDevices;
    private RefreshSessionScope deferredRefreshScope = RefreshSessionScope.None;
    private bool isSnapshotStale = true;
    private MasterVolumeState? pendingObservedMasterVolumeState;
    private PanelSnapshot? lastSnapshot;
    private DateTimeOffset lastSnapshotUpdatedAt;
    private int interactionDepth;
    private int refreshGeneration;
    private bool scheduledRefreshCanReuseDevices;
    private RefreshSessionScope scheduledRefreshScope = RefreshSessionScope.None;
    private Task? latestRefreshTask;
    private bool isMasterVolumeUpdateQueued;
    private bool isSessionVolumeUpdateQueued;
    private bool isSessionVolumeCommitFlushRunning;
    private bool isMasterVolumePollInFlight;
    private long lastObservedMasterVolumeEventTick;
    private WndProcDelegate? windowProcDelegate;
    private IntPtr originalWindowProc;

    public MainWindow()
    {
        RuntimeLog.Reset();
        RuntimeLog.Write("应用启动");
        Interlocked.Exchange(ref lastObservedMasterVolumeEventTick, Environment.TickCount64);
        InitializeComponent();
        refreshDispatcher = new StaThreadDispatcher("AudioRoute.Refresh");
        trayDispatcher = new StaThreadDispatcher("AudioRoute.Tray");
        trayIconHost = new ShellNotifyIconHost(TrayIconGuid, TrayIconId, TrayCallbackMessage);
        trayIconManager = new TrayIconManager(() => hwnd, trayIconHost);
        trayIconManager.TogglePanelRequested += () => { _ = TogglePanelVisibilityAsync(); return Task.CompletedTask; };
        trayIconManager.ShowPanelRequested += () => { _ = ShowOrBringToFrontAsync(); return Task.CompletedTask; };
        trayIconManager.ExitRequested += RequestExit;
        trayIconManager.ErrorOccurred += ShowError;
        ConfigureWindow();

        Activated += OnWindowActivated;
        Closed += OnWindowClosed;

        refreshTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        refreshTimer.Interval = SnapshotHealthCheckInterval;
        refreshTimer.Tick += (_, _) =>
        {
            if (!isExitRequested && panelController.IsPanelVisible)
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

            if (!isExitRequested && (panelController.IsPanelVisible || lastSnapshot is not null))
                RefreshData(canReuseDevices, scope);
        };

        masterVolumePollTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        masterVolumePollTimer.Interval = MasterVolumePollInterval;
        masterVolumePollTimer.IsRepeating = true;
        masterVolumePollTimer.Tick += async (_, _) =>
        {
            if (isExitRequested || !trayIconManager.IsCreated || isMasterVolumePollInFlight)
                return;

            if (Environment.TickCount64 - Interlocked.Read(ref lastObservedMasterVolumeEventTick) < (long)MasterVolumePollQuietWindow.TotalMilliseconds)
                return;

            isMasterVolumePollInFlight = true;
            try
            {
                var polledState = await trayDispatcher.InvokeAsync(MasterVolumeService.TryGetMasterVolumeState);
                if (!isExitRequested && trayIconManager.IsCreated && trayIconManager.ShouldQueueTrayIconUpdate(polledState))
                {
                    RuntimeLog.Write($"主音量轮询命中: state={TrayIconManager.FormatMasterVolumeStateForLog(polledState)}");
                    trayIconManager.ScheduleTrayIconUpdate(polledState, force: true);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[AudioRoute] 主音量轮询失败: {ex}");
            }
            finally
            {
                isMasterVolumePollInFlight = false;
            }
        };

        audioChangeMonitor = new AudioChangeMonitor();
        audioChangeMonitor.Changed += OnAudioEnvironmentChanged;
        audioChangeMonitor.SessionStructureChanged += OnObservedSessionStructureChanged;
        audioChangeMonitor.SessionVolumeChanged += OnObservedSessionVolumeChanged;
        audioChangeMonitor.MasterVolumeChanged += OnObservedMasterVolumeChanged;
        audioChangeMonitor.SessionDisplayNameChanged += OnObservedSessionDisplayNameChanged;

        EnsureConfigured();
    }

    public bool IsPanelVisible => panelController.IsPanelVisible;

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
        panelController.AttachForegroundEventHook();
        trayIconManager.Initialize();
        masterVolumePollTimer.Start();
        RefreshData();
    }

    private void HandlePanelDeactivated()
    {
        panelController.HandlePanelDeactivated();
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

        lock (sessionVolumeUpdateSync)
        {
            pendingObservedSessionVolumeChanges[CreateSessionFlowKey(e.SessionKey, e.Flow)] = e;
            if (isSessionVolumeUpdateQueued)
                return;

            isSessionVolumeUpdateQueued = true;
        }

        if (!DispatcherQueue.TryEnqueue(DispatcherQueuePriority.High, FlushObservedSessionVolumeChanged))
        {
            lock (sessionVolumeUpdateSync)
            {
                isSessionVolumeUpdateQueued = false;
            }
        }
    }

    private void OnObservedMasterVolumeChanged(object? sender, ObservedMasterVolumeChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        Interlocked.Exchange(ref lastObservedMasterVolumeEventTick, Environment.TickCount64);

        lock (masterVolumeUpdateSync)
        {
            pendingObservedMasterVolumeState = e.State;
            if (isMasterVolumeUpdateQueued)
                return;

            isMasterVolumeUpdateQueued = true;
        }

        if (!DispatcherQueue.TryEnqueue(DispatcherQueuePriority.High, FlushObservedMasterVolumeChanged))
        {
            lock (masterVolumeUpdateSync)
            {
                isMasterVolumeUpdateQueued = false;
            }
        }
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

        if (!panelController.IsPanelVisible && lastSnapshot is null)
            return;

        ScheduleAudioRefresh(canReuseDevices: false, scope: RefreshSessionScope.All);
    }

    private void HandleObservedSessionStructureChanged(ObservedSessionStructureChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        isSnapshotStale = true;

        if (!panelController.IsPanelVisible && lastSnapshot is null)
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

        if (!panelController.IsPanelVisible)
            return;

        if (sessionCards.TryGetValue(e.SessionKey, out var card))
            card.ApplyObservedVolume(e.Flow, e.Volume, e.IsMuted);
    }

    private void FlushObservedSessionVolumeChanged()
    {
        List<ObservedSessionVolumeChangedEventArgs> pendingChanges;
        lock (sessionVolumeUpdateSync)
        {
            if (pendingObservedSessionVolumeChanges.Count == 0)
            {
                isSessionVolumeUpdateQueued = false;
                return;
            }

            pendingChanges = pendingObservedSessionVolumeChanges.Values.ToList();
            pendingObservedSessionVolumeChanges.Clear();
            isSessionVolumeUpdateQueued = false;
        }

        foreach (var change in pendingChanges)
            HandleObservedSessionVolumeChanged(change);
    }

    private void HandleObservedMasterVolumeChanged(ObservedMasterVolumeChangedEventArgs e)
    {
        if (isExitRequested || !trayIconManager.IsCreated)
            return;

        if (trayIconManager.ShouldQueueTrayIconUpdate(e.State))
            trayIconManager.ScheduleTrayIconUpdate(e.State, force: true);
    }

    private void FlushObservedMasterVolumeChanged()
    {
        MasterVolumeState? currentState;
        lock (masterVolumeUpdateSync)
        {
            currentState = pendingObservedMasterVolumeState;
            pendingObservedMasterVolumeState = null;
            isMasterVolumeUpdateQueued = false;
        }

        HandleObservedMasterVolumeChanged(new ObservedMasterVolumeChangedEventArgs(currentState));
    }

    private static string CreateSessionFlowKey(string sessionKey, EDataFlow flow)
    {
        return $"{(int)flow}:{sessionKey}";
    }

    private void HandleObservedSessionDisplayNameChanged(ObservedSessionDisplayNameChangedEventArgs e)
    {
        if (isExitRequested)
            return;

        if (string.IsNullOrWhiteSpace(e.DisplayName))
        {
            isSnapshotStale = true;

            if (!panelController.IsPanelVisible && lastSnapshot is null)
                return;

            ScheduleAudioRefresh(
                canReuseDevices: lastSnapshot is not null,
                scope: GetRefreshScope(e.Flow));
            return;
        }

        if (!TryUpdateCachedSessionDisplayName(e.SessionKey, e.Flow, e.DisplayName, out var updatedAppSession))
            return;

        if (!panelController.IsPanelVisible || lastSnapshot is null || updatedAppSession is null)
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

            if (panelController.IsPanelVisible)
                ApplyPanelSnapshot(snapshot, previousSnapshot);
        }
        catch when (isExitRequested)
        {
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] 刷新音频会话失败: {ex}");
            if (refreshVersion == refreshGeneration && panelController.IsPanelVisible)
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

            if (panelController.IsPanelVisible)
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
        var shouldStartFlush = false;
        lock (sessionVolumeCommitSync)
        {
            pendingSessionVolumeCommits[CreateSessionFlowKey(e.Session.SessionKey, e.Session.Flow)] =
                new PendingSessionVolumeCommit(e, sender as SessionCardControl);

            if (!isSessionVolumeCommitFlushRunning)
            {
                isSessionVolumeCommitFlushRunning = true;
                shouldStartFlush = true;
            }
        }

        if (shouldStartFlush)
            _ = FlushPendingSessionVolumeCommitsAsync();
    }

    private async Task FlushPendingSessionVolumeCommitsAsync()
    {
        while (!isExitRequested)
        {
            List<PendingSessionVolumeCommit> pendingCommits;
            lock (sessionVolumeCommitSync)
            {
                if (pendingSessionVolumeCommits.Count == 0)
                {
                    isSessionVolumeCommitFlushRunning = false;
                    return;
                }

                pendingCommits = pendingSessionVolumeCommits.Values.ToList();
                pendingSessionVolumeCommits.Clear();
            }

            SessionVolumeCommitBatchResult result;
            try
            {
                result = await refreshDispatcher.InvokeAsync(() =>
                {
                    var succeeded = new List<PendingSessionVolumeCommit>(pendingCommits.Count);
                    var failures = new List<SessionVolumeCommitFailure>();

                    foreach (var pendingCommit in pendingCommits)
                    {
                        try
                        {
                            var change = pendingCommit.Change;
                            var updated = false;
                            try
                            {
                                updated = audioChangeMonitor.TrySetSessionVolume(change.Session.SessionKey, change.Session.Flow, change.Volume);
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"[AudioRoute] 快速调节音量路径失败，回退到枚举路径: {ex}");
                            }

                            if (!updated)
                                AudioSessionService.SetSessionVolume(change.Session.SessionKey, change.Session.Flow, change.Volume);

                            succeeded.Add(pendingCommit);
                        }
                        catch (Exception ex)
                        {
                            failures.Add(new SessionVolumeCommitFailure(pendingCommit, ex));
                        }
                    }

                    return new SessionVolumeCommitBatchResult(succeeded, failures);
                });
            }
            catch (Exception ex)
            {
                result = new SessionVolumeCommitBatchResult(
                    Array.Empty<PendingSessionVolumeCommit>(),
                    pendingCommits.Select(commit => new SessionVolumeCommitFailure(commit, ex)).ToArray());
            }

            foreach (var succeededCommit in result.Succeeded)
            {
                var change = succeededCommit.Change;
                UpdateCachedSessionVolume(change.Session.SessionKey, change.Session.Flow, change.Volume, isMuted: change.Volume <= 0.005f);
            }

            if (result.Failures.Count == 0)
                continue;

            foreach (var failure in result.Failures)
            {
                failure.Commit.SourceCard?.NotifyVolumeCommitFailed();
                RuntimeLog.Write($"主页调节音量: failed session={failure.Commit.Change.Session.SessionKey}, flow={failure.Commit.Change.Session.Flow}, message={failure.Exception.Message}");
                Trace.WriteLine($"[AudioRoute] 调整音量失败: {failure.Exception}");
            }

            hasDeferredRefresh = true;
            RefreshData();
            ShowError($"调整音量失败: {result.Failures[0].Exception.Message}");
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
        panelController.CancelDeactivateHide();

        if (!isConfigured && hwnd == IntPtr.Zero)
            Activate();

        EnsureConfigured();

        if (!isConfigured || isExitRequested || panelController.IsVisibilityTransitioning)
            return Task.CompletedTask;

        if (panelController.IsPanelVisible)
        {
            panelController.BringPanelToFront(() => Activate());
            return Task.CompletedTask;
        }

        return ShowPanelAsync();
    }

    private Task TogglePanelVisibilityAsync()
    {
        panelController.CancelDeactivateHide();

        if (!isConfigured && hwnd == IntPtr.Zero)
            Activate();

        EnsureConfigured();

        if (!isConfigured || isExitRequested || panelController.IsVisibilityTransitioning)
            return Task.CompletedTask;

        return panelController.IsPanelVisible
            ? HideToTrayAsync()
            : ShowPanelAsync();
    }

    public void PrepareForExit()
    {
        if (isExitRequested)
            return;

        isExitRequested = true;
        panelController.StopTimers();
        refreshTimer.Stop();
        audioChangeRefreshTimer.Stop();
        panelController.DetachForegroundEventHook();
        audioChangeMonitor.Changed -= OnAudioEnvironmentChanged;
        audioChangeMonitor.SessionStructureChanged -= OnObservedSessionStructureChanged;
        audioChangeMonitor.SessionVolumeChanged -= OnObservedSessionVolumeChanged;
        audioChangeMonitor.MasterVolumeChanged -= OnObservedMasterVolumeChanged;
        audioChangeMonitor.SessionDisplayNameChanged -= OnObservedSessionDisplayNameChanged;
        audioChangeMonitor.Dispose();
        masterVolumePollTimer.Stop();
        isMasterVolumePollInFlight = false;
        lock (masterVolumeUpdateSync)
        {
            pendingObservedMasterVolumeState = null;
            isMasterVolumeUpdateQueued = false;
        }
        panelController.Dispose();
        trayIconManager.Dispose();
        DetachWindowProc();
        trayDispatcher.Dispose();
        refreshDispatcher.Dispose();
        TrayVolumeIconService.Dispose();
        AudioPolicyManager.Cleanup();
    }

    private async Task ShowPanelAsync()
    {
        if (!isConfigured || isExitRequested)
            return;

        await panelController.ShowPanelAsync(() => Activate());
    }

    private async Task HideToTrayAsync()
    {
        if (!isConfigured || isExitRequested)
            return;

        refreshTimer.Stop();
        await panelController.HideToTrayAsync();
        trayIconManager.SuppressReopenUntil(DateTimeOffset.UtcNow + TrayReopenSuppressionDuration);
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
        NativeMethods.ApplyWindowCornerPreference(hwnd);
        AttachWindowProc();
        appWindow.Resize(new SizeInt32(PanelWidth, PanelHeight));
        panelController = new PanelController(hwnd, appWindow, DispatcherQueue, PanelSurface, trayIconManager);
        trayIconManager.IsPanelVisible = () => isConfigured && panelController.IsPanelVisible;
        trayIconManager.CancelDeactivateHideRequested = () =>
        {
            if (isConfigured)
                panelController.CancelDeactivateHide();
        };
        panelController.IsExitRequested = () => isExitRequested;
        panelController.IsInteracting = () => interactionDepth > 0;
        panelController.OnBeforeShow = async () =>
        {
            await WaitForInitialSnapshotWarmupAsync();
            ShowCachedSnapshotOrLoading();
            refreshTimer.Start();
            WindowRoot.UpdateLayout();
            PanelSurface.UpdateLayout();
        };
        panelController.PanelShown += () => RefreshDataIfNeeded(SnapshotRefreshOnOpenAge);
        panelController.PanelHidden += () => trayIconManager.SuppressReopenUntil(DateTimeOffset.UtcNow + TrayReopenSuppressionDuration);
        panelController.UpdateWindowPosition();
        NativeMethods.SetTopMost(hwnd);
        NativeMethods.SetWindowCloaked(hwnd, true);
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

            var updatedSession = session with { Volume = volume, IsMuted = isMuted };
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

            var updatedSession = session with { BoundDeviceId = boundDeviceId, BoundDeviceSummary = boundDeviceSummary };
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

            var updatedSession = session with { DisplayName = displayName };
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
            left.OutputSession == right.OutputSession &&
            left.InputSession == right.InputSession;
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
            if (trayIconManager.TryHandleWmCommand(commandId))
                return IntPtr.Zero;
        }

        return originalWindowProc != IntPtr.Zero
            ? NativeMethods.CallWindowProc(originalWindowProc, windowHandle, message, wParam, lParam)
            : NativeMethods.DefWindowProc(windowHandle, message, wParam, lParam);
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
}
