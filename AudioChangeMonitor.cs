using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudioAudioSessionState = NAudio.CoreAudioApi.Interfaces.AudioSessionState;
using NAudioDataFlow = NAudio.CoreAudioApi.DataFlow;
using NAudioDeviceState = NAudio.CoreAudioApi.DeviceState;
using NAudioPropertyKey = NAudio.CoreAudioApi.PropertyKey;

namespace AudioRoute;

internal sealed class ObservedSessionVolumeChangedEventArgs : EventArgs
{
    public ObservedSessionVolumeChangedEventArgs(string sessionKey, EDataFlow flow, string deviceId, float volume, bool isMuted)
    {
        SessionKey = sessionKey;
        Flow = flow;
        DeviceId = deviceId;
        Volume = volume;
        IsMuted = isMuted;
    }

    public string SessionKey { get; }

    public EDataFlow Flow { get; }

    public string DeviceId { get; }

    public float Volume { get; }

    public bool IsMuted { get; }
}

internal sealed class ObservedMasterVolumeChangedEventArgs : EventArgs
{
    public ObservedMasterVolumeChangedEventArgs(MasterVolumeState? state)
    {
        State = state;
    }

    public MasterVolumeState? State { get; }
}

internal sealed class ObservedSessionDisplayNameChangedEventArgs : EventArgs
{
    public ObservedSessionDisplayNameChangedEventArgs(string sessionKey, EDataFlow flow, string? displayName)
    {
        SessionKey = sessionKey;
        Flow = flow;
        DisplayName = displayName;
    }

    public string SessionKey { get; }

    public EDataFlow Flow { get; }

    public string? DisplayName { get; }
}

internal sealed class ObservedSessionStructureChangedEventArgs : EventArgs
{
    public ObservedSessionStructureChangedEventArgs(EDataFlow flow)
    {
        Flow = flow;
    }

    public EDataFlow Flow { get; }
}

internal sealed class AudioChangeMonitor : IMMNotificationClient, IDisposable
{
    private const int TopologyRebuildDebounceMilliseconds = 250;
    private readonly StaThreadDispatcher monitorDispatcher;
    private readonly Timer topologyRebuildTimer;
    private readonly List<DeviceRegistration> deviceRegistrations = new();
    private readonly Dictionary<string, SessionRegistration> sessionRegistrations = new(StringComparer.OrdinalIgnoreCase);
    private DefaultRenderEndpointRegistration? defaultRenderEndpointRegistration;
    private MMDeviceEnumerator? deviceEnumerator;
    private volatile bool disposed;
    private int topologyRebuildPending;
    private int topologyWorkQueued;

    public AudioChangeMonitor()
    {
        monitorDispatcher = new StaThreadDispatcher("AudioRoute.Monitor");
        topologyRebuildTimer = new Timer(
            _ => QueuePendingTopologyRebuild(),
            null,
            Timeout.Infinite,
            Timeout.Infinite);
        var initialization = monitorDispatcher.InvokeAsync(InitializeMonitorCore);
        _ = initialization.ContinueWith(
            completedWork =>
            {
                if (completedWork.Exception is not null)
                    RuntimeLog.WriteException("初始化音频监听器失败", completedWork.Exception.GetBaseException());
            },
            CancellationToken.None,
            TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    public event EventHandler? Changed;

    public event EventHandler<ObservedSessionVolumeChangedEventArgs>? SessionVolumeChanged;

    public event EventHandler<ObservedMasterVolumeChangedEventArgs>? MasterVolumeChanged;

    public event EventHandler<ObservedSessionDisplayNameChangedEventArgs>? SessionDisplayNameChanged;

    public event EventHandler<ObservedSessionStructureChangedEventArgs>? SessionStructureChanged;

    public bool TrySetSessionState(
        string sessionKey,
        EDataFlow flow,
        string? deviceId,
        float volume,
        bool isMuted)
    {
        if (disposed)
            return false;

        var clampedVolume = Math.Clamp(volume, 0f, 1f);

        return monitorDispatcher.Invoke(() =>
        {
            if (disposed)
                return false;

            if (TrySetSessionStateCore(sessionKey, flow, deviceId, clampedVolume, isMuted))
                return true;

            RebuildSubscriptions();
            return TrySetSessionStateCore(sessionKey, flow, deviceId, clampedVolume, isMuted);
        });
    }

    private bool TrySetSessionStateCore(
        string sessionKey,
        EDataFlow flow,
        string? deviceId,
        float volume,
        bool isMuted)
    {
        List<SessionRegistration>? matchingRegistrations = null;

        foreach (var registration in sessionRegistrations.Values)
        {
            if (registration.Flow != flow ||
                !string.Equals(registration.SessionKey, sessionKey, StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(deviceId) &&
                 !string.Equals(registration.DeviceId, deviceId, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            matchingRegistrations ??= new List<SessionRegistration>();
            matchingRegistrations.Add(registration);
        }

        if (matchingRegistrations is null || matchingRegistrations.Count == 0)
            return false;

        var updated = false;
        foreach (var registration in matchingRegistrations)
        {
            try
            {
                var simpleAudioVolume = registration.Session.SimpleAudioVolume;
                try
                {
                    simpleAudioVolume.Volume = volume;
                    simpleAudioVolume.Mute = isMuted;

                    updated = true;
                }
                finally
                {
                    simpleAudioVolume.Dispose();
                }
            }
            catch (Exception ex)
            {
                RuntimeLog.WriteException(
                    $"设置已监听会话状态失败: key={sessionKey}, flow={flow}, device={registration.DeviceId}",
                    ex);
            }
        }

        return updated;
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;
        topologyRebuildTimer.Dispose();

        try
        {
            var cleanup = monitorDispatcher.InvokeAsync(DisposeMonitorCore);
            _ = cleanup.ContinueWith(
                completedWork =>
                {
                    if (completedWork.Exception is not null)
                        RuntimeLog.WriteException("释放音频监听线程资源失败", completedWork.Exception.GetBaseException());
                },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("投递音频监听线程清理任务失败", ex);
        }
        finally
        {
            monitorDispatcher.Dispose();
        }
    }

    private void DisposeMonitorCore()
    {
        var enumerator = deviceEnumerator;
        deviceEnumerator = null;
        if (enumerator is null)
            return;

        try
        {
            enumerator.UnregisterEndpointNotificationCallback(this);
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("注销设备端点通知失败", ex);
        }

        DefaultRenderEndpointRegistration? previousDefaultRegistration;
        List<SessionRegistration> previousSessionRegistrations;
        List<DeviceRegistration> previousDeviceRegistrations;
        DetachRegistrationsCore(
            out previousDefaultRegistration,
            out previousSessionRegistrations,
            out previousDeviceRegistrations);

        DisposeRegistrations(previousDefaultRegistration, previousSessionRegistrations, previousDeviceRegistrations);

        enumerator.Dispose();
    }

    private void InitializeMonitorCore()
    {
        if (disposed)
            return;

        var enumerator = new MMDeviceEnumerator();
        try
        {
            enumerator.RegisterEndpointNotificationCallback(this);
            deviceEnumerator = enumerator;
            RebuildSubscriptions();
        }
        catch
        {
            deviceEnumerator = enumerator;
            DisposeMonitorCore();
            throw;
        }
    }

    void IMMNotificationClient.OnDefaultDeviceChanged(NAudioDataFlow flow, Role role, string defaultDeviceId)
    {
        RuntimeLog.Write($"默认设备变化: flow={flow}, role={role}, id={defaultDeviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceAdded(string pwstrDeviceId)
    {
        RuntimeLog.Write($"设备新增: id={pwstrDeviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceRemoved(string deviceId)
    {
        RuntimeLog.Write($"设备移除: id={deviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceStateChanged(string deviceId, NAudioDeviceState newState)
    {
        RuntimeLog.Write($"设备状态变化: id={deviceId}, state={newState}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnPropertyValueChanged(string pwstrDeviceId, NAudioPropertyKey key)
    {
        RuntimeLog.Write($"设备属性变化: id={pwstrDeviceId}, key={key}");
    }

    private void HandleDeviceTopologyChanged()
    {
        if (disposed)
            return;

        RuntimeLog.Write("监听器拓扑变化: queued");
        DeviceEnumerator.InvalidateCache();
        Interlocked.Exchange(ref topologyRebuildPending, 1);
        ScheduleTopologyRebuild();
    }

    private void ScheduleTopologyRebuild()
    {
        if (disposed)
            return;

        try
        {
            topologyRebuildTimer.Change(TopologyRebuildDebounceMilliseconds, Timeout.Infinite);
        }
        catch (ObjectDisposedException)
        {
        }
    }

    private void QueuePendingTopologyRebuild()
    {
        if (disposed || Interlocked.Exchange(ref topologyWorkQueued, 1) != 0)
            return;

        try
        {
            var work = monitorDispatcher.InvokeAsync(ProcessPendingTopologyRebuild);
            _ = work.ContinueWith(
                completedWork =>
                {
                    Interlocked.Exchange(ref topologyWorkQueued, 0);

                    if (completedWork.Exception is not null)
                        RuntimeLog.WriteException("音频监听器后台重建失败", completedWork.Exception.GetBaseException());

                    if (!disposed && Volatile.Read(ref topologyRebuildPending) != 0)
                        ScheduleTopologyRebuild();
                },
                CancellationToken.None,
                TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }
        catch (Exception ex)
        {
            Interlocked.Exchange(ref topologyWorkQueued, 0);
            RuntimeLog.WriteException("投递音频监听器重建任务失败", ex);
        }
    }

    private void ProcessPendingTopologyRebuild()
    {
        if (disposed || Interlocked.Exchange(ref topologyRebuildPending, 0) == 0)
            return;

        var startedAt = Stopwatch.GetTimestamp();
        RuntimeLog.Write("监听器拓扑变化: begin");
        RebuildSubscriptions();
        RaiseChanged();
        RuntimeLog.Write($"监听器拓扑变化: done, elapsed={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}ms");
    }

    private void RebuildSubscriptions()
    {
        if (disposed || deviceEnumerator is null)
            return;

        DefaultRenderEndpointRegistration? previousDefaultRegistration;
        List<SessionRegistration> previousSessionRegistrations;
        List<DeviceRegistration> previousDeviceRegistrations;

        DetachRegistrationsCore(
            out previousDefaultRegistration,
            out previousSessionRegistrations,
            out previousDeviceRegistrations);

        DisposeRegistrations(previousDefaultRegistration, previousSessionRegistrations, previousDeviceRegistrations);

        if (disposed)
            return;

        RuntimeLog.Write("监听器重建订阅: rebuilding");
        RebuildSubscriptionsCore();
        RuntimeLog.Write("监听器重建订阅: rebuilt");
    }

    private void RebuildSubscriptionsCore()
    {
        RebuildDefaultRenderEndpointRegistrationCore();
        RegisterFlowCore(NAudioDataFlow.Render);
        RegisterFlowCore(NAudioDataFlow.Capture);
    }

    private void RebuildDefaultRenderEndpointRegistrationCore()
    {
        DefaultRenderEndpointRegistration? nextRegistration = null;

        try
        {
            var device = deviceEnumerator!.GetDefaultAudioEndpoint(NAudioDataFlow.Render, Role.Multimedia);
            var endpointVolume = device.AudioEndpointVolume;
            var deviceId = device.ID;
            var deviceName = device.FriendlyName;
            AudioEndpointVolumeNotificationDelegate notificationHandler = notification =>
                HandleMasterVolumeNotification(
                    notification.MasterVolume,
                    notification.Muted,
                    deviceId,
                    deviceName);

            endpointVolume.OnVolumeNotification += notificationHandler;
            nextRegistration = new DefaultRenderEndpointRegistration(device, endpointVolume, notificationHandler);
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("注册默认输出设备音量监听器失败", ex);
        }

        if (disposed)
        {
            nextRegistration?.Dispose();
            return;
        }

        defaultRenderEndpointRegistration = nextRegistration;
        nextRegistration = null;

        nextRegistration?.Dispose();

        RaiseMasterVolumeChanged(MasterVolumeService.TryGetMasterVolumeState());
    }

    private void RegisterFlowCore(NAudioDataFlow flow)
    {
        MMDeviceCollection? devices = null;
        var appFlow = ToAppFlow(flow);

        try
        {
            devices = deviceEnumerator!.EnumerateAudioEndPoints(flow, NAudioDeviceState.Active);
            for (var index = 0; index < devices.Count; index++)
            {
                MMDevice? device = null;
                AudioSessionManager? sessionManager = null;
                SessionCollection? sessions = null;

                try
                {
                    device = devices[index];
                    var deviceId = device.ID;
                    sessionManager = device.AudioSessionManager;

                    AudioSessionManager.SessionCreatedDelegate sessionCreatedHandler =
                        (_, newSessionControl) => QueueSessionCreated(appFlow, deviceId, newSessionControl);
                    var deviceRegistration = new DeviceRegistration(deviceId, device, sessionManager, sessionCreatedHandler);
                    sessionManager.OnSessionCreated += deviceRegistration.SessionCreatedHandler;
                    sessionManager.RefreshSessions();

                    if (disposed)
                    {
                        deviceRegistration.Dispose();
                        device = null;
                        sessionManager = null;
                        continue;
                    }

                    deviceRegistrations.Add(deviceRegistration);

                    device = null;
                    sessionManager = null;

                    sessions = deviceRegistration.SessionManager.Sessions;
                    for (var sessionIndex = 0; sessionIndex < sessions.Count; sessionIndex++)
                    {
                        AudioSessionControl? session = null;

                        try
                        {
                            session = sessions[sessionIndex];
                            if (TryTrackSessionCore(session, appFlow, deviceRegistration.DeviceId))
                                session = null;
                        }
                        finally
                        {
                            session?.Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    RuntimeLog.WriteException($"注册音频会话监听器失败: flow={flow}, index={index}", ex);
                }
                finally
                {
                    sessionManager?.Dispose();
                    device?.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException($"枚举音频会话监听目标失败: flow={flow}", ex);
        }
    }

    private bool TryTrackSessionCore(AudioSessionControl session, EDataFlow flow, string deviceId)
    {
        if (disposed)
            return false;

        var sessionId = $"{deviceId}|{GetSessionIdentity(session)}";
        var sessionKey = GetAppSessionKey(session);

        if (disposed || sessionRegistrations.ContainsKey(sessionId))
            return false;

        var handler = new SessionEventsHandler(
            onSessionStateChanged: state => HandleSessionStateChanged(sessionId, flow, state),
            onSessionDisconnected: () => QueueSessionRemoval(sessionId, flow),
            onDisplayNameChanged: displayName => HandleSessionDisplayNameChanged(sessionKey, flow, displayName),
            onVolumeChanged: (volume, isMuted) => HandleSessionVolumeChanged(sessionKey, flow, deviceId, volume, isMuted));

        session.RegisterEventClient(handler);
        var registration = new SessionRegistration(sessionId, sessionKey, flow, deviceId, session, handler);
        if (!disposed && !sessionRegistrations.ContainsKey(sessionId))
        {
            sessionRegistrations.Add(sessionId, registration);
            return true;
        }

        registration.Dispose();
        return false;
    }

    private void QueueSessionCreated(EDataFlow flow, string deviceId, IAudioSessionControl newSessionControl)
    {
        if (disposed)
            return;

        try
        {
            var work = monitorDispatcher.InvokeAsync(() => TrackSessionCreatedCore(flow, deviceId, newSessionControl));
            _ = work.ContinueWith(
                completedWork =>
                {
                    if (completedWork.Exception is not null)
                        RuntimeLog.WriteException("后台监听新音频会话失败", completedWork.Exception.GetBaseException());
                },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("投递新音频会话监听任务失败", ex);
        }
    }

    private void TrackSessionCreatedCore(EDataFlow flow, string deviceId, IAudioSessionControl newSessionControl)
    {
        AudioSessionControl? session = null;
        var shouldRaise = false;

        try
        {
            session = new AudioSessionControl(newSessionControl);

            if (TryTrackSessionCore(session, flow, deviceId))
            {
                session = null;
                shouldRaise = true;
            }
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("监听新音频会话失败", ex);
            shouldRaise = true;
        }
        finally
        {
            session?.Dispose();
        }

        if (shouldRaise)
            RaiseSessionStructureChanged(flow);
    }

    private void HandleSessionStateChanged(string sessionId, EDataFlow flow, NAudioAudioSessionState state)
    {
        if (state == NAudioAudioSessionState.AudioSessionStateExpired)
            QueueSessionRemoval(sessionId, flow);
    }

    private void HandleSessionVolumeChanged(string sessionKey, EDataFlow flow, string deviceId, float volume, bool isMuted)
    {
        if (disposed)
            return;

        SessionVolumeChanged?.Invoke(
            this,
            new ObservedSessionVolumeChangedEventArgs(sessionKey, flow, deviceId, volume, isMuted));
    }

    private void HandleSessionDisplayNameChanged(string sessionKey, EDataFlow flow, string? displayName)
    {
        if (disposed)
            return;

        SessionDisplayNameChanged?.Invoke(this, new ObservedSessionDisplayNameChangedEventArgs(sessionKey, flow, displayName));
    }

    private void HandleMasterVolumeNotification(
        float volume,
        bool isMuted,
        string deviceId,
        string deviceName)
    {
        if (disposed)
            return;

        RaiseMasterVolumeChanged(MasterVolumeService.CreateState(volume, isMuted, deviceId, deviceName));
    }

    private void RemoveSessionAndRaiseStructureChanged(string sessionId, EDataFlow flow)
    {
        var removedRegistration = RemoveSessionCore(sessionId);

        removedRegistration?.Dispose();

        RaiseSessionStructureChanged(flow);
    }

    private void QueueSessionRemoval(string sessionId, EDataFlow flow)
    {
        if (disposed)
            return;

        try
        {
            var work = monitorDispatcher.InvokeAsync(() => RemoveSessionAndRaiseStructureChanged(sessionId, flow));
            _ = work.ContinueWith(
                completedWork =>
                {
                    if (completedWork.Exception is not null)
                        RuntimeLog.WriteException("后台移除音频会话监听器失败", completedWork.Exception.GetBaseException());
                },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException("投递音频会话移除任务失败", ex);
        }
    }

    private SessionRegistration? RemoveSessionCore(string sessionId)
    {
        if (!sessionRegistrations.Remove(sessionId, out var registration))
            return null;

        return registration;
    }

    private void DetachRegistrationsCore(
        out DefaultRenderEndpointRegistration? previousDefaultRegistration,
        out List<SessionRegistration> previousSessionRegistrations,
        out List<DeviceRegistration> previousDeviceRegistrations)
    {
        previousDefaultRegistration = defaultRenderEndpointRegistration;
        defaultRenderEndpointRegistration = null;

        previousSessionRegistrations = sessionRegistrations.Values.ToList();
        sessionRegistrations.Clear();

        previousDeviceRegistrations = deviceRegistrations.ToList();
        deviceRegistrations.Clear();
    }

    private static void DisposeRegistrations(
        DefaultRenderEndpointRegistration? previousDefaultRegistration,
        IReadOnlyList<SessionRegistration> previousSessionRegistrations,
        IReadOnlyList<DeviceRegistration> previousDeviceRegistrations)
    {
        var startedAt = Stopwatch.GetTimestamp();
        RuntimeLog.Write(
            $"监听器释放旧订阅: begin, endpoint={previousDefaultRegistration is not null}, " +
            $"sessions={previousSessionRegistrations.Count}, devices={previousDeviceRegistrations.Count}");

        RuntimeLog.Write("监听器释放默认端点订阅: begin");
        DisposeRegistration(previousDefaultRegistration, "释放默认端点订阅失败");
        RuntimeLog.Write("监听器释放默认端点订阅: done");

        RuntimeLog.Write($"监听器释放会话订阅: begin, count={previousSessionRegistrations.Count}");
        foreach (var registration in previousSessionRegistrations)
            DisposeRegistration(registration, $"释放会话订阅失败: session={registration.SessionId}");
        RuntimeLog.Write("监听器释放会话订阅: done");

        RuntimeLog.Write($"监听器释放设备订阅: begin, count={previousDeviceRegistrations.Count}");
        foreach (var registration in previousDeviceRegistrations)
            DisposeRegistration(registration, "释放设备会话创建订阅失败");
        RuntimeLog.Write("监听器释放设备订阅: done");

        RuntimeLog.Write($"监听器释放旧订阅: done, elapsed={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}ms");
    }

    private static void DisposeRegistration(IDisposable? registration, string failureContext)
    {
        if (registration is null)
            return;

        try
        {
            registration.Dispose();
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException(failureContext, ex);
        }
    }

    private void RaiseChanged()
    {
        if (disposed)
            return;

        Changed?.Invoke(this, EventArgs.Empty);
    }

    private void RaiseMasterVolumeChanged(MasterVolumeState? state)
    {
        if (disposed)
            return;

        MasterVolumeChanged?.Invoke(this, new ObservedMasterVolumeChangedEventArgs(state));
    }

    private void RaiseSessionStructureChanged(EDataFlow flow)
    {
        if (disposed)
            return;

        SessionStructureChanged?.Invoke(this, new ObservedSessionStructureChangedEventArgs(flow));
    }

    private static string GetSessionIdentity(AudioSessionControl session)
    {
        try
        {
            var sessionInstanceId = session.GetSessionInstanceIdentifier;
            if (!string.IsNullOrWhiteSpace(sessionInstanceId))
                return sessionInstanceId;
        }
        catch
        {
        }

        try
        {
            var sessionIdentifier = session.GetSessionIdentifier;
            if (!string.IsNullOrWhiteSpace(sessionIdentifier))
                return sessionIdentifier;
        }
        catch
        {
        }

        if (session.IsSystemSoundsSession)
            return "system";

        return $"pid:{session.GetProcessID}";
    }

    private static string GetAppSessionKey(AudioSessionControl session)
    {
        if (session.IsSystemSoundsSession)
            return "system";

        var processId = checked((int)session.GetProcessID);
        return processId <= 0 ? "system" : $"pid:{processId}";
    }

    private static EDataFlow ToAppFlow(NAudioDataFlow flow)
    {
        return flow switch
        {
            NAudioDataFlow.Render => EDataFlow.eRender,
            NAudioDataFlow.Capture => EDataFlow.eCapture,
            _ => EDataFlow.eAll
        };
    }

    private sealed class DeviceRegistration : IDisposable
    {
        public DeviceRegistration(
            string deviceId,
            MMDevice device,
            AudioSessionManager sessionManager,
            AudioSessionManager.SessionCreatedDelegate sessionCreatedHandler)
        {
            DeviceId = deviceId;
            Device = device;
            SessionManager = sessionManager;
            SessionCreatedHandler = sessionCreatedHandler;
        }

        public string DeviceId { get; }

        public MMDevice Device { get; }

        public AudioSessionManager SessionManager { get; }

        public AudioSessionManager.SessionCreatedDelegate SessionCreatedHandler { get; }

        public void Dispose()
        {
            try
            {
                SessionManager.OnSessionCreated -= SessionCreatedHandler;
            }
            catch (Exception ex)
            {
                RuntimeLog.WriteException("注销音频会话创建监听器失败", ex);
            }

            SessionManager.Dispose();
            Device.Dispose();
        }
    }

    private sealed class SessionRegistration : IDisposable
    {
        public SessionRegistration(
            string sessionId,
            string sessionKey,
            EDataFlow flow,
            string deviceId,
            AudioSessionControl session,
            SessionEventsHandler handler)
        {
            SessionId = sessionId;
            SessionKey = sessionKey;
            Flow = flow;
            DeviceId = deviceId;
            Session = session;
            Handler = handler;
        }

        public string SessionId { get; }

        public string SessionKey { get; }

        public EDataFlow Flow { get; }

        public string DeviceId { get; }

        public AudioSessionControl Session { get; }

        public SessionEventsHandler Handler { get; }

        public void Dispose()
        {
            try
            {
                Session.UnRegisterEventClient(Handler);
            }
            catch (Exception ex)
            {
                RuntimeLog.WriteException($"注销音频会话事件监听器失败: session={SessionId}", ex);
            }

            Session.Dispose();
        }
    }

    private sealed class DefaultRenderEndpointRegistration : IDisposable
    {
        public DefaultRenderEndpointRegistration(
            MMDevice device,
            AudioEndpointVolume endpointVolume,
            AudioEndpointVolumeNotificationDelegate notificationHandler)
        {
            Device = device;
            EndpointVolume = endpointVolume;
            NotificationHandler = notificationHandler;
        }

        public MMDevice Device { get; }

        public AudioEndpointVolume EndpointVolume { get; }

        public AudioEndpointVolumeNotificationDelegate NotificationHandler { get; }

        public void Dispose()
        {
            try
            {
                EndpointVolume.OnVolumeNotification -= NotificationHandler;
            }
            catch (Exception ex)
            {
                RuntimeLog.WriteException("注销默认端点音量监听器失败", ex);
            }

            EndpointVolume.Dispose();
            Device.Dispose();
        }
    }

    private sealed class SessionEventsHandler : IAudioSessionEventsHandler
    {
        private readonly Action<NAudioAudioSessionState> onSessionStateChanged;
        private readonly Action onSessionDisconnected;
        private readonly Action<string?> onDisplayNameChanged;
        private readonly Action<float, bool> onVolumeChanged;

        public SessionEventsHandler(
            Action<NAudioAudioSessionState> onSessionStateChanged,
            Action onSessionDisconnected,
            Action<string?> onDisplayNameChanged,
            Action<float, bool> onVolumeChanged)
        {
            this.onSessionStateChanged = onSessionStateChanged;
            this.onSessionDisconnected = onSessionDisconnected;
            this.onDisplayNameChanged = onDisplayNameChanged;
            this.onVolumeChanged = onVolumeChanged;
        }

        public void OnDisplayNameChanged(string displayName)
        {
            onDisplayNameChanged(displayName);
        }

        public void OnIconPathChanged(string iconPath)
        {
        }

        public void OnVolumeChanged(float volume, bool isMuted)
        {
            onVolumeChanged(volume, isMuted);
        }

        public void OnChannelVolumeChanged(uint channelCount, IntPtr newVolumes, uint channelIndex)
        {
        }

        public void OnGroupingParamChanged(ref Guid groupingId)
        {
        }

        public void OnStateChanged(NAudioAudioSessionState state)
        {
            onSessionStateChanged(state);
        }

        public void OnSessionDisconnected(AudioSessionDisconnectReason disconnectReason)
        {
            onSessionDisconnected();
        }
    }
}
