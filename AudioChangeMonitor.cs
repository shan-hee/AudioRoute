using System;
using System.Collections.Generic;
using System.Diagnostics;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudioAudioSessionState = NAudio.CoreAudioApi.Interfaces.AudioSessionState;
using NAudioDataFlow = NAudio.CoreAudioApi.DataFlow;
using NAudioDeviceState = NAudio.CoreAudioApi.DeviceState;
using NAudioPropertyKey = NAudio.CoreAudioApi.PropertyKey;

namespace AudioRoute;

internal sealed class ObservedSessionVolumeChangedEventArgs : EventArgs
{
    public ObservedSessionVolumeChangedEventArgs(string sessionKey, EDataFlow flow, float volume, bool isMuted)
    {
        SessionKey = sessionKey;
        Flow = flow;
        Volume = volume;
        IsMuted = isMuted;
    }

    public string SessionKey { get; }

    public EDataFlow Flow { get; }

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
    private readonly object syncRoot = new();
    private readonly MMDeviceEnumerator deviceEnumerator = new();
    private readonly List<DeviceRegistration> deviceRegistrations = new();
    private readonly Dictionary<string, SessionRegistration> sessionRegistrations = new(StringComparer.OrdinalIgnoreCase);
    private DefaultRenderEndpointRegistration? defaultRenderEndpointRegistration;
    private bool disposed;

    public AudioChangeMonitor()
    {
        deviceEnumerator.RegisterEndpointNotificationCallback(this);
        RebuildSubscriptions();
    }

    public event EventHandler? Changed;

    public event EventHandler<ObservedSessionVolumeChangedEventArgs>? SessionVolumeChanged;

    public event EventHandler<ObservedMasterVolumeChangedEventArgs>? MasterVolumeChanged;

    public event EventHandler<ObservedSessionDisplayNameChangedEventArgs>? SessionDisplayNameChanged;

    public event EventHandler<ObservedSessionStructureChangedEventArgs>? SessionStructureChanged;

    public bool TrySetSessionVolume(string sessionKey, EDataFlow flow, float volume)
    {
        if (disposed)
            return false;

        List<SessionRegistration>? matchingRegistrations = null;
        var clampedVolume = Math.Clamp(volume, 0f, 1f);

        lock (syncRoot)
        {
            foreach (var registration in sessionRegistrations.Values)
            {
                if (registration.Flow != flow ||
                    !string.Equals(registration.SessionKey, sessionKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                matchingRegistrations ??= new List<SessionRegistration>();
                matchingRegistrations.Add(registration);
            }
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
                    simpleAudioVolume.Volume = clampedVolume;
                    updated = true;
                }
                finally
                {
                    simpleAudioVolume.Dispose();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[AudioRoute] Failed to set tracked session volume: key={sessionKey}, flow={flow}, {ex}");
            }
        }

        return updated;
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;

        try
        {
            deviceEnumerator.UnregisterEndpointNotificationCallback(this);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] Failed to unregister endpoint notifications: {ex}");
        }

        lock (syncRoot)
            ClearRegistrationsCore();

        deviceEnumerator.Dispose();
    }

    void IMMNotificationClient.OnDefaultDeviceChanged(NAudioDataFlow flow, Role role, string defaultDeviceId)
    {
        Trace.WriteLine($"[AudioRoute] Default device changed: flow={flow}, role={role}, id={defaultDeviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceAdded(string pwstrDeviceId)
    {
        Trace.WriteLine($"[AudioRoute] Device added: id={pwstrDeviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceRemoved(string deviceId)
    {
        Trace.WriteLine($"[AudioRoute] Device removed: id={deviceId}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnDeviceStateChanged(string deviceId, NAudioDeviceState newState)
    {
        Trace.WriteLine($"[AudioRoute] Device state changed: id={deviceId}, state={newState}");
        HandleDeviceTopologyChanged();
    }

    void IMMNotificationClient.OnPropertyValueChanged(string pwstrDeviceId, NAudioPropertyKey key)
    {
        Trace.WriteLine($"[AudioRoute] Device property changed: id={pwstrDeviceId}");
        HandleDeviceTopologyChanged();
    }

    private void HandleDeviceTopologyChanged()
    {
        if (disposed)
            return;

        DeviceEnumerator.InvalidateCache();

        lock (syncRoot)
            RebuildSubscriptionsCore();

        RaiseChanged();
    }

    private void RebuildSubscriptions()
    {
        lock (syncRoot)
            RebuildSubscriptionsCore();
    }

    private void RebuildSubscriptionsCore()
    {
        ClearRegistrationsCore();
        RebuildDefaultRenderEndpointRegistrationCore();
        RegisterFlowCore(NAudioDataFlow.Render);
        RegisterFlowCore(NAudioDataFlow.Capture);
    }

    private void RebuildDefaultRenderEndpointRegistrationCore()
    {
        defaultRenderEndpointRegistration?.Dispose();
        defaultRenderEndpointRegistration = null;

        try
        {
            var device = deviceEnumerator.GetDefaultAudioEndpoint(NAudioDataFlow.Render, Role.Multimedia);
            var endpointVolume = device.AudioEndpointVolume;
            AudioEndpointVolumeNotificationDelegate notificationHandler = notification =>
                HandleMasterVolumeNotification(notification.MasterVolume, notification.Muted);

            endpointVolume.OnVolumeNotification += notificationHandler;
            defaultRenderEndpointRegistration = new DefaultRenderEndpointRegistration(device, endpointVolume, notificationHandler);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] Failed to register default render endpoint volume watcher: {ex}");
        }

        RaiseMasterVolumeChanged(MasterVolumeService.TryGetMasterVolumeState());
    }

    private void RegisterFlowCore(NAudioDataFlow flow)
    {
        MMDeviceCollection? devices = null;
        var appFlow = ToAppFlow(flow);

        try
        {
            devices = deviceEnumerator.EnumerateAudioEndPoints(flow, NAudioDeviceState.Active);
            for (var index = 0; index < devices.Count; index++)
            {
                MMDevice? device = null;
                AudioSessionManager? sessionManager = null;
                SessionCollection? sessions = null;

                try
                {
                    device = devices[index];
                    sessionManager = device.AudioSessionManager;

                    AudioSessionManager.SessionCreatedDelegate sessionCreatedHandler =
                        (sender, newSessionControl) => HandleSessionCreated(appFlow, sender, newSessionControl);
                    var deviceRegistration = new DeviceRegistration(device, sessionManager, sessionCreatedHandler);
                    sessionManager.OnSessionCreated += deviceRegistration.SessionCreatedHandler;
                    sessionManager.RefreshSessions();

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
                            if (TryTrackSessionCore(session, appFlow))
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
                    Trace.WriteLine($"[AudioRoute] Failed to register audio session watchers: flow={flow}, index={index}, {ex}");
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
            Trace.WriteLine($"[AudioRoute] Failed to enumerate audio session watcher targets: flow={flow}, {ex}");
        }
    }

    private bool TryTrackSessionCore(AudioSessionControl session, EDataFlow flow)
    {
        if (disposed)
            return false;

        var sessionId = GetSessionIdentity(session);
        if (sessionRegistrations.ContainsKey(sessionId))
            return false;

        var sessionKey = GetAppSessionKey(session);
        var handler = new SessionEventsHandler(
            onSessionStateChanged: state => HandleSessionStateChanged(sessionId, flow, state),
            onSessionDisconnected: () => RemoveSessionAndRaiseStructureChanged(sessionId, flow),
            onDisplayNameChanged: displayName => HandleSessionDisplayNameChanged(sessionKey, flow, displayName),
            onVolumeChanged: (volume, isMuted) => HandleSessionVolumeChanged(sessionKey, flow, volume, isMuted));

        session.RegisterEventClient(handler);
        sessionRegistrations.Add(sessionId, new SessionRegistration(sessionId, sessionKey, flow, session, handler));
        return true;
    }

    private void HandleSessionCreated(EDataFlow flow, object? sender, IAudioSessionControl newSessionControl)
    {
        AudioSessionControl? session = null;
        var shouldRaise = false;

        try
        {
            session = new AudioSessionControl(newSessionControl);

            lock (syncRoot)
            {
                if (TryTrackSessionCore(session, flow))
                {
                    session = null;
                    shouldRaise = true;
                }
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[AudioRoute] Failed to track new audio session: {ex}");
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
            RemoveSessionAndRaiseStructureChanged(sessionId, flow);
    }

    private void HandleSessionVolumeChanged(string sessionKey, EDataFlow flow, float volume, bool isMuted)
    {
        if (disposed)
            return;

        SessionVolumeChanged?.Invoke(this, new ObservedSessionVolumeChangedEventArgs(sessionKey, flow, volume, isMuted));
    }

    private void HandleSessionDisplayNameChanged(string sessionKey, EDataFlow flow, string? displayName)
    {
        if (disposed)
            return;

        SessionDisplayNameChanged?.Invoke(this, new ObservedSessionDisplayNameChangedEventArgs(sessionKey, flow, displayName));
    }

    private void HandleMasterVolumeNotification(float volume, bool isMuted)
    {
        if (disposed)
            return;

        RaiseMasterVolumeChanged(MasterVolumeService.CreateState(volume, isMuted));
    }

    private void RemoveSessionAndRaiseStructureChanged(string sessionId, EDataFlow flow)
    {
        lock (syncRoot)
            RemoveSessionCore(sessionId);

        RaiseSessionStructureChanged(flow);
    }

    private void RemoveSessionCore(string sessionId)
    {
        if (!sessionRegistrations.Remove(sessionId, out var registration))
            return;

        registration.Dispose();
    }

    private void ClearRegistrationsCore()
    {
        defaultRenderEndpointRegistration?.Dispose();
        defaultRenderEndpointRegistration = null;

        foreach (var registration in sessionRegistrations.Values)
            registration.Dispose();

        sessionRegistrations.Clear();

        foreach (var registration in deviceRegistrations)
            registration.Dispose();

        deviceRegistrations.Clear();
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
        public DeviceRegistration(MMDevice device, AudioSessionManager sessionManager, AudioSessionManager.SessionCreatedDelegate sessionCreatedHandler)
        {
            Device = device;
            SessionManager = sessionManager;
            SessionCreatedHandler = sessionCreatedHandler;
        }

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
                Trace.WriteLine($"[AudioRoute] Failed to unregister session-created listener: {ex}");
            }

            SessionManager.Dispose();
            Device.Dispose();
        }
    }

    private sealed class SessionRegistration : IDisposable
    {
        public SessionRegistration(string sessionId, string sessionKey, EDataFlow flow, AudioSessionControl session, SessionEventsHandler handler)
        {
            SessionId = sessionId;
            SessionKey = sessionKey;
            Flow = flow;
            Session = session;
            Handler = handler;
        }

        public string SessionId { get; }

        public string SessionKey { get; }

        public EDataFlow Flow { get; }

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
                Trace.WriteLine($"[AudioRoute] Failed to unregister session event listener: session={SessionId}, {ex}");
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
                Trace.WriteLine($"[AudioRoute] Failed to unregister endpoint volume listener: {ex}");
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
