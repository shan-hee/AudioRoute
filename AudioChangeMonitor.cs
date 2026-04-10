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

internal sealed class AudioChangeMonitor : IMMNotificationClient, IDisposable
{
    private readonly object syncRoot = new();
    private readonly MMDeviceEnumerator deviceEnumerator = new();
    private readonly List<DeviceRegistration> deviceRegistrations = new();
    private readonly Dictionary<string, SessionRegistration> sessionRegistrations = new(StringComparer.OrdinalIgnoreCase);
    private bool disposed;

    public AudioChangeMonitor()
    {
        deviceEnumerator.RegisterEndpointNotificationCallback(this);
        RebuildSubscriptions();
    }

    public event EventHandler? Changed;

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
        RegisterFlowCore(NAudioDataFlow.Render);
        RegisterFlowCore(NAudioDataFlow.Capture);
    }

    private void RegisterFlowCore(NAudioDataFlow flow)
    {
        MMDeviceCollection? devices = null;

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

                    var deviceRegistration = new DeviceRegistration(device, sessionManager, HandleSessionCreated);
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
                            if (TryTrackSessionCore(session))
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

    private bool TryTrackSessionCore(AudioSessionControl session)
    {
        if (disposed)
            return false;

        var sessionId = GetSessionIdentity(session);
        if (sessionRegistrations.ContainsKey(sessionId))
            return false;

        var handler = new SessionEventsHandler(
            onSessionStateChanged: state => HandleSessionStateChanged(sessionId, state),
            onSessionDisconnected: () => RemoveSessionAndRaise(sessionId),
            onMetadataChanged: RaiseChanged);

        session.RegisterEventClient(handler);
        sessionRegistrations.Add(sessionId, new SessionRegistration(sessionId, session, handler));
        return true;
    }

    private void HandleSessionCreated(object? sender, IAudioSessionControl newSessionControl)
    {
        AudioSessionControl? session = null;
        var shouldRaise = false;

        try
        {
            session = new AudioSessionControl(newSessionControl);

            lock (syncRoot)
            {
                if (TryTrackSessionCore(session))
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
            RaiseChanged();
    }

    private void HandleSessionStateChanged(string sessionId, NAudioAudioSessionState state)
    {
        if (state == NAudioAudioSessionState.AudioSessionStateExpired)
            RemoveSessionAndRaise(sessionId);
    }

    private void RemoveSessionAndRaise(string sessionId)
    {
        lock (syncRoot)
            RemoveSessionCore(sessionId);

        RaiseChanged();
    }

    private void RemoveSessionCore(string sessionId)
    {
        if (!sessionRegistrations.Remove(sessionId, out var registration))
            return;

        registration.Dispose();
    }

    private void ClearRegistrationsCore()
    {
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
        public SessionRegistration(string sessionId, AudioSessionControl session, SessionEventsHandler handler)
        {
            SessionId = sessionId;
            Session = session;
            Handler = handler;
        }

        public string SessionId { get; }

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

    private sealed class SessionEventsHandler : IAudioSessionEventsHandler
    {
        private readonly Action<NAudioAudioSessionState> onSessionStateChanged;
        private readonly Action onSessionDisconnected;
        private readonly Action onMetadataChanged;

        public SessionEventsHandler(Action<NAudioAudioSessionState> onSessionStateChanged, Action onSessionDisconnected, Action onMetadataChanged)
        {
            this.onSessionStateChanged = onSessionStateChanged;
            this.onSessionDisconnected = onSessionDisconnected;
            this.onMetadataChanged = onMetadataChanged;
        }

        public void OnDisplayNameChanged(string displayName)
        {
            onMetadataChanged();
        }

        public void OnIconPathChanged(string iconPath)
        {
            onMetadataChanged();
        }

        public void OnVolumeChanged(float volume, bool isMuted)
        {
        }

        public void OnChannelVolumeChanged(uint channelCount, IntPtr newVolumes, uint channelIndex)
        {
        }

        public void OnGroupingParamChanged(ref Guid groupingId)
        {
            onMetadataChanged();
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
