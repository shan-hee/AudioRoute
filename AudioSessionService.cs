using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NAudio.CoreAudioApi;
using NAudioDataFlow = NAudio.CoreAudioApi.DataFlow;
using NAudioDeviceState = NAudio.CoreAudioApi.DeviceState;

namespace AudioRoute;

public static class AudioSessionService
{
    public static IReadOnlyList<MixerSessionInfo> GetActiveSessions(EDataFlow flow)
    {
        var devices = DeviceEnumerator.EnumerateDevices(flow);
        var deviceMap = devices.ToDictionary(device => device.Id, StringComparer.OrdinalIgnoreCase);
        var defaultDeviceName = devices.FirstOrDefault(device => device.IsDefault)?.Name ?? "System Default";
        var aggregates = new Dictionary<string, SessionAggregate>(StringComparer.OrdinalIgnoreCase);

        using var enumerator = new MMDeviceEnumerator();
        var audioDevices = enumerator.EnumerateAudioEndPoints(ToNaudioFlow(flow), NAudioDeviceState.Active);

        for (var deviceIndex = 0; deviceIndex < audioDevices.Count; deviceIndex++)
        {
            using var audioDevice = audioDevices[deviceIndex];
            var sessions = audioDevice.AudioSessionManager.Sessions;

            for (var sessionIndex = 0; sessionIndex < sessions.Count; sessionIndex++)
            {
                using var session = sessions[sessionIndex];
                if (!ShouldIncludeSession(session))
                    continue;

                var processId = checked((int)session.GetProcessID);
                var sessionKey = BuildSessionKey(session, processId);
                if (!aggregates.TryGetValue(sessionKey, out var aggregate))
                {
                    aggregate = CreateAggregate(session, flow, processId);
                    aggregates.Add(sessionKey, aggregate);
                }

                aggregate.AddSample(audioDevice.FriendlyName, session.SimpleAudioVolume.Volume, session.SimpleAudioVolume.Mute);
            }
        }

        return aggregates.Values
            .Select(aggregate => aggregate.ToSnapshot(flow, deviceMap, defaultDeviceName))
            .OrderBy(session => session.IsSystemSession ? 0 : 1)
            .ThenBy(session => session.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public static void SetSessionVolume(string sessionKey, EDataFlow flow, float volume)
    {
        var clampedVolume = Math.Clamp(volume, 0f, 1f);
        UpdateMatchingSessions(sessionKey, flow, session => session.SimpleAudioVolume.Volume = clampedVolume);
    }

    private static void UpdateMatchingSessions(string sessionKey, EDataFlow flow, Action<AudioSessionControl> update)
    {
        using var enumerator = new MMDeviceEnumerator();
        var audioDevices = enumerator.EnumerateAudioEndPoints(ToNaudioFlow(flow), NAudioDeviceState.Active);

        for (var deviceIndex = 0; deviceIndex < audioDevices.Count; deviceIndex++)
        {
            using var audioDevice = audioDevices[deviceIndex];
            var sessions = audioDevice.AudioSessionManager.Sessions;

            for (var sessionIndex = 0; sessionIndex < sessions.Count; sessionIndex++)
            {
                using var session = sessions[sessionIndex];
                var processId = checked((int)session.GetProcessID);
                if (!string.Equals(BuildSessionKey(session, processId), sessionKey, StringComparison.OrdinalIgnoreCase))
                    continue;

                update(session);
            }
        }
    }

    private static SessionAggregate CreateAggregate(AudioSessionControl session, EDataFlow flow, int processId)
    {
        if (session.IsSystemSoundsSession || processId <= 0)
        {
            return new SessionAggregate(
                sessionKey: BuildSessionKey(session, processId),
                processId: 0,
                flow: flow,
                displayName: "System Sounds",
                processName: "System Sounds",
                executablePath: null,
                isSystemSession: true);
        }

        var processName = $"PID {processId}";
        var displayName = session.DisplayName;
        string? executablePath = null;

        try
        {
            using var process = Process.GetProcessById(processId);
            processName = process.ProcessName;
            executablePath = TryGetProcessPath(process);

            if (string.IsNullOrWhiteSpace(displayName))
                displayName = TryGetFileDescription(executablePath);

            if (string.IsNullOrWhiteSpace(displayName) && !string.IsNullOrWhiteSpace(process.MainWindowTitle))
                displayName = process.MainWindowTitle;
        }
        catch
        {
            // Ignore processes that have already exited or deny access.
        }

        if (string.IsNullOrWhiteSpace(displayName))
            displayName = processName;

        return new SessionAggregate(
            sessionKey: BuildSessionKey(session, processId),
            processId: processId,
            flow: flow,
            displayName: displayName,
            processName: processName,
            executablePath: executablePath,
            isSystemSession: false);
    }

    private static bool ShouldIncludeSession(AudioSessionControl session)
    {
        return session.IsSystemSoundsSession || session.GetProcessID > 0;
    }

    private static string BuildSessionKey(AudioSessionControl session, int processId)
    {
        if (session.IsSystemSoundsSession || processId <= 0)
            return "system";

        return $"pid:{processId}";
    }

    private static string? TryGetProcessPath(Process process)
    {
        try
        {
            return process.MainModule?.FileName;
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetFileDescription(string? executablePath)
    {
        if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            return null;

        try
        {
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(executablePath);
            return string.IsNullOrWhiteSpace(fileVersionInfo.FileDescription)
                ? Path.GetFileNameWithoutExtension(executablePath)
                : fileVersionInfo.FileDescription;
        }
        catch
        {
            return Path.GetFileNameWithoutExtension(executablePath);
        }
    }

    private static NAudioDataFlow ToNaudioFlow(EDataFlow flow)
    {
        return flow switch
        {
            EDataFlow.eRender => NAudioDataFlow.Render,
            EDataFlow.eCapture => NAudioDataFlow.Capture,
            _ => throw new NotSupportedException($"Unsupported flow: {flow}")
        };
    }

    private sealed class SessionAggregate
    {
        private readonly HashSet<string> actualDeviceNames = new(StringComparer.OrdinalIgnoreCase);
        private int sampleCount;
        private float totalVolume;
        private bool muted = true;

        public SessionAggregate(string sessionKey, int processId, EDataFlow flow, string displayName, string processName, string? executablePath, bool isSystemSession)
        {
            SessionKey = sessionKey;
            ProcessId = processId;
            Flow = flow;
            DisplayName = displayName;
            ProcessName = processName;
            ExecutablePath = executablePath;
            IsSystemSession = isSystemSession;
        }

        public string SessionKey { get; }

        public int ProcessId { get; }

        public EDataFlow Flow { get; }

        public string DisplayName { get; }

        public string ProcessName { get; }

        public string? ExecutablePath { get; }

        public bool IsSystemSession { get; }

        public void AddSample(string actualDeviceName, float volume, bool isMuted)
        {
            if (!string.IsNullOrWhiteSpace(actualDeviceName))
                actualDeviceNames.Add(actualDeviceName);

            totalVolume += volume;
            sampleCount++;
            muted &= isMuted;
        }

        public MixerSessionInfo ToSnapshot(EDataFlow flow, IReadOnlyDictionary<string, AudioDevice> deviceMap, string defaultDeviceName)
        {
            string? boundDeviceId = null;
            var boundDeviceSummary = $"Follow system default ({defaultDeviceName})";

            if (ProcessId > 0)
            {
                boundDeviceId = AudioPolicyManager.GetAppDefaultDevice((uint)ProcessId, flow);
                if (!string.IsNullOrWhiteSpace(boundDeviceId) && deviceMap.TryGetValue(boundDeviceId, out var boundDevice))
                    boundDeviceSummary = boundDevice.Name;
                else if (!string.IsNullOrWhiteSpace(boundDeviceId))
                    boundDeviceSummary = boundDeviceId;
            }

            var actualSummary = actualDeviceNames.Count switch
            {
                0 => "No active session",
                1 => actualDeviceNames.First(),
                _ => $"Multiple devices: {string.Join(", ", actualDeviceNames.Take(3))}"
            };

            return new MixerSessionInfo
            {
                SessionKey = SessionKey,
                DisplayName = DisplayName,
                ProcessName = ProcessName,
                ExecutablePath = ExecutablePath,
                Flow = Flow,
                ProcessId = ProcessId,
                Volume = sampleCount == 0 ? 0f : totalVolume / sampleCount,
                IsMuted = muted,
                IsSystemSession = IsSystemSession,
                ActualDeviceSummary = actualSummary,
                BoundDeviceId = boundDeviceId,
                BoundDeviceSummary = boundDeviceSummary
            };
        }
    }
}
