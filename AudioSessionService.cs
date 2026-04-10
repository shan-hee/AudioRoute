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
    private static readonly TimeSpan ProcessMetadataCacheDuration = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan FileDescriptionCacheDuration = TimeSpan.FromMinutes(30);
    private static readonly ExpiringCache<int, ProcessMetadata> ProcessMetadataCache = new(ProcessMetadataCacheDuration, 128);
    private static readonly ExpiringCache<string, string> FileDescriptionCache = new(FileDescriptionCacheDuration, 256, StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyList<MixerSessionInfo> GetActiveSessions(EDataFlow flow, IReadOnlyDictionary<string, AudioDevice>? deviceMap = null)
    {
        deviceMap ??= CreateDeviceMap(DeviceEnumerator.EnumerateDevices(flow), flow);
        var aggregates = new Dictionary<string, SessionAggregate>(StringComparer.OrdinalIgnoreCase);

        using var enumerator = new MMDeviceEnumerator();
        var audioDevices = enumerator.EnumerateAudioEndPoints(ToNaudioFlow(flow), NAudioDeviceState.Active);

        for (var deviceIndex = 0; deviceIndex < audioDevices.Count; deviceIndex++)
        {
            using var audioDevice = audioDevices[deviceIndex];
            var audioSessionManager = audioDevice.AudioSessionManager;
            try
            {
                var sessions = audioSessionManager.Sessions;
                try
                {
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

                        var simpleAudioVolume = session.SimpleAudioVolume;
                        try
                        {
                            aggregate.AddSample(audioDevice.FriendlyName, simpleAudioVolume.Volume, simpleAudioVolume.Mute);
                        }
                        finally
                        {
                            DisposeIfNeeded(simpleAudioVolume);
                        }
                    }
                }
                finally
                {
                    DisposeIfNeeded(sessions);
                }
            }
            finally
            {
                DisposeIfNeeded(audioSessionManager);
            }
        }

        return aggregates.Values
            .Select(aggregate => aggregate.ToSnapshot(flow, deviceMap))
            .OrderBy(session => session.IsSystemSession ? 0 : 1)
            .ThenBy(session => session.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public static void SetSessionVolume(string sessionKey, EDataFlow flow, float volume)
    {
        var clampedVolume = Math.Clamp(volume, 0f, 1f);
        var sessionMatcher = SessionMatcher.Create(sessionKey);
        UpdateMatchingSessions(sessionMatcher, flow, session =>
        {
            var simpleAudioVolume = session.SimpleAudioVolume;
            try
            {
                simpleAudioVolume.Volume = clampedVolume;
            }
            finally
            {
                DisposeIfNeeded(simpleAudioVolume);
            }
        });
    }

    private static void UpdateMatchingSessions(SessionMatcher sessionMatcher, EDataFlow flow, Action<AudioSessionControl> update)
    {
        using var enumerator = new MMDeviceEnumerator();
        var audioDevices = enumerator.EnumerateAudioEndPoints(ToNaudioFlow(flow), NAudioDeviceState.Active);

        for (var deviceIndex = 0; deviceIndex < audioDevices.Count; deviceIndex++)
        {
            using var audioDevice = audioDevices[deviceIndex];
            var audioSessionManager = audioDevice.AudioSessionManager;
            try
            {
                var sessions = audioSessionManager.Sessions;
                try
                {
                    for (var sessionIndex = 0; sessionIndex < sessions.Count; sessionIndex++)
                    {
                        using var session = sessions[sessionIndex];
                        if (!sessionMatcher.IsMatch(session))
                            continue;

                        update(session);
                    }
                }
                finally
                {
                    DisposeIfNeeded(sessions);
                }
            }
            finally
            {
                DisposeIfNeeded(audioSessionManager);
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
                displayName: "系统声音",
                processName: "系统声音",
                executablePath: null,
                isSystemSession: true);
        }

        var displayName = session.DisplayName;
        var processMetadata = TryGetProcessMetadata(processId);
        var processName = processMetadata?.ProcessName ?? $"PID {processId}";
        var executablePath = processMetadata?.ExecutablePath;

        if (string.IsNullOrWhiteSpace(displayName))
            displayName = processMetadata?.FileDescription;

        if (string.IsNullOrWhiteSpace(displayName))
            displayName = processMetadata?.MainWindowTitle;

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

    private static ProcessMetadata? TryGetProcessMetadata(int processId)
    {
        if (ProcessMetadataCache.TryGetValue(processId, out var cachedMetadata))
            return cachedMetadata;

        try
        {
            using var process = Process.GetProcessById(processId);
            var metadata = new ProcessMetadata(
                process.ProcessName,
                TryGetProcessPath(process),
                null,
                string.IsNullOrWhiteSpace(process.MainWindowTitle) ? null : process.MainWindowTitle);

            metadata = metadata with
            {
                FileDescription = TryGetCachedFileDescription(metadata.ExecutablePath)
            };

            ProcessMetadataCache.Set(processId, metadata);
            return metadata;
        }
        catch
        {
            // Ignore processes that have already exited or deny access.
            return null;
        }
    }

    private static string? TryGetCachedFileDescription(string? executablePath)
    {
        if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            return null;

        if (FileDescriptionCache.TryGetValue(executablePath, out var cachedDescription))
            return cachedDescription;

        var description = TryGetFileDescription(executablePath);
        if (string.IsNullOrWhiteSpace(description))
            return null;

        FileDescriptionCache.Set(executablePath, description);
        return description;
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

    private static void DisposeIfNeeded(object? instance)
    {
        if (instance is IDisposable disposable)
            disposable.Dispose();
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

        public MixerSessionInfo ToSnapshot(EDataFlow flow, IReadOnlyDictionary<string, AudioDevice> deviceMap)
        {
            string? boundDeviceId = null;
            var boundDeviceSummary = "跟随系统默认";

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
                0 => "无活跃会话",
                1 => actualDeviceNames.First(),
                _ => $"多个设备: {string.Join(", ", actualDeviceNames.Take(3))}"
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

    private sealed record ProcessMetadata(
        string ProcessName,
        string? ExecutablePath,
        string? FileDescription,
        string? MainWindowTitle);

    private readonly record struct SessionMatcher(bool IsSystemSession, int ProcessId)
    {
        public static SessionMatcher Create(string sessionKey)
        {
            if (string.Equals(sessionKey, "system", StringComparison.OrdinalIgnoreCase))
                return new SessionMatcher(true, 0);

            if (sessionKey.StartsWith("pid:", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(sessionKey.AsSpan(4), out var processId))
                    return new SessionMatcher(false, processId);
            }

            return new SessionMatcher(false, int.MinValue);
        }

        public bool IsMatch(AudioSessionControl session)
        {
            if (IsSystemSession)
                return session.IsSystemSoundsSession;

            if (ProcessId <= 0)
                return false;

            return !session.IsSystemSoundsSession &&
                checked((int)session.GetProcessID) == ProcessId;
        }
    }
}
