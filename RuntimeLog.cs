using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace AudioRoute;

internal static class RuntimeLog
{
    private const int ArchivedLogRetentionCount = 5;
    private static readonly object SyncRoot = new();
    private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "AudioRoute.log");
    private static readonly string ArchivedLogSearchPattern = "AudioRoute-*.log";

    public static string PathOnDisk => LogPath;

    public static void Reset()
    {
        try
        {
            lock (SyncRoot)
            {
                if (File.Exists(LogPath))
                {
                    var logFile = new FileInfo(LogPath);
                    if (logFile.Length > 0)
                        File.Move(LogPath, CreateArchivePathCore());
                    else
                        File.Delete(LogPath);
                }

                DeleteExpiredArchivesCore();
            }
        }
        catch
        {
        }
    }

    public static void Write(string message)
    {
        try
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [T{Environment.CurrentManagedThreadId}] {message}";
            lock (SyncRoot)
            {
                File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch
        {
        }
    }

    public static bool TryOpenCurrentLog(out string? errorMessage)
    {
        try
        {
            lock (SyncRoot)
            {
                EnsureCurrentLogExistsCore();
            }

            _ = Process.Start(new ProcessStartInfo
            {
                FileName = LogPath,
                UseShellExecute = true
            });

            errorMessage = null;
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return false;
        }
    }

    private static void EnsureCurrentLogExistsCore()
    {
        if (File.Exists(LogPath))
            return;

        File.WriteAllText(LogPath, string.Empty, Encoding.UTF8);
    }

    private static string CreateArchivePathCore()
    {
        var directory = Path.GetDirectoryName(LogPath) ?? Path.GetTempPath();
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
        var candidatePath = Path.Combine(directory, $"AudioRoute-{timestamp}.log");
        var suffix = 1;

        while (File.Exists(candidatePath))
        {
            candidatePath = Path.Combine(directory, $"AudioRoute-{timestamp}-{suffix}.log");
            suffix++;
        }

        return candidatePath;
    }

    private static void DeleteExpiredArchivesCore()
    {
        var directory = Path.GetDirectoryName(LogPath) ?? Path.GetTempPath();
        var archivedLogs = new DirectoryInfo(directory)
            .GetFiles(ArchivedLogSearchPattern)
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .Skip(ArchivedLogRetentionCount);

        foreach (var archivedLog in archivedLogs)
        {
            try
            {
                archivedLog.Delete();
            }
            catch
            {
            }
        }
    }
}
