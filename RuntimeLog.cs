using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace AudioRoute;

internal static class RuntimeLog
{
    private const int ArchivedLogRetentionCount = 10;
    private const long MaxCurrentLogLength = 5 * 1024 * 1024;
    private const string ArchivedLogSearchPattern = "AudioRoute-*.log";
    private static readonly object SyncRoot = new();
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AudioRoute",
        "Logs");
    private static readonly string LogPath = Path.Combine(LogDirectory, "AudioRoute.log");

    public static string PathOnDisk => LogPath;

    public static void BeginSession()
    {
        try
        {
            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
                RotateOversizedLogCore();
                DeleteExpiredArchivesCore();
                DeleteLegacyTempLogsCore();
            }
        }
        catch
        {
        }

        Write("========== 新运行会话 ==========");
    }

    public static void Write(string message)
    {
        try
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [T{Environment.CurrentManagedThreadId}] {message}";
            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch
        {
        }
    }

    public static void WriteException(string operation, Exception exception)
    {
        Write($"{operation}: {exception}");
    }

    public static bool TryOpenCurrentLog(out string? errorMessage)
    {
        try
        {
            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
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

    private static void RotateOversizedLogCore()
    {
        if (!File.Exists(LogPath))
            return;

        var logFile = new FileInfo(LogPath);
        if (logFile.Length < MaxCurrentLogLength)
            return;

        File.Move(LogPath, CreateArchivePathCore());
    }

    private static string CreateArchivePathCore()
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
        var candidatePath = Path.Combine(LogDirectory, $"AudioRoute-{timestamp}.log");
        var suffix = 1;

        while (File.Exists(candidatePath))
        {
            candidatePath = Path.Combine(LogDirectory, $"AudioRoute-{timestamp}-{suffix}.log");
            suffix++;
        }

        return candidatePath;
    }

    private static void DeleteExpiredArchivesCore()
    {
        var archivedLogs = new DirectoryInfo(LogDirectory)
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

    private static void DeleteLegacyTempLogsCore()
    {
        var tempDirectory = new DirectoryInfo(Path.GetTempPath());
        foreach (var legacyLog in tempDirectory.GetFiles("AudioRoute*.log"))
        {
            try
            {
                legacyLog.Delete();
            }
            catch
            {
            }
        }
    }
}
