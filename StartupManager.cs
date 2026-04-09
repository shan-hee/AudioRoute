using System;
using System.Diagnostics;
using Microsoft.Win32;

namespace AudioRoute;

public static class StartupManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupValueName = "AudioRoute";
    private const string StartupArgument = "--startup";

    public static bool IsEnabled()
    {
        using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        var value = runKey?.GetValue(StartupValueName) as string;
        return !string.IsNullOrWhiteSpace(value);
    }

    public static void SetEnabled(bool enabled)
    {
        using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
            ?? Registry.CurrentUser.CreateSubKey(RunKeyPath);

        if (runKey is null)
            throw new InvalidOperationException("无法打开开机启动注册表项。");

        if (enabled)
        {
            runKey.SetValue(StartupValueName, BuildStartupCommand(), RegistryValueKind.String);
        }
        else
        {
            runKey.DeleteValue(StartupValueName, throwOnMissingValue: false);
        }
    }

    public static bool IsStartupLaunch(string? arguments)
    {
        return !string.IsNullOrWhiteSpace(arguments) &&
               arguments.Contains(StartupArgument, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildStartupCommand()
    {
        var executablePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(executablePath))
            executablePath = Process.GetCurrentProcess().MainModule?.FileName;

        if (string.IsNullOrWhiteSpace(executablePath))
            throw new InvalidOperationException("无法确定当前程序路径。");

        return $"\"{executablePath}\" {StartupArgument}";
    }
}
