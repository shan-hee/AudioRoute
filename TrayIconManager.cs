using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;

namespace AudioRoute;

internal sealed class TrayIconManager : IDisposable
{
    private const int MaxTrayRetryAttempts = 3;
    private const uint HomeTrayMenuItemId = 1001;
    private const uint ExitTrayMenuItemId = 1002;
    private const uint StartupTrayMenuItemId = 1003;
    private const uint ViewLogTrayMenuItemId = 1004;

    private readonly Func<IntPtr> getWindowHandle;
    private readonly ShellNotifyIconHost trayIconHost;
    private readonly DispatcherQueueTimer trayRetryTimer;
    private MasterVolumeState? lastTrayVolumeState;
    private MasterVolumeState? pendingTrayRetryState;
    private DateTimeOffset lastPrimaryTrayInvokeAt;
    private DateTimeOffset suppressPrimaryTrayInvokeUntil;
    private bool disposed;
    private int trayRetryAttempt;

    public TrayIconManager(
        Func<IntPtr> getWindowHandle,
        ShellNotifyIconHost trayIconHost,
        DispatcherQueue dispatcherQueue)
    {
        this.getWindowHandle = getWindowHandle;
        this.trayIconHost = trayIconHost;
        trayRetryTimer = dispatcherQueue.CreateTimer();
        trayRetryTimer.IsRepeating = false;
        trayRetryTimer.Tick += OnTrayRetryTimerTick;
        trayIconHost.MessageReceived += OnTrayIconMessageReceived;
        trayIconHost.TaskbarCreated += OnTrayIconTaskbarCreated;
        trayIconHost.EnvironmentChanged += OnTrayIconEnvironmentChanged;
    }

    public event Func<Task>? TogglePanelRequested;
    public event Func<Task>? ShowPanelRequested;
    public event Action? ExitRequested;
    public event Action<string>? ErrorOccurred;
    public Func<bool>? IsPanelVisible { get; set; }
    public Action? CancelDeactivateHideRequested { get; set; }

    public bool IsCreated => trayIconHost.IsCreated;

    public void Initialize()
    {
        if (GetWindowHandle() == IntPtr.Zero)
            return;

        if (trayIconHost.IsCreated)
        {
            UpdateTrayIcon(force: true);
            return;
        }

        TrayVolumeIconService.RefreshEnvironment(force: true);
        var currentState = MasterVolumeService.TryGetMasterVolumeState();
        ApplyTrayIconUpdate(currentState, force: true);
    }

    public void SuppressReopenUntil(DateTimeOffset until)
    {
        suppressPrimaryTrayInvokeUntil = until;
    }

    public void UpdateTrayIcon(bool force = false)
    {
        ScheduleTrayIconUpdate(MasterVolumeService.TryGetMasterVolumeState(), force);
    }

    public void UpdateTrayIcon(MasterVolumeState? currentState, bool force = false)
    {
        ScheduleTrayIconUpdate(currentState, force);
    }

    public bool ShouldQueueTrayIconUpdate(MasterVolumeState? currentState)
    {
        return !EqualityComparer<MasterVolumeState?>.Default.Equals(currentState, lastTrayVolumeState);
    }

    public void ScheduleTrayIconUpdate(MasterVolumeState? currentState, bool force = false)
    {
        if (!trayIconHost.IsCreated && !force)
            return;

        if (!force && !ShouldQueueTrayIconUpdate(currentState))
            return;

        ApplyTrayIconUpdate(currentState, force);
    }

    public bool TryGetIconRect(out ShellNotifyIconRect rect)
    {
        return trayIconHost.TryGetIconRect(out rect);
    }

    public void HandleTrayMessage(uint trayMessage, IntPtr invokePointData, bool isPanelVisible, Action cancelDeactivateHide)
    {
        if (disposed)
            return;

        if (trayMessage == NativeMethods.NinSelect || trayMessage == NativeMethods.WmLButtonUp)
        {
            if (!isPanelVisible && DateTimeOffset.UtcNow < suppressPrimaryTrayInvokeUntil)
                return;

            if (DateTimeOffset.UtcNow - lastPrimaryTrayInvokeAt < TimeSpan.FromMilliseconds(200))
                return;

            lastPrimaryTrayInvokeAt = DateTimeOffset.UtcNow;
            cancelDeactivateHide();
            _ = TogglePanelRequested?.Invoke();

            return;
        }

        if (trayMessage == NativeMethods.WmRButtonUp || trayMessage == NativeMethods.WmContextMenu)
        {
            cancelDeactivateHide();
            ShowTrayContextMenu(GetTrayMenuAnchorPoint(invokePointData));
        }
    }

    public void HandleTaskbarCreated()
    {
        if (disposed)
            return;

        CancelTrayIconRetry();
        lastTrayVolumeState = null;
        TrayVolumeIconService.Invalidate();
        Initialize();
    }

    public void HandleEnvironmentChanged()
    {
        if (disposed)
            return;

        CancelTrayIconRetry();
        TrayVolumeIconService.Invalidate();
        UpdateTrayIcon(force: true);
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;
        CancelTrayIconRetry();
        trayRetryTimer.Tick -= OnTrayRetryTimerTick;
        trayIconHost.MessageReceived -= OnTrayIconMessageReceived;
        trayIconHost.TaskbarCreated -= OnTrayIconTaskbarCreated;
        trayIconHost.EnvironmentChanged -= OnTrayIconEnvironmentChanged;
        trayIconHost.Dispose();
        lastTrayVolumeState = null;
    }

    public static string FormatMasterVolumeStateForLog(MasterVolumeState? state)
    {
        if (state is not MasterVolumeState value)
            return "null";

        return $"{value.Percentage}%/{value.IconKind}/muted={value.IsMuted}";
    }

    private void ApplyTrayIconUpdate(MasterVolumeState? currentState, bool force)
    {
        if (!trayIconHost.IsCreated && !force)
            return;

        if (!force && EqualityComparer<MasterVolumeState?>.Default.Equals(currentState, lastTrayVolumeState))
            return;

        try
        {
            var wasCreated = trayIconHost.IsCreated;
            var iconHandle = TrayVolumeIconService.GetIconHandle(currentState?.IconKind ?? TrayVolumeIconKind.NoDevice);
            if (iconHandle == IntPtr.Zero)
                return;

            if (!trayIconHost.UpdateIcon(iconHandle, BuildTrayToolTip(currentState)))
            {
                ScheduleTrayIconRetry(currentState);
                return;
            }

            if (!trayIconHost.IsCreated)
                return;

            CancelTrayIconRetry();
            lastTrayVolumeState = currentState;
            if (!wasCreated)
                RuntimeLog.Write($"托盘创建成功: state={FormatMasterVolumeStateForLog(currentState)}");
        }
        catch (Exception ex)
        {
            RuntimeLog.WriteException(
                $"托盘图标更新异常: state={FormatMasterVolumeStateForLog(currentState)}",
                ex);
            ScheduleTrayIconRetry(currentState);
        }
    }

    private void ScheduleTrayIconRetry(MasterVolumeState? currentState)
    {
        if (disposed)
            return;

        pendingTrayRetryState = currentState;
        if (trayRetryAttempt >= MaxTrayRetryAttempts)
        {
            RuntimeLog.Write(
                $"托盘重建重试已耗尽: attempts={trayRetryAttempt}, " +
                $"state={FormatMasterVolumeStateForLog(currentState)}");
            return;
        }

        trayRetryAttempt++;
        trayRetryTimer.Stop();
        trayRetryTimer.Interval = trayRetryAttempt switch
        {
            1 => TimeSpan.FromMilliseconds(250),
            2 => TimeSpan.FromSeconds(1),
            _ => TimeSpan.FromSeconds(3)
        };
        trayRetryTimer.Start();
        RuntimeLog.Write(
            $"托盘更新失败，已安排重建: attempt={trayRetryAttempt}/{MaxTrayRetryAttempts}, " +
            $"delay={trayRetryTimer.Interval.TotalMilliseconds:F0}ms, " +
            $"state={FormatMasterVolumeStateForLog(currentState)}");
    }

    private void CancelTrayIconRetry()
    {
        trayRetryTimer.Stop();
        pendingTrayRetryState = null;
        trayRetryAttempt = 0;
    }

    private void OnTrayRetryTimerTick(DispatcherQueueTimer sender, object args)
    {
        sender.Stop();
        if (disposed)
            return;

        ApplyTrayIconUpdate(pendingTrayRetryState, force: true);
    }

    private static string BuildTrayToolTip(MasterVolumeState? volumeState)
    {
        if (volumeState is null)
            return "AudioRoute 未检测到默认输出设备";

        return volumeState.Value.IsMuted
            ? $"AudioRoute 主音量 {volumeState.Value.Percentage}% 已静音"
            : $"AudioRoute 主音量 {volumeState.Value.Percentage}%";
    }

    private NativePoint GetTrayMenuAnchorPoint(IntPtr invokePointData)
    {
        var callbackPoint = NativeMethods.GetPointFromLParam(invokePointData);
        if (callbackPoint.X != 0 || callbackPoint.Y != 0)
            return callbackPoint;

        if (trayIconHost.TryGetIconRect(out var trayIconRect))
        {
            return new NativePoint
            {
                X = trayIconRect.Left,
                Y = trayIconRect.Bottom
            };
        }

        return NativeMethods.GetCursorPos(out var cursorPosition)
            ? cursorPosition
            : default;
    }

    private void ShowTrayContextMenu(NativePoint anchorPoint)
    {
        var hwnd = GetWindowHandle();
        if (hwnd == IntPtr.Zero)
            return;

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
            _ = NativeMethods.AppendMenu(menuHandle, NativeMethods.MfString, ViewLogTrayMenuItemId, "查看日志");
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
            trayIconHost.SetFocus();
        }
        finally
        {
            _ = NativeMethods.DestroyMenu(menuHandle);
        }
    }

    private void OnTrayIconMessageReceived(object? sender, ShellNotifyIconMessageEventArgs e)
    {
        HandleTrayMessage(
            e.TrayMessage,
            e.InvokePointData,
            IsPanelVisible?.Invoke() ?? false,
            () => CancelDeactivateHideRequested?.Invoke());
    }

    private void OnTrayIconTaskbarCreated(object? sender, EventArgs e)
    {
        HandleTaskbarCreated();
    }

    private void OnTrayIconEnvironmentChanged(object? sender, EventArgs e)
    {
        HandleEnvironmentChanged();
    }

    private void HandleWmCommand(uint commandId)
    {
        if (commandId == HomeTrayMenuItemId)
        {
            _ = ShowPanelRequested?.Invoke();
            return;
        }

        if (commandId == ExitTrayMenuItemId)
        {
            ExitRequested?.Invoke();
            return;
        }

        if (commandId == StartupTrayMenuItemId)
        {
            ToggleStartupRegistration();
            return;
        }

        if (commandId == ViewLogTrayMenuItemId)
        {
            OpenRuntimeLog();
        }
    }

    public bool TryHandleWmCommand(uint commandId)
    {
        if (commandId == HomeTrayMenuItemId ||
            commandId == ExitTrayMenuItemId ||
            commandId == StartupTrayMenuItemId ||
            commandId == ViewLogTrayMenuItemId)
        {
            HandleWmCommand(commandId);
            return true;
        }

        return false;
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
            ErrorOccurred?.Invoke($"切换开机自启失败: {ex.Message}");
        }
    }

    private void OpenRuntimeLog()
    {
        if (RuntimeLog.TryOpenCurrentLog(out var errorMessage))
            return;

        ErrorOccurred?.Invoke($"打开日志失败: {errorMessage}");
    }

    private IntPtr GetWindowHandle() => getWindowHandle();
}
