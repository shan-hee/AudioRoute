using System;
using System.Diagnostics;
using System.Numerics;
using System.Threading.Tasks;
using Microsoft.UI.Composition;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Hosting;
using Windows.Graphics;

namespace AudioRoute;

internal sealed class PanelController : IDisposable
{
    private const int PanelWidth = 400;
    private const int PanelHeight = 460;
    private const int ScreenMargin = 18;
    private static readonly TimeSpan DeactivateHideDelay = TimeSpan.FromMilliseconds(140);
    private static readonly TimeSpan ForegroundMonitorFallbackInterval = TimeSpan.FromMilliseconds(650);
    private static readonly TimeSpan HideSuppressionDuration = TimeSpan.FromMilliseconds(320);
    private static readonly TimeSpan PanelOpenOffsetAnimationDuration = TimeSpan.FromMilliseconds(240);
    private static readonly TimeSpan PanelOpenOpacityAnimationDuration = TimeSpan.FromMilliseconds(220);
    private static readonly TimeSpan PanelCloseOffsetAnimationDuration = TimeSpan.FromMilliseconds(120);
    private static readonly TimeSpan PanelCloseOpacityAnimationDuration = TimeSpan.FromMilliseconds(100);

    private readonly IntPtr hwnd;
    private readonly AppWindow appWindow;
    private readonly DispatcherQueue dispatcherQueue;
    private readonly UIElement panelSurface;
    private readonly TrayIconManager trayIconManager;
    private readonly DispatcherQueueTimer deactivateHideTimer;
    private readonly DispatcherQueueTimer foregroundMonitorTimer;

    private bool isPanelVisible;
    private bool isVisibilityTransitioning;
    private bool allowDeactivateHide;
    private DateTimeOffset suppressHideUntil;
    private WinEventDelegate? foregroundEventDelegate;
    private IntPtr foregroundEventHook;
    private bool disposed;

    public PanelController(
        IntPtr hwnd,
        AppWindow appWindow,
        DispatcherQueue dispatcherQueue,
        UIElement panelSurface,
        TrayIconManager trayIconManager)
    {
        this.hwnd = hwnd;
        this.appWindow = appWindow;
        this.dispatcherQueue = dispatcherQueue;
        this.panelSurface = panelSurface;
        this.trayIconManager = trayIconManager;

        deactivateHideTimer = dispatcherQueue.CreateTimer();
        deactivateHideTimer.Interval = DeactivateHideDelay;
        deactivateHideTimer.IsRepeating = false;
        deactivateHideTimer.Tick += async (_, _) =>
        {
            deactivateHideTimer.Stop();

            if (ShouldHideForLostForeground())
                await HideToTrayAsync();
        };

        foregroundMonitorTimer = dispatcherQueue.CreateTimer();
        foregroundMonitorTimer.Interval = ForegroundMonitorFallbackInterval;
        foregroundMonitorTimer.IsRepeating = true;
        foregroundMonitorTimer.Tick += async (_, _) =>
        {
            if (ShouldHideForLostForeground())
                await HideToTrayAsync();
        };
    }

    public bool IsPanelVisible => isPanelVisible;
    public bool IsVisibilityTransitioning => isVisibilityTransitioning;

    /// <summary>
    /// Called by MainWindow to query whether deactivate-hide should be suppressed
    /// due to active user interaction (e.g. volume slider drag).
    /// </summary>
    public Func<bool>? IsInteracting { get; set; }

    /// <summary>
    /// Called by MainWindow to query whether exit has been requested.
    /// </summary>
    public Func<bool>? IsExitRequested { get; set; }

    /// <summary>
    /// Called before the panel is shown, to allow MainWindow to prepare cached snapshot / loading UI.
    /// </summary>
    public Func<Task>? OnBeforeShow { get; set; }

    /// <summary>
    /// Raised after the panel becomes visible.
    /// </summary>
    public event Action? PanelShown;

    /// <summary>
    /// Raised after the panel is hidden.
    /// </summary>
    public event Action? PanelHidden;

    /// <summary>
    /// Raised when the panel loses foreground and should schedule a deactivate-hide check.
    /// </summary>
    #pragma warning disable CS0067
    public event Action? PanelDeactivated;
    #pragma warning restore CS0067

    private bool CheckExitRequested() => IsExitRequested?.Invoke() ?? false;
    private bool CheckInteracting() => IsInteracting?.Invoke() ?? false;

    public void AttachForegroundEventHook()
    {
        if (foregroundEventHook != IntPtr.Zero)
            return;

        foregroundEventDelegate = ForegroundEventProc;
        foregroundEventHook = NativeMethods.SetWinEventHook(
            NativeMethods.EventSystemForeground,
            NativeMethods.EventSystemForeground,
            IntPtr.Zero,
            foregroundEventDelegate,
            0,
            0,
            NativeMethods.WineventOutofcontext | NativeMethods.WineventSkipOwnProcess);

        if (foregroundEventHook == IntPtr.Zero)
            Trace.WriteLine("[AudioRoute] Failed to attach foreground event hook, fallback timer will be used.");
    }

    public void DetachForegroundEventHook()
    {
        if (foregroundEventHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(foregroundEventHook);
            foregroundEventHook = IntPtr.Zero;
        }

        foregroundEventDelegate = null;
    }

    public void HandlePanelDeactivated()
    {
        if (!allowDeactivateHide || CheckExitRequested() || !isPanelVisible || isVisibilityTransitioning || CheckInteracting())
            return;

        ScheduleDeactivateHide();
    }

    public async Task ShowPanelAsync(Action activate)
    {
        if (CheckExitRequested() || isVisibilityTransitioning || isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;

        try
        {
            if (OnBeforeShow is not null)
                await OnBeforeShow();

            if (CheckExitRequested() || isPanelVisible)
                return;

            UpdateWindowPosition();

            var visual = ElementCompositionPreview.GetElementVisual(panelSurface);
            visual.Opacity = 0f;
            visual.Offset = new Vector3(0f, 26f, 0f);

            NativeMethods.ShowWindow(hwnd, NativeMethods.SwShow);
            activate();
            NativeMethods.SetForegroundWindow(hwnd);
            NativeMethods.SetTopMost(hwnd);
            NativeMethods.SetWindowCloaked(hwnd, false);

            await AnimatePanelAsync(visual, new Vector3(0f, 26f, 0f), Vector3.Zero, 0f, 1f);
            NativeMethods.SetForegroundWindow(hwnd);
            isPanelVisible = true;
            suppressHideUntil = DateTimeOffset.UtcNow + HideSuppressionDuration;
            allowDeactivateHide = true;
            StartForegroundMonitorFallbackIfNeeded();
            PanelShown?.Invoke();
        }
        finally
        {
            isVisibilityTransitioning = false;
            allowDeactivateHide = isPanelVisible && !CheckExitRequested();
        }
    }

    public async Task HideToTrayAsync()
    {
        if (CheckExitRequested() || isVisibilityTransitioning || !isPanelVisible)
            return;

        CancelDeactivateHide();
        isVisibilityTransitioning = true;
        allowDeactivateHide = false;
        foregroundMonitorTimer.Stop();

        try
        {
            var visual = ElementCompositionPreview.GetElementVisual(panelSurface);
            await AnimatePanelAsync(
                visual,
                Vector3.Zero,
                new Vector3(0f, 24f, 0f),
                1f,
                0f,
                PanelCloseOffsetAnimationDuration,
                PanelCloseOpacityAnimationDuration);
            NativeMethods.SetWindowCloaked(hwnd, true);
            NativeMethods.ShowWindow(hwnd, NativeMethods.SwHide);
            isPanelVisible = false;
            PanelHidden?.Invoke();
        }
        finally
        {
            isVisibilityTransitioning = false;
        }
    }

    public void BringPanelToFront(Action activate)
    {
        if (CheckExitRequested())
            return;

        UpdateWindowPosition();
        NativeMethods.ShowWindow(hwnd, NativeMethods.SwShow);
        activate();
        NativeMethods.SetForegroundWindow(hwnd);
        NativeMethods.SetTopMost(hwnd);
        NativeMethods.SetWindowCloaked(hwnd, false);
        suppressHideUntil = DateTimeOffset.UtcNow + HideSuppressionDuration;
        allowDeactivateHide = true;
        StartForegroundMonitorFallbackIfNeeded();
    }

    public void CancelDeactivateHide()
    {
        if (deactivateHideTimer.IsRunning)
            deactivateHideTimer.Stop();
    }

    public void StopTimers()
    {
        allowDeactivateHide = false;
        deactivateHideTimer.Stop();
        foregroundMonitorTimer.Stop();
    }

    public void UpdateWindowPosition()
    {
        if (TryGetTrayAnchoredPanelPosition(out var panelPosition))
        {
            appWindow.Move(panelPosition);
            return;
        }

        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
        var displayArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary);
        var workArea = displayArea.WorkArea;
        var x = workArea.X + workArea.Width - PanelWidth - ScreenMargin;
        var y = workArea.Y + workArea.Height - PanelHeight - ScreenMargin;
        appWindow.Move(new PointInt32(x, y));
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;
        StopTimers();
        DetachForegroundEventHook();
    }

    private void ScheduleDeactivateHide()
    {
        deactivateHideTimer.Stop();
        deactivateHideTimer.Start();
    }

    private bool ShouldHideForLostForeground()
    {
        if (!allowDeactivateHide || CheckExitRequested() || !isPanelVisible || isVisibilityTransitioning || CheckInteracting())
            return false;

        if (DateTimeOffset.UtcNow < suppressHideUntil)
            return false;

        var foregroundWindow = NativeMethods.GetForegroundWindow();
        if (foregroundWindow == IntPtr.Zero)
            return true;

        return NativeMethods.GetWindowProcessId(foregroundWindow) != Environment.ProcessId;
    }

    private void StartForegroundMonitorFallbackIfNeeded()
    {
        if (foregroundEventHook != IntPtr.Zero)
            return;

        if (!foregroundMonitorTimer.IsRunning)
            foregroundMonitorTimer.Start();
    }

    private void ForegroundEventProc(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr eventWindow,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime)
    {
        if (eventType != NativeMethods.EventSystemForeground ||
            idObject != NativeMethods.ObjIdWindow ||
            CheckExitRequested() ||
            !isPanelVisible ||
            !allowDeactivateHide)
        {
            return;
        }

        dispatcherQueue.TryEnqueue(() => HandlePanelDeactivated());
    }

    private bool TryGetTrayAnchoredPanelPosition(out PointInt32 panelPosition)
    {
        panelPosition = default;

        if (!trayIconManager.TryGetIconRect(out var trayIconRect))
            return false;

        var trayCenter = new PointInt32(
            trayIconRect.Left + ((trayIconRect.Right - trayIconRect.Left) / 2),
            trayIconRect.Top + ((trayIconRect.Bottom - trayIconRect.Top) / 2));
        var displayArea = DisplayArea.GetFromPoint(trayCenter, DisplayAreaFallback.Nearest);
        if (displayArea is null)
            return false;

        panelPosition = CalculatePanelPosition(displayArea.WorkArea, trayIconRect);
        return true;
    }

    private static PointInt32 CalculatePanelPosition(RectInt32 workArea, ShellNotifyIconRect trayIconRect)
    {
        var minX = workArea.X + ScreenMargin;
        var minY = workArea.Y + ScreenMargin;
        var maxX = Math.Max(minX, workArea.X + workArea.Width - PanelWidth - ScreenMargin);
        var maxY = Math.Max(minY, workArea.Y + workArea.Height - PanelHeight - ScreenMargin);

        return DetectTrayDockEdge(workArea, trayIconRect) switch
        {
            TrayDockEdge.Left => new PointInt32(minX, maxY),
            TrayDockEdge.Top => new PointInt32(maxX, minY),
            TrayDockEdge.Right => new PointInt32(maxX, maxY),
            _ => new PointInt32(maxX, maxY)
        };
    }

    private static TrayDockEdge DetectTrayDockEdge(RectInt32 workArea, ShellNotifyIconRect trayIconRect)
    {
        var workAreaLeft = workArea.X;
        var workAreaTop = workArea.Y;
        var workAreaRight = workArea.X + workArea.Width;
        var workAreaBottom = workArea.Y + workArea.Height;

        var leftOverflow = Math.Max(0, workAreaLeft - trayIconRect.Left);
        var topOverflow = Math.Max(0, workAreaTop - trayIconRect.Top);
        var rightOverflow = Math.Max(0, trayIconRect.Right - workAreaRight);
        var bottomOverflow = Math.Max(0, trayIconRect.Bottom - workAreaBottom);

        if (bottomOverflow > 0 || rightOverflow > 0 || leftOverflow > 0 || topOverflow > 0)
        {
            if (bottomOverflow >= rightOverflow && bottomOverflow >= leftOverflow && bottomOverflow >= topOverflow)
                return TrayDockEdge.Bottom;

            if (rightOverflow >= leftOverflow && rightOverflow >= topOverflow)
                return TrayDockEdge.Right;

            if (leftOverflow >= topOverflow)
                return TrayDockEdge.Left;

            return TrayDockEdge.Top;
        }

        var leftDistance = Math.Abs(trayIconRect.Left - workAreaLeft);
        var topDistance = Math.Abs(trayIconRect.Top - workAreaTop);
        var rightDistance = Math.Abs(workAreaRight - trayIconRect.Right);
        var bottomDistance = Math.Abs(workAreaBottom - trayIconRect.Bottom);
        var nearestDistance = Math.Min(Math.Min(leftDistance, rightDistance), Math.Min(topDistance, bottomDistance));

        if (nearestDistance == bottomDistance)
            return TrayDockEdge.Bottom;

        if (nearestDistance == rightDistance)
            return TrayDockEdge.Right;

        if (nearestDistance == leftDistance)
            return TrayDockEdge.Left;

        return TrayDockEdge.Top;
    }

    private static Task AnimatePanelAsync(
        Visual visual,
        Vector3 fromOffset,
        Vector3 toOffset,
        float fromOpacity,
        float toOpacity,
        TimeSpan? offsetDuration = null,
        TimeSpan? opacityDuration = null)
    {
        var compositor = visual.Compositor;
        var batch = compositor.CreateScopedBatch(CompositionBatchTypes.Animation);
        var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        var easing = compositor.CreateCubicBezierEasingFunction(
            new Vector2(0.18f, 0.9f),
            new Vector2(0.24f, 1f));

        var offsetAnimation = compositor.CreateVector3KeyFrameAnimation();
        offsetAnimation.InsertKeyFrame(0f, fromOffset);
        offsetAnimation.InsertKeyFrame(1f, toOffset, easing);
        offsetAnimation.Duration = offsetDuration ?? PanelOpenOffsetAnimationDuration;

        var opacityAnimation = compositor.CreateScalarKeyFrameAnimation();
        opacityAnimation.InsertKeyFrame(0f, fromOpacity);
        opacityAnimation.InsertKeyFrame(1f, toOpacity, easing);
        opacityAnimation.Duration = opacityDuration ?? PanelOpenOpacityAnimationDuration;

        visual.StartAnimation(nameof(visual.Offset), offsetAnimation);
        visual.StartAnimation(nameof(visual.Opacity), opacityAnimation);

        batch.Completed += (_, _) => completed.TrySetResult();
        batch.End();

        return completed.Task;
    }
}
