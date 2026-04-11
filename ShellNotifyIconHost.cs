using System;
using System.Runtime.InteropServices;

namespace AudioRoute;

internal sealed class ShellNotifyIconHost : IDisposable
{
    private const uint NimAdd = 0x00000000;
    private const uint NimModify = 0x00000001;
    private const uint NimDelete = 0x00000002;
    private const uint NimSetFocus = 0x00000003;
    private const uint NimSetVersion = 0x00000004;
    private const uint NifMessage = 0x00000001;
    private const uint NifIcon = 0x00000002;
    private const uint NifTip = 0x00000004;
    private const uint NifGuid = 0x00000020;
    private const uint NifShowTip = 0x00000080;
    private const uint NotifyIconVersion4 = 4;
    private const uint WmThemeChanged = 0x031A;
    private const uint WmSettingChange = 0x001A;
    private const uint WmDisplayChange = 0x007E;
    private const uint WmDpiChanged = 0x02E0;
    private const int GwlWndProc = -4;

    private readonly Guid guid;
    private readonly uint iconId;
    private readonly uint callbackMessage;
    private readonly uint taskbarCreatedMessage;
    private bool disposed;
    private IntPtr windowHandle;
    private IntPtr originalWindowProc;
    private WndProcDelegate? windowProcDelegate;

    public ShellNotifyIconHost(Guid guid, uint iconId, uint callbackMessage)
    {
        this.guid = guid;
        this.iconId = iconId;
        this.callbackMessage = callbackMessage;
        taskbarCreatedMessage = RegisterWindowMessage("TaskbarCreated");
    }

    public event EventHandler<ShellNotifyIconMessageEventArgs>? MessageReceived;

    public event EventHandler? TaskbarCreated;

    public event EventHandler? EnvironmentChanged;

    public bool IsCreated { get; private set; }

    public bool UpdateIcon(IntPtr iconHandle, string toolTip)
    {
        if (disposed || iconHandle == IntPtr.Zero)
            return false;

        if (!EnsureMessageWindow())
            return false;

        var data = CreateNotifyIconData(iconHandle, toolTip);
        if (!IsCreated)
            return TryAddIcon(ref data);

        if (ShellNotifyIcon(NimModify, ref data))
            return true;

        IsCreated = false;
        return TryAddIcon(ref data);
    }

    public void DeleteIcon()
    {
        if (!IsCreated || windowHandle == IntPtr.Zero)
            return;

        var data = CreateNotifyIconData(IntPtr.Zero, string.Empty);
        _ = ShellNotifyIcon(NimDelete, ref data);
        IsCreated = false;
    }

    public bool TryGetIconRect(out ShellNotifyIconRect rect)
    {
        rect = default;
        if (windowHandle == IntPtr.Zero)
            return false;

        var identifier = new NotifyIconIdentifier
        {
            cbSize = (uint)Marshal.SizeOf<NotifyIconIdentifier>(),
            hWnd = windowHandle,
            uID = iconId,
            guidItem = guid
        };

        if (ShellNotifyIconGetRect(ref identifier, out var nativeRect) != 0)
            return false;

        rect = new ShellNotifyIconRect(nativeRect.Left, nativeRect.Top, nativeRect.Right, nativeRect.Bottom);
        return true;
    }

    public void SetFocus()
    {
        if (!IsCreated || windowHandle == IntPtr.Zero)
            return;

        var data = new NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NotifyIconData>(),
            hWnd = windowHandle,
            uID = iconId,
            guidItem = guid
        };

        _ = ShellNotifyIcon(NimSetFocus, ref data);
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;
        DeleteIcon();

        if (windowHandle != IntPtr.Zero && originalWindowProc != IntPtr.Zero)
        {
            _ = SetWindowLongPtr(windowHandle, GwlWndProc, originalWindowProc);
            originalWindowProc = IntPtr.Zero;
        }

        if (windowHandle != IntPtr.Zero)
        {
            _ = DestroyWindow(windowHandle);
            windowHandle = IntPtr.Zero;
        }

        windowProcDelegate = null;
    }

    private bool EnsureMessageWindow()
    {
        if (windowHandle != IntPtr.Zero)
            return true;

        windowHandle = CreateWindowEx(
            0,
            "Static",
            "AudioRoute.TrayMessageWindow",
            0,
            0,
            0,
            0,
            0,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero);

        if (windowHandle == IntPtr.Zero)
            return false;

        windowProcDelegate = WindowProc;
        var pointer = Marshal.GetFunctionPointerForDelegate(windowProcDelegate);
        originalWindowProc = SetWindowLongPtr(windowHandle, GwlWndProc, pointer);
        return true;
    }

    private bool TryAddIcon(ref NotifyIconData data)
    {
        if (!ShellNotifyIcon(NimAdd, ref data))
            return false;

        data.uVersionOrTimeout = NotifyIconVersion4;
        _ = ShellNotifyIcon(NimSetVersion, ref data);
        IsCreated = true;
        return true;
    }

    private NotifyIconData CreateNotifyIconData(IntPtr iconHandle, string toolTip)
    {
        var flags = NifMessage | NifGuid;

        if (iconHandle != IntPtr.Zero)
            flags |= NifIcon;

        if (!string.IsNullOrWhiteSpace(toolTip))
            flags |= NifTip | NifShowTip;

        return new NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NotifyIconData>(),
            hWnd = windowHandle,
            uID = iconId,
            uFlags = flags,
            uCallbackMessage = callbackMessage,
            hIcon = iconHandle,
            szTip = toolTip,
            szInfo = string.Empty,
            szInfoTitle = string.Empty,
            uVersionOrTimeout = 0,
            guidItem = guid
        };
    }

    private IntPtr WindowProc(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam)
    {
        if (message == taskbarCreatedMessage)
        {
            IsCreated = false;
            TaskbarCreated?.Invoke(this, EventArgs.Empty);
            return IntPtr.Zero;
        }

        if (message == WmThemeChanged ||
            message == WmSettingChange ||
            message == WmDisplayChange ||
            message == WmDpiChanged)
        {
            EnvironmentChanged?.Invoke(this, EventArgs.Empty);
            return IntPtr.Zero;
        }

        if (message == callbackMessage)
        {
            MessageReceived?.Invoke(this, new ShellNotifyIconMessageEventArgs(GetLowWord(lParam), wParam));
            return IntPtr.Zero;
        }

        return originalWindowProc != IntPtr.Zero
            ? CallWindowProc(originalWindowProc, hWnd, message, wParam, lParam)
            : DefWindowProc(hWnd, message, wParam, lParam);
    }

    private static uint GetLowWord(IntPtr value)
    {
        return (uint)(value.ToInt64() & 0xFFFF);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool Shell_NotifyIcon(uint dwMessage, ref NotifyIconData lpData);

    private static bool ShellNotifyIcon(uint message, ref NotifyIconData data) => Shell_NotifyIcon(message, ref data);

    [DllImport("shell32.dll", EntryPoint = "Shell_NotifyIconGetRect", SetLastError = true)]
    private static extern int ShellNotifyIconGetRect(ref NotifyIconIdentifier identifier, out NativeRect iconLocation);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateWindowExW")]
    private static extern IntPtr CreateWindowEx(
        int dwExStyle,
        string lpClassName,
        string lpWindowName,
        int dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", EntryPoint = "CallWindowProcW", SetLastError = true)]
    private static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", EntryPoint = "DefWindowProcW", SetLastError = true)]
    private static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint RegisterWindowMessage(string lpString);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NotifyIconData
    {
        public uint cbSize;
        public IntPtr hWnd;
        public uint uID;
        public uint uFlags;
        public uint uCallbackMessage;
        public IntPtr hIcon;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szTip;

        public uint dwState;
        public uint dwStateMask;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szInfo;

        public uint uVersionOrTimeout;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;

        public uint dwInfoFlags;
        public Guid guidItem;
        public IntPtr hBalloonIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NotifyIconIdentifier
    {
        public uint cbSize;
        public IntPtr hWnd;
        public uint uID;
        public Guid guidItem;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}

internal sealed class ShellNotifyIconMessageEventArgs : EventArgs
{
    public ShellNotifyIconMessageEventArgs(uint trayMessage, IntPtr invokePointData)
    {
        TrayMessage = trayMessage;
        InvokePointData = invokePointData;
    }

    public uint TrayMessage { get; }

    public IntPtr InvokePointData { get; }
}

internal readonly record struct ShellNotifyIconRect(int Left, int Top, int Right, int Bottom);
