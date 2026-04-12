using System;
using System.Runtime.InteropServices;

namespace AudioRoute;

[StructLayout(LayoutKind.Sequential)]
internal struct NativePoint
{
    public int X;
    public int Y;
}

internal delegate void WinEventDelegate(
    IntPtr hWinEventHook,
    uint eventType,
    IntPtr hwnd,
    int idObject,
    int idChild,
    uint eventThread,
    uint eventTime);

internal delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

internal static class NativeMethods
{
    private const int GwlExStyle = -20;
    private const int GwlStyle = -16;
    private const int GwlWndProc = -4;
    private const int WsBorder = 0x00800000;
    private const int WsCaption = 0x00C00000;
    private const int WsDlgFrame = 0x00400000;
    private const int WsExAppWindow = 0x00040000;
    private const int WsExToolWindow = 0x00000080;
    private const int WsThickFrame = 0x00040000;
    private const uint DwmaCloak = 13;
    private const uint DwmaWindowCornerPreference = 33;
    private const int DwmWindowCornerPreferenceRound = 2;
    public const int SwHide = 0;
    public const int SwShow = 5;
    public const uint WaInactive = 0;
    public const uint WmApp = 0x8000;
    public const uint WmActivate = 0x0006;
    public const uint WmActivateApp = 0x001C;
    public const uint WmCommand = 0x0111;
    public const uint WmContextMenu = 0x007B;
    public const uint WmLButtonUp = 0x0202;
    public const uint WmNull = 0x0000;
    public const uint WmRButtonUp = 0x0205;
    public const uint EventSystemForeground = 0x0003;
    public const uint NinSelect = 0x0400;
    public const uint MfString = 0x00000000;
    public const uint MfChecked = 0x00000008;
    public const uint MfSeparator = 0x00000800;
    public const uint WineventOutofcontext = 0x0000;
    public const uint WineventSkipOwnProcess = 0x0002;
    public const int ObjIdWindow = 0;
    private const uint SwpFrameChanged = 0x0020;
    private const uint SwpNomove = 0x0002;
    private const uint SwpNozorder = 0x0004;
    private const uint SwpNosize = 0x0001;
    private const uint SwpNoactivate = 0x0010;
    public const uint TpmBottomAlign = 0x0020;
    public const uint TpmLeftAlign = 0x0000;
    public const uint TpmRightButton = 0x0002;
    private static readonly IntPtr HwndTopMost = new(-1);

    public static void ApplyToolWindowStyle(IntPtr hwnd)
    {
        var style = GetWindowLongPtr(hwnd, GwlStyle).ToInt64();
        style &= ~(WsCaption | WsThickFrame | WsBorder | WsDlgFrame);
        _ = SetWindowLongPtr(hwnd, GwlStyle, new IntPtr(style));

        var exStyle = GetWindowLongPtr(hwnd, GwlExStyle).ToInt64();
        exStyle |= WsExToolWindow;
        exStyle &= ~WsExAppWindow;
        _ = SetWindowLongPtr(hwnd, GwlExStyle, new IntPtr(exStyle));

        _ = SetWindowPos(
            hwnd,
            IntPtr.Zero,
            0,
            0,
            0,
            0,
            SwpNomove | SwpNosize | SwpNozorder | SwpNoactivate | SwpFrameChanged);
    }

    public static void SetTopMost(IntPtr hwnd)
    {
        _ = SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNoactivate);
    }

    public static void ShowWindow(IntPtr hwnd, int command)
    {
        _ = ShowWindowNative(hwnd, command);
    }

    public static void SetWindowCloaked(IntPtr hwnd, bool cloaked)
    {
        var value = cloaked ? 1 : 0;
        _ = DwmSetWindowAttribute(hwnd, DwmaCloak, ref value, Marshal.SizeOf<int>());
    }

    public static void ApplyWindowCornerPreference(IntPtr hwnd)
    {
        var value = DwmWindowCornerPreferenceRound;
        _ = DwmSetWindowAttribute(hwnd, DwmaWindowCornerPreference, ref value, Marshal.SizeOf<int>());
    }

    public static IntPtr SetWindowProc(IntPtr hwnd, IntPtr windowProc)
    {
        return SetWindowLongPtr(hwnd, GwlWndProc, windowProc);
    }

    public static uint GetLowWord(IntPtr value)
    {
        return (uint)(value.ToInt64() & 0xFFFF);
    }

    public static NativePoint GetPointFromLParam(IntPtr value)
    {
        var raw = value.ToInt64();
        return new NativePoint
        {
            X = unchecked((short)(raw & 0xFFFF)),
            Y = unchecked((short)((raw >> 16) & 0xFFFF))
        };
    }

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
    private static extern bool ShowWindowNative(IntPtr hWnd, int nCmdShow);

    [DllImport("dwmapi.dll", SetLastError = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, uint dwAttribute, ref int pvAttribute, int cbAttribute);

    [DllImport("user32.dll", EntryPoint = "CallWindowProcW", SetLastError = true)]
    public static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", EntryPoint = "DefWindowProcW", SetLastError = true)]
    public static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr CreatePopupMenu();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool AppendMenu(IntPtr hMenu, uint uFlags, uint uIDNewItem, string lpNewItem);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyMenu(IntPtr hMenu);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetCursorPos(out NativePoint cursorPosition);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, IntPtr prcRect);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SetWinEventHook(
        uint eventMin,
        uint eventMax,
        IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc,
        uint idProcess,
        uint idThread,
        uint dwFlags);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    public static uint GetWindowProcessId(IntPtr hWnd)
    {
        _ = GetWindowThreadProcessId(hWnd, out var processId);
        return processId;
    }
}
