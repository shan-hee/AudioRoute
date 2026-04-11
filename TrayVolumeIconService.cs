using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace AudioRoute;

public static class TrayVolumeIconService
{
    private const int MinimumIconSize = 16;
    private const int SmCxSmIcon = 49;
    private const int SmCySmIcon = 50;
    private const float TargetGlyphScale = 0.9f;
    private const string SndVolSsoPath = @"%SystemRoot%\System32\SndVolSSO.dll";
    private const string SystemPersonalizeKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string SystemUsesLightThemeValueName = "SystemUsesLightTheme";
    private const int ColorWindowText = 8;
    private const uint SpiGetHighContrast = 0x0042;
    private const int HcfHighContrastOn = 0x00000001;
    private const uint RtIcon = 3;
    private const uint RtGroupIcon = 14;
    private const uint DefaultIconCursorVersion = 0x00030000;
    private const uint LoadLibraryAsDataFile = 0x00000002;
    private const uint LoadLibraryAsImageResource = 0x00000020;

    private static readonly string SystemVolumeDllPath = Environment.ExpandEnvironmentVariables(SndVolSsoPath);
    private static readonly object SyncRoot = new();
    private static readonly Dictionary<IconCacheKey, IntPtr> CachedIcons = new();
    private static TrayIconEnvironment? currentEnvironment;

    public static IntPtr GetIconHandle(TrayVolumeIconKind iconKind)
    {
        lock (SyncRoot)
        {
            var environment = EnsureEnvironment();
            var cacheKey = new IconCacheKey(iconKind, environment.Dpi, environment.Theme);
            if (CachedIcons.TryGetValue(cacheKey, out var iconHandle) && iconHandle != IntPtr.Zero)
                return iconHandle;

            iconHandle = CreateIconHandle(cacheKey);
            CachedIcons[cacheKey] = iconHandle;
            return iconHandle;
        }
    }

    public static void Dispose()
    {
        lock (SyncRoot)
        {
            DisposeCachedIcons();
            currentEnvironment = null;
        }
    }

    public static void Invalidate()
    {
        lock (SyncRoot)
        {
            DisposeCachedIcons();
            currentEnvironment = null;
        }
    }

    public static void RefreshEnvironment(bool force = false)
    {
        lock (SyncRoot)
        {
            var nextEnvironment = new TrayIconEnvironment(ResolveTaskbarDpi(), ResolveTaskbarTheme());
            if (!force && currentEnvironment is TrayIconEnvironment existingEnvironment && existingEnvironment.Equals(nextEnvironment))
                return;

            DisposeCachedIcons();
            currentEnvironment = nextEnvironment;
        }
    }

    private static void DisposeCachedIcons()
    {
        foreach (var iconHandle in CachedIcons.Values)
        {
            if (iconHandle != IntPtr.Zero)
                NativeTrayMethods.DestroyIcon(iconHandle);
        }

        CachedIcons.Clear();
    }

    private static TrayIconEnvironment EnsureEnvironment()
    {
        if (currentEnvironment is TrayIconEnvironment environment)
            return environment;

        environment = new TrayIconEnvironment(ResolveTaskbarDpi(), ResolveTaskbarTheme());
        currentEnvironment = environment;
        return environment;
    }

    private static IntPtr CreateIconHandle(IconCacheKey cacheKey)
    {
        try
        {
            var systemIcon = LoadSystemVolumeIconHandle(cacheKey.IconKind, cacheKey.Dpi);
            if (systemIcon != IntPtr.Zero)
                return ApplyTheme(systemIcon, cacheKey.IconKind, cacheKey.Theme);
        }
        catch
        {
            // Fall back to the previous glyph-based icon path if the system resources
            // are unavailable so the tray icon still renders on unusual environments.
        }

        return CreateFallbackGlyphIconHandle(cacheKey.IconKind, cacheKey.Dpi, cacheKey.Theme);
    }

    private static IntPtr LoadSystemVolumeIconHandle(TrayVolumeIconKind iconKind, uint dpi)
    {
        var resourceId = GetSystemIconResourceId(iconKind);
        var iconSize = ResolveIconSize(dpi);
        var moduleHandle = NativeTrayMethods.LoadLibraryEx(
            SystemVolumeDllPath,
            IntPtr.Zero,
            LoadLibraryAsDataFile | LoadLibraryAsImageResource);

        if (moduleHandle == IntPtr.Zero)
            return IntPtr.Zero;

        try
        {
            var groupResource = NativeTrayMethods.FindResource(moduleHandle, new IntPtr(resourceId), new IntPtr(RtGroupIcon));
            if (groupResource == IntPtr.Zero)
                return IntPtr.Zero;

            var groupResourceHandle = NativeTrayMethods.LoadResource(moduleHandle, groupResource);
            if (groupResourceHandle == IntPtr.Zero)
                return IntPtr.Zero;

            var groupResourceData = NativeTrayMethods.LockResource(groupResourceHandle);
            if (groupResourceData == IntPtr.Zero)
                return IntPtr.Zero;

            var iconId = NativeTrayMethods.LookupIconIdFromDirectoryEx(groupResourceData, true, iconSize, iconSize, 0);
            if (iconId == 0)
                return IntPtr.Zero;

            var iconResource = NativeTrayMethods.FindResource(moduleHandle, new IntPtr(iconId), new IntPtr(RtIcon));
            if (iconResource == IntPtr.Zero)
                return IntPtr.Zero;

            var iconResourceHandle = NativeTrayMethods.LoadResource(moduleHandle, iconResource);
            if (iconResourceHandle == IntPtr.Zero)
                return IntPtr.Zero;

            var iconResourceData = NativeTrayMethods.LockResource(iconResourceHandle);
            if (iconResourceData == IntPtr.Zero)
                return IntPtr.Zero;

            var iconResourceSize = NativeTrayMethods.SizeofResource(moduleHandle, iconResource);
            if (iconResourceSize == 0)
                return IntPtr.Zero;

            return NativeTrayMethods.CreateIconFromResourceEx(
                iconResourceData,
                iconResourceSize,
                true,
                DefaultIconCursorVersion,
                iconSize,
                iconSize,
                0);
        }
        finally
        {
            _ = NativeTrayMethods.FreeLibrary(moduleHandle);
        }
    }

    private static IntPtr ApplyTheme(IntPtr iconHandle, TrayVolumeIconKind iconKind, TrayIconTheme theme)
    {
        return theme switch
        {
            TrayIconTheme.Light => RecolorIcon(iconHandle, iconKind, Color.Black),
            TrayIconTheme.HighContrast => RecolorIcon(iconHandle, iconKind, ResolveHighContrastTextColor()),
            _ => iconHandle
        };
    }

    private static IntPtr RecolorIcon(IntPtr iconHandle, TrayVolumeIconKind iconKind, Color targetColor)
    {
        using var icon = Icon.FromHandle(iconHandle);
        using var bitmap = icon.ToBitmap();

        var recolorLimit = Math.Clamp((int)Math.Ceiling(bitmap.Width * GetFillPercent(iconKind)), 0, bitmap.Width);
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < recolorLimit; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 220)
                    bitmap.SetPixel(x, y, Color.FromArgb(pixel.A, targetColor.R, targetColor.G, targetColor.B));
            }
        }

        var recoloredIconHandle = bitmap.GetHicon();
        NativeTrayMethods.DestroyIcon(iconHandle);
        return recoloredIconHandle;
    }

    private static float GetFillPercent(TrayVolumeIconKind iconKind) => iconKind == TrayVolumeIconKind.NoDevice ? 0.4f : 1f;

    private static IntPtr CreateFallbackGlyphIconHandle(TrayVolumeIconKind iconKind, uint dpi, TrayIconTheme theme)
    {
        var iconSize = ResolveIconSize(dpi);
        using var bitmap = new Bitmap(iconSize, iconSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

        using var brush = new SolidBrush(ResolveFallbackGlyphColor(theme));
        graphics.Clear(Color.Transparent);
        DrawGlyph(graphics, brush, GetFallbackGlyph(iconKind), iconSize);
        return bitmap.GetHicon();
    }

    private static void DrawGlyph(Graphics graphics, Brush brush, string glyph, int iconSize)
    {
        var fontFamily = ResolveFontFamily();
        var probeFontSize = Math.Max(iconSize, 24);
        using var probeFont = new Font(fontFamily, probeFontSize, FontStyle.Regular, GraphicsUnit.Pixel);
        var probeBounds = MeasureGlyphBounds(graphics, glyph, probeFont);
        if (probeBounds.Width <= 0 || probeBounds.Height <= 0)
            return;

        var targetGlyphSize = Math.Max(iconSize * TargetGlyphScale, iconSize - 1f);
        var scale = Math.Min(targetGlyphSize / probeBounds.Width, targetGlyphSize / probeBounds.Height);
        var fontSize = Math.Max(1f, (float)Math.Round(probeFontSize * scale));

        using var font = new Font(fontFamily, fontSize, FontStyle.Regular, GraphicsUnit.Pixel);
        using var format = (StringFormat)StringFormat.GenericTypographic.Clone();
        format.FormatFlags |= StringFormatFlags.NoClip;
        var glyphBounds = MeasureGlyphBounds(graphics, glyph, font);
        if (glyphBounds.Width <= 0 || glyphBounds.Height <= 0)
            return;

        var verticalOffset = Math.Max(0.25f, iconSize * 0.03f);
        var drawPoint = new PointF(
            (float)Math.Round(((iconSize - glyphBounds.Width) / 2f) - glyphBounds.X),
            (float)Math.Round(((iconSize - glyphBounds.Height) / 2f) - glyphBounds.Y - verticalOffset));

        graphics.DrawString(glyph, font, brush, drawPoint, format);
    }

    private static RectangleF MeasureGlyphBounds(Graphics graphics, string glyph, Font font)
    {
        using var format = (StringFormat)StringFormat.GenericTypographic.Clone();
        format.FormatFlags |= StringFormatFlags.NoClip;
        format.SetMeasurableCharacterRanges(new[] { new CharacterRange(0, glyph.Length) });

        var layoutRect = new RectangleF(0f, 0f, font.Size * 4f, font.Size * 4f);
        var regions = graphics.MeasureCharacterRanges(glyph, font, layoutRect, format);
        if (regions.Length == 0)
            return RectangleF.Empty;

        using var region = regions[0];
        return region.GetBounds(graphics);
    }

    private static int ResolveIconSize(uint dpi)
    {
        var width = NativeTrayMethods.GetSystemMetricsForDpi(SmCxSmIcon, dpi);
        var height = NativeTrayMethods.GetSystemMetricsForDpi(SmCySmIcon, dpi);
        if (width <= 0 || height <= 0)
        {
            width = NativeTrayMethods.GetSystemMetrics(SmCxSmIcon);
            height = NativeTrayMethods.GetSystemMetrics(SmCySmIcon);
        }

        return Math.Max(MinimumIconSize, Math.Max(width, height));
    }

    private static uint ResolveTaskbarDpi()
    {
        var taskbarHandle = NativeTrayMethods.FindWindow("Shell_TrayWnd", null);
        if (taskbarHandle != IntPtr.Zero)
        {
            var dpi = NativeTrayMethods.GetDpiForWindow(taskbarHandle);
            if (dpi != 0)
                return dpi;
        }

        return 96;
    }

    private static TrayIconTheme ResolveTaskbarTheme()
    {
        if (IsHighContrastEnabled())
            return TrayIconTheme.HighContrast;

        return IsTaskbarUsingLightTheme()
            ? TrayIconTheme.Light
            : TrayIconTheme.Dark;
    }

    private static bool IsHighContrastEnabled()
    {
        var highContrast = new HighContrastInfo
        {
            cbSize = Marshal.SizeOf<HighContrastInfo>()
        };

        return NativeTrayMethods.SystemParametersInfo(
            SpiGetHighContrast,
            (uint)highContrast.cbSize,
            ref highContrast,
            0) &&
            (highContrast.dwFlags & HcfHighContrastOn) != 0;
    }

    private static bool IsTaskbarUsingLightTheme()
    {
        using var personalizeKey = Registry.CurrentUser.OpenSubKey(SystemPersonalizeKeyPath);
        var value = personalizeKey?.GetValue(SystemUsesLightThemeValueName);
        return value switch
        {
            int intValue => intValue != 0,
            long longValue => longValue != 0,
            _ => false
        };
    }

    private static Color ResolveHighContrastTextColor()
    {
        var colorValue = NativeTrayMethods.GetSysColor(ColorWindowText);
        return Color.FromArgb(
            (int)(colorValue & 0xFF),
            (int)((colorValue >> 8) & 0xFF),
            (int)((colorValue >> 16) & 0xFF));
    }

    private static Color ResolveFallbackGlyphColor(TrayIconTheme theme)
    {
        return theme switch
        {
            TrayIconTheme.Light => Color.Black,
            TrayIconTheme.HighContrast => ResolveHighContrastTextColor(),
            _ => Color.FromArgb(246, 246, 246)
        };
    }

    private static FontFamily ResolveFontFamily()
    {
        foreach (var family in FontFamily.Families)
        {
            if (string.Equals(family.Name, "Segoe Fluent Icons", StringComparison.OrdinalIgnoreCase))
                return family;
        }

        return new FontFamily("Segoe MDL2 Assets");
    }

    private static string GetFallbackGlyph(TrayVolumeIconKind iconKind)
    {
        return iconKind switch
        {
            TrayVolumeIconKind.Muted => "\uE74F",
            TrayVolumeIconKind.ZeroBars => "\uE992",
            TrayVolumeIconKind.Low => "\uE993",
            TrayVolumeIconKind.Medium => "\uE994",
            TrayVolumeIconKind.NoDevice => "\uE74F",
            _ => "\uE767"
        };
    }

    private static int GetSystemIconResourceId(TrayVolumeIconKind iconKind)
    {
        return iconKind switch
        {
            TrayVolumeIconKind.Muted => 120,
            TrayVolumeIconKind.ZeroBars => 121,
            TrayVolumeIconKind.Low => 122,
            TrayVolumeIconKind.Medium => 123,
            TrayVolumeIconKind.High => 124,
            TrayVolumeIconKind.NoDevice => 125,
            _ => 124
        };
    }

    private readonly record struct IconCacheKey(TrayVolumeIconKind IconKind, uint Dpi, TrayIconTheme Theme);
    private readonly record struct TrayIconEnvironment(uint Dpi, TrayIconTheme Theme);

    private enum TrayIconTheme
    {
        Dark,
        Light,
        HighContrast
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct HighContrastInfo
    {
        public int cbSize;
        public int dwFlags;
        public IntPtr lpszDefaultScheme;
    }

    private static class NativeTrayMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        public static extern int GetSystemMetricsForDpi(int nIndex, uint dpi);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int LookupIconIdFromDirectoryEx(IntPtr presbits, bool fIcon, int cxDesired, int cyDesired, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr CreateIconFromResourceEx(IntPtr presbits, uint dwResSize, bool fIcon, uint dwVer, int cxDesired, int cyDesired, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref HighContrastInfo pvParam, uint fWinIni);

        [DllImport("user32.dll")]
        public static extern uint GetSysColor(int nIndex);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "LoadLibraryExW")]
        public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "FindResourceW")]
        public static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LockResource(IntPtr hResData);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);
    }
}
