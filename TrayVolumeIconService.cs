using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;

namespace AudioRoute;

public static class TrayVolumeIconService
{
    private const int MinimumIconSize = 16;
    private const int SmCxSmIcon = 49;
    private const int SmCySmIcon = 50;
    private const float TargetGlyphScale = 0.9f;
    private static readonly Dictionary<TrayVolumeIconKind, IntPtr> CachedIcons = new();
    private static int cachedIconSize;

    public static IntPtr GetIconHandle(TrayVolumeIconKind iconKind)
    {
        var iconSize = ResolveIconSize();
        if (cachedIconSize != iconSize)
        {
            DisposeCachedIcons();
            cachedIconSize = iconSize;
        }

        if (CachedIcons.TryGetValue(iconKind, out var iconHandle) && iconHandle != IntPtr.Zero)
            return iconHandle;

        iconHandle = CreateIconHandle(iconKind, iconSize);
        CachedIcons[iconKind] = iconHandle;
        return iconHandle;
    }

    public static void Dispose()
    {
        DisposeCachedIcons();
        cachedIconSize = 0;
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

    private static IntPtr CreateIconHandle(TrayVolumeIconKind iconKind, int iconSize)
    {
        using var bitmap = new Bitmap(iconSize, iconSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

        using var brush = new SolidBrush(Color.FromArgb(246, 246, 246));
        graphics.Clear(Color.Transparent);
        DrawGlyph(graphics, brush, GetGlyph(iconKind), iconSize);
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

    private static int ResolveIconSize()
    {
        var width = NativeTrayMethods.GetSystemMetrics(SmCxSmIcon);
        var height = NativeTrayMethods.GetSystemMetrics(SmCySmIcon);
        return Math.Max(MinimumIconSize, Math.Max(width, height));
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

    private static string GetGlyph(TrayVolumeIconKind iconKind)
    {
        return iconKind switch
        {
            TrayVolumeIconKind.Muted => "\uE74F",
            TrayVolumeIconKind.Low => "\uE993",
            TrayVolumeIconKind.Medium => "\uE994",
            _ => "\uE767"
        };
    }

    private static class NativeTrayMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);
    }
}
