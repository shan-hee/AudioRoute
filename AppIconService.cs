using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage;
using Windows.Storage.FileProperties;

namespace AudioRoute;

public static class AppIconService
{
    private static readonly TimeSpan IconCacheDuration = TimeSpan.FromMinutes(10);
    private const int MaxIconCacheEntries = 128;
    private static readonly ExpiringCache<IconCacheKey, Task<BitmapImage?>> IconCache = new(IconCacheDuration, MaxIconCacheEntries);

    public static Task<BitmapImage?> TryLoadIconAsync(string executablePath, uint size = 32)
    {
        if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            return Task.FromResult<BitmapImage?>(null);

        var normalizedPath = Path.GetFullPath(executablePath);
        var cacheKey = new IconCacheKey(normalizedPath, size);
        if (IconCache.TryGetValue(cacheKey, out var cachedIcon))
            return cachedIcon!;

        var loadTask = LoadIconAsync(normalizedPath, size);
        IconCache.Set(cacheKey, loadTask);
        return loadTask;
    }

    private static async Task<BitmapImage?> LoadIconAsync(string executablePath, uint size)
    {
        try
        {
            var file = await StorageFile.GetFileFromPathAsync(executablePath);
            using var thumbnail = await file.GetThumbnailAsync(
                ThumbnailMode.SingleItem,
                size,
                ThumbnailOptions.UseCurrentScale);

            if (thumbnail is null)
                return null;

            var bitmap = new BitmapImage();
            await bitmap.SetSourceAsync(thumbnail);
            return bitmap;
        }
        catch
        {
            return null;
        }
    }

    private readonly record struct IconCacheKey(string ExecutablePath, uint Size);
}
