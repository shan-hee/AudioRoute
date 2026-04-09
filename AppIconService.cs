using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage;
using Windows.Storage.FileProperties;

namespace AudioRoute;

public static class AppIconService
{
    private static readonly object IconCacheSync = new();
    private static readonly Dictionary<IconCacheKey, CachedIconLoad> IconCache = new();
    private static readonly TimeSpan IconCacheDuration = TimeSpan.FromMinutes(10);
    private const int MaxIconCacheEntries = 128;

    public static Task<BitmapImage?> TryLoadIconAsync(string executablePath, uint size = 32)
    {
        if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            return Task.FromResult<BitmapImage?>(null);

        var normalizedPath = Path.GetFullPath(executablePath);
        var cacheKey = new IconCacheKey(normalizedPath, size);
        var now = DateTimeOffset.UtcNow;

        lock (IconCacheSync)
        {
            if (IconCache.TryGetValue(cacheKey, out var cachedIcon) &&
                cachedIcon.ExpiresAt > now)
            {
                return cachedIcon.LoadTask;
            }

            var loadTask = LoadIconAsync(normalizedPath, size);
            IconCache[cacheKey] = new CachedIconLoad(loadTask, now + IconCacheDuration);

            if (IconCache.Count > MaxIconCacheEntries)
                TrimExpiredIconCacheEntries(now);

            return loadTask;
        }
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

    private static void TrimExpiredIconCacheEntries(DateTimeOffset now)
    {
        var expiredKeys = new List<IconCacheKey>();

        foreach (var entry in IconCache)
        {
            if (entry.Value.ExpiresAt <= now)
                expiredKeys.Add(entry.Key);
        }

        foreach (var key in expiredKeys)
            IconCache.Remove(key);

        if (IconCache.Count <= MaxIconCacheEntries)
            return;

        using var enumerator = IconCache.Keys.GetEnumerator();
        if (enumerator.MoveNext())
            IconCache.Remove(enumerator.Current);
    }

    private readonly record struct IconCacheKey(string ExecutablePath, uint Size);

    private sealed record CachedIconLoad(Task<BitmapImage?> LoadTask, DateTimeOffset ExpiresAt);
}
