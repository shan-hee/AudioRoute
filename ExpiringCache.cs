using System;
using System.Collections.Generic;
using System.Linq;

namespace AudioRoute;

internal sealed class ExpiringCache<TKey, TValue> where TKey : notnull
{
    private readonly object syncRoot = new();
    private readonly Dictionary<TKey, CacheEntry> entries;
    private readonly TimeSpan defaultLifetime;
    private readonly int capacity;
    private long accessStamp;

    public ExpiringCache(TimeSpan defaultLifetime, int capacity, IEqualityComparer<TKey>? comparer = null)
    {
        if (capacity <= 0)
            throw new ArgumentOutOfRangeException(nameof(capacity));

        this.defaultLifetime = defaultLifetime;
        this.capacity = capacity;
        entries = new Dictionary<TKey, CacheEntry>(comparer ?? EqualityComparer<TKey>.Default);
    }

    public bool TryGetValue(TKey key, out TValue? value)
    {
        var now = DateTimeOffset.UtcNow;

        lock (syncRoot)
        {
            if (entries.TryGetValue(key, out var entry))
            {
                if (entry.ExpiresAt > now)
                {
                    entry.LastAccessStamp = ++accessStamp;
                    value = entry.Value;
                    return true;
                }

                entries.Remove(key);
            }
        }

        value = default;
        return false;
    }

    public void Set(TKey key, TValue value)
    {
        Set(key, value, defaultLifetime);
    }

    public void Set(TKey key, TValue value, TimeSpan lifetime)
    {
        var now = DateTimeOffset.UtcNow;

        lock (syncRoot)
        {
            entries[key] = new CacheEntry(value, now + lifetime, ++accessStamp);
            PruneCore(now);
        }
    }

    public void Remove(TKey key)
    {
        lock (syncRoot)
            entries.Remove(key);
    }

    public void RemoveWhere(Func<TKey, bool> predicate)
    {
        lock (syncRoot)
        {
            var keysToRemove = entries.Keys.Where(predicate).ToArray();
            foreach (var key in keysToRemove)
                entries.Remove(key);
        }
    }

    public void Clear()
    {
        lock (syncRoot)
            entries.Clear();
    }

    private void PruneCore(DateTimeOffset now)
    {
        if (entries.Count == 0)
            return;

        var expiredKeys = entries
            .Where(pair => pair.Value.ExpiresAt <= now)
            .Select(pair => pair.Key)
            .ToArray();

        foreach (var key in expiredKeys)
            entries.Remove(key);

        if (entries.Count <= capacity)
            return;

        var overflowKeys = entries
            .OrderBy(pair => pair.Value.LastAccessStamp)
            .Take(entries.Count - capacity)
            .Select(pair => pair.Key)
            .ToArray();

        foreach (var key in overflowKeys)
            entries.Remove(key);
    }

    private sealed class CacheEntry
    {
        public CacheEntry(TValue value, DateTimeOffset expiresAt, long lastAccessStamp)
        {
            Value = value;
            ExpiresAt = expiresAt;
            LastAccessStamp = lastAccessStamp;
        }

        public TValue Value { get; }

        public DateTimeOffset ExpiresAt { get; }

        public long LastAccessStamp { get; set; }
    }
}
