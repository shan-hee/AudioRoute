using System;
using System.Collections.Generic;

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
            List<TKey>? keysToRemove = null;
            foreach (var key in entries.Keys)
            {
                if (predicate(key))
                {
                    keysToRemove ??= new List<TKey>();
                    keysToRemove.Add(key);
                }
            }

            if (keysToRemove is not null)
            {
                foreach (var key in keysToRemove)
                    entries.Remove(key);
            }
        }
    }

    public void Clear()
    {
        lock (syncRoot)
            entries.Clear();
    }

    private void PruneCore(DateTimeOffset now)
    {
        if (entries.Count <= capacity)
            return;

        List<TKey>? expiredKeys = null;
        foreach (var pair in entries)
        {
            if (pair.Value.ExpiresAt <= now)
            {
                expiredKeys ??= new List<TKey>();
                expiredKeys.Add(pair.Key);
            }
        }

        if (expiredKeys is not null)
        {
            foreach (var key in expiredKeys)
                entries.Remove(key);
        }

        if (entries.Count <= capacity)
            return;

        var removeCount = entries.Count - capacity;
        var stamps = new long[entries.Count];
        var index = 0;
        foreach (var pair in entries)
            stamps[index++] = pair.Value.LastAccessStamp;

        Array.Sort(stamps);
        var threshold = stamps[removeCount - 1];

        List<TKey>? overflowKeys = null;
        var removed = 0;
        foreach (var pair in entries)
        {
            if (removed >= removeCount)
                break;

            if (pair.Value.LastAccessStamp <= threshold)
            {
                overflowKeys ??= new List<TKey>();
                overflowKeys.Add(pair.Key);
                removed++;
            }
        }

        if (overflowKeys is not null)
        {
            foreach (var key in overflowKeys)
                entries.Remove(key);
        }
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
