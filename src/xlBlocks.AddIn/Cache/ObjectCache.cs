namespace XlBlocks.AddIn.Cache;

using System;
using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Murmur;
using Numeral;
using XlBlocks.AddIn.Types;

public delegate void CacheInvalidatedEventHandler(object sender, EventArgs e);

internal record ObjectCacheItemInfo(string Reference, string HexKey, Type Type);

internal class ObjectCache
{
    private static readonly Encoding _encoding = Encoding.UTF8;
    private static readonly uint _hashSeed = 12387;
    private static readonly string _affix = "@";
    private static readonly string _compoundKeySplit = "_";
    private static readonly Regex _hexRegex = new($"\\{_affix}([0-9A-Fa-f]{{32}}(?:_[0-9A-Fa-f]{{32}})?)\\{_affix}",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    [ThreadStatic]
    private static HashAlgorithm? _murmur128;

    private readonly ConcurrentDictionary<byte[], object> _cache = new(new ByteArrayComparer());
    private readonly ConcurrentDictionary<byte[], string> _referenceCache = new(new ByteArrayComparer());

    public event CacheInvalidatedEventHandler? CacheInvalidated;

    protected virtual void OnCacheInvalidated(EventArgs e)
    {
        CacheInvalidated?.Invoke(this, e);
    }

    private static string HexFromHash(byte[] hash)
    {
        return $"{_affix}{HexConverter.GetString(hash)}{_affix}";
    }

    internal static byte[] GetCacheKey(string reference, out string hexKey)
    {
        _murmur128 ??= MurmurHash.Create128(_hashSeed, managed: false);

        var hash = _murmur128.ComputeHash(_encoding.GetBytes(reference));
        hexKey = HexFromHash(hash);
        return hash;
    }

    private static bool TryCacheKeyFromHexKey(string hexKey, out byte[] hash)
    {
        if (hexKey.Contains(_compoundKeySplit))
        {
            var splitHexKey = hexKey.Split(_compoundKeySplit);
            if (TryCacheKeyFromHexKey(splitHexKey[0], out var hash1) && TryCacheKeyFromHexKey(splitHexKey[1], out var hash2))
            {
                hash = new byte[32];
                hash1.CopyTo(hash, 0);
                hash2.CopyTo(hash, 16);
                return true;
            }
            else
            {
                hash = Array.Empty<byte>();
                return false;
            }
        }

        var hexMatch = _hexRegex.Match(hexKey);
        if (!hexMatch.Success)
        {
            hash = Array.Empty<byte>();
            return false;
        }

        var purHexStr = hexMatch.Groups[1].Value;
        hash = HexConverter.GetBytes(purHexStr);
        return true;
    }

    internal static byte[] GetCacheCollectionKey(byte[] referenceHash, string referenceHex, string innerKey, out string hexKey)
    {
        var innerKeyHash = GetCacheKey(innerKey, out var innerKeyHex);
        hexKey = $"{referenceHex}{_compoundKeySplit}{innerKeyHex}";

        var fullHash = new byte[referenceHash.Length + innerKeyHash.Length];
        referenceHash.CopyTo(fullHash, 0);
        innerKeyHash.CopyTo(fullHash.AsSpan(referenceHash.Length));
        return fullHash;
    }

    internal static byte[] GetCacheCollectionKey(byte[] referenceHash, string referenceHex, int innerIndex, out string hexKey)
    {
        var innerKeyHash = GetCacheKey(innerIndex.ToString(), out var innerKeyHex);
        hexKey = $"{referenceHex}{_compoundKeySplit}{innerKeyHex}";

        var fullHash = new byte[referenceHash.Length + innerKeyHash.Length];
        referenceHash.CopyTo(fullHash, 0);
        innerKeyHash.CopyTo(fullHash.AsSpan(referenceHash.Length));
        return fullHash;
    }

    public int Count => _cache.Count;

    public IList<ObjectCacheItemInfo> GetItemInfo()
    {
        return _cache.Select((kvp) =>
        {
            return new ObjectCacheItemInfo(_referenceCache[kvp.Key], HexFromHash(kvp.Key), _cache[kvp.Key].GetType());
        }).ToList();
    }

    public void AddOrReplace(string reference, object value, out string hexKey)
    {
        var hash = GetCacheKey(reference, out hexKey);
        _cache[hash] = value;
        _referenceCache[hash] = reference;
        OnCacheInvalidated(EventArgs.Empty);
    }

    internal sealed class ElementCacher : IElementCacher
    {
        private readonly string _reference;
        private readonly byte[] _referenceHash;
        private readonly string _referenceHex;
        private readonly ObjectCache _objectCache;

        public string CollectionHexKey => _referenceHex;

        public ElementCacher(string reference, ObjectCache cache)
        {
            _reference = reference;
            _referenceHash = GetCacheKey(reference, out _referenceHex);
            _objectCache = cache;
        }

        public string CacheElement(object element, int index)
        {
            var elementHash = GetCacheCollectionKey(_referenceHash, _referenceHex, index, out var hexKey);
            _objectCache._cache[elementHash] = element;
            _objectCache._referenceCache[elementHash] = $"{_reference}_{index}";
            _objectCache.OnCacheInvalidated(EventArgs.Empty);
            return hexKey;
        }

        public string CacheElement(object element, string key)
        {
            var elementHash = GetCacheCollectionKey(_referenceHash, _referenceHex, key, out var hexKey);
            _objectCache._cache[elementHash] = element;
            _objectCache._referenceCache[elementHash] = $"{_reference}_{key}";
            _objectCache.OnCacheInvalidated(EventArgs.Empty);
            return hexKey;
        }
    }

    public ElementCacher GetElementCacher(string reference)
    {
        return new ElementCacher(reference, this);
    }

    public bool ContainsKey(string hexKey)
    {
        if (hexKey is not null && TryCacheKeyFromHexKey(hexKey, out var hash))
            return _cache.ContainsKey(hash);

        return false;
    }

    public bool Remove(string hexKey)
    {
        return Remove(hexKey, out var _);
    }

    public bool Remove(string hexKey, out object? value)
    {
        value = default;
        if (hexKey is not null && TryCacheKeyFromHexKey(hexKey, out var hash))
        {
            var success = _cache.Remove(hash, out value);
            OnCacheInvalidated(EventArgs.Empty);
            return success;
        }

        return false;
    }

    public bool TryGetValue(string hexKey, out object? value)
    {
        value = default;
        if (hexKey != null && TryCacheKeyFromHexKey(hexKey, out var hash))
            return _cache.TryGetValue(hash, out value);

        return false;
    }

    public bool TryGetReference(string hexKey, out string? reference)
    {
        reference = default;
        if (hexKey != null && TryCacheKeyFromHexKey(hexKey, out var hash))
            return _referenceCache.TryGetValue(hash, out reference);

        return false;
    }

    public void Clear()
    {
        _cache.Clear();
        OnCacheInvalidated(EventArgs.Empty);
    }
}

internal sealed class ByteArrayComparer : EqualityComparer<byte[]>
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool SpanEquals(ReadOnlySpan<byte> a, ReadOnlySpan<byte> b)
    {
        return a.SequenceEqual(b);
    }

    public override bool Equals(byte[]? first, byte[]? second)
    {
        if (first == null || second == null)
        {
            return first == second;
        }

        if (ReferenceEquals(first, second))
        {
            return true;
        }

        if (first.Length != second.Length)
        {
            return false;
        }

        return SpanEquals(first, second);
    }

    public override int GetHashCode(byte[] obj)
    {
        if (obj == null)
            throw new ArgumentNullException(nameof(obj));

        // a four byte comparison should segment a lot of the hashes before resorting to Equals
        // note cached collection elements will have 32-byte keys which will all share first 16 bytes
        var offset = 16 * Math.Max((obj.Length / 16) - 1, 0);
        return obj[offset] | (obj[offset + 1] << 8) | (obj[offset + 2] << 16) | (obj[offset + 3] << 24);
    }
}
