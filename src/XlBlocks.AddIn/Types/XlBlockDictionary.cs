namespace XlBlocks.AddIn.Types;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Utilities;

internal class XlBlockDictionary : IXlBlockCopyableObject<XlBlockDictionary>, IXlBlockArrayableObject, IXlBlockCacheableCollection, IEnumerable
{
    private readonly IDictionary _dictionary;

    private readonly Type _keyType;
    public Type KeyType => _keyType;

    public int Count => _dictionary.Count;

    private static Func<IDictionary?, Dictionary<T, object>> DictionaryGenerator<T>() where T : notnull
    {
        return (IDictionary? dict) => dict is null ?
            new Dictionary<T, object>() :
            new Dictionary<T, object>((IDictionary<T, object>)dict);
    }

    private static readonly Dictionary<Type, Func<IDictionary?, IDictionary>> _generatorDelegates = new()
    {
        [typeof(object)] = DictionaryGenerator<object>(),
        [typeof(string)] = DictionaryGenerator<string>(),
        [typeof(double)] = DictionaryGenerator<double>(),
        [typeof(bool)] = DictionaryGenerator<bool>(),
        [typeof(DateTime)] = DictionaryGenerator<DateTime>(),
        [typeof(int)] = DictionaryGenerator<int>(),
        [typeof(uint)] = DictionaryGenerator<uint>(),
        [typeof(long)] = DictionaryGenerator<long>(),
        [typeof(short)] = DictionaryGenerator<short>(),
        [typeof(ushort)] = DictionaryGenerator<ushort>(),
        [typeof(decimal)] = DictionaryGenerator<decimal>()
    };

    internal XlBlockDictionary(IEnumerable<(object Key, object Value)> keyValues, Type keyType)
    {
        _keyType = keyType;
        _dictionary = DictionaryUtilities.BuildTypedDictionary(keyType, keyValues);
    }

    internal XlBlockDictionary(IDictionary dictionary, Type keyType)
    {
        _keyType = keyType;
        _dictionary = dictionary;
    }

    private XlBlockDictionary(Type keyType)
    {
        _keyType = keyType;
        _dictionary = DictionaryUtilities.BuildTypedDictionary(keyType);
    }

    public IEnumerator GetEnumerator() => _dictionary.GetEnumerator();

    public XlBlockDictionary Copy()
    {
        var dictionary = DictionaryUtilities.BuildTypedDictionary(_keyType, _dictionary);
        return new XlBlockDictionary(dictionary, _keyType);
    }

    public void AddOrReplace(object key, object value)
    {
        if (!ParamTypeConverter.TryConvertToProvidedType(key, _keyType, out var convertedKey))
            throw new ArgumentException($"Key '{key}' cannot be added to dictionary with key type of '{_keyType.Name}'");

        _dictionary[convertedKey] = value;
    }

    public void AddOrReplace(XlBlockRange keys, XlBlockRange values, string onErrors)
    {
        if (keys.Count != values.Count)
            throw new ArgumentException("keys and values must have same length");

        var keyValues = CheckKeysWithValues(keys, values, onErrors);
        foreach (var (Key, Value) in keyValues)
            AddOrReplace(Key, Value);
    }

    public void Remove(object key)
    {
        if (ParamTypeConverter.TryConvertToProvidedType(key, _keyType, out var convertedKey))
            _dictionary.Remove(convertedKey);
    }

    public void Remove(XlBlockRange keys)
    {
        foreach (var key in keys)
            Remove(key);
    }

    public object? this[object key]
    {
        get
        {
            if (!ParamTypeConverter.TryConvertToProvidedType(key, _keyType, out var convertedKey))
                throw new ArgumentException($"key '{key}' is not of type '{_keyType.Name}'");

            return _dictionary[convertedKey];
        }
        set
        {
            if (!ParamTypeConverter.TryConvertToProvidedType(key, _keyType, out var convertedKey))
                return;
            _dictionary[convertedKey] = value;
        }
    }

    public object[] Keys => _dictionary.Keys.Cast<object>().ToArray();

    public object[] Values => _dictionary.Values.Cast<object>().ToArray();

    public bool ContainsKey(object key)
    {
        return ParamTypeConverter.TryConvertToProvidedType(key, _keyType, out var convertedKey) && _dictionary.Contains(convertedKey);
    }

    public override string ToString()
    {
        return $"An XlBlocks dictionary with key type '{KeyType.Name}' and {Count} entries";
    }

    public object[,] AsArray(RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        return AsArray(false, "Key", "Value", orientation);
    }

    public object[,] AsArray(bool includeHeader, string keyColumn = "Key", string valueColumn = "Value",
        RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        var outputArray = orientation == RangeOrientation.ByColumn ?
            new object[_dictionary.Count + (includeHeader ? 1 : 0), 2] :
            new object[2, _dictionary.Count + (includeHeader ? 1 : 0)];

        var i = 0;
        if (includeHeader)
        {
            outputArray[0, 0] = keyColumn;
            if (orientation == RangeOrientation.ByColumn)
                outputArray[0, 1] = valueColumn;
            else
                outputArray[1, 0] = valueColumn;
            i++;
        }

        foreach (var key in _dictionary.Keys)
        {
            var value = _dictionary[key];
            if (orientation == RangeOrientation.ByColumn)
            {
                outputArray[i, 0] = key;
                if (value is not null)
                    outputArray[i, 1] = value;
            }
            else
            {
                outputArray[0, i] = key;
                if (value is not null)
                    outputArray[1, i] = value;
            }
            i++;
        }
        return outputArray;
    }

    public IXlBlockCacheableCollection CacheCollectionValues(IElementCacher elementCacher)
    {
        if (KeyType != typeof(string))
            return this;

        var cacheKeyDict = new Dictionary<string, object>();
        foreach (var key in _dictionary.Keys.Cast<string>())
        {
            var val = _dictionary[key];
            if (val is null || !ShouldCacheElement(val))
            {
                cacheKeyDict[key] = val!;
                continue;
            }

            var elementHexKey = elementCacher.CacheElement(val, key);
            cacheKeyDict[key] = elementHexKey;
        }
        return new XlBlockDictionary(cacheKeyDict, KeyType);
    }

    public bool ShouldCacheElement(object? element)
    {
        // don't cache items that are value types or strings
        return element is not null && !element.GetType().IsValueType && element is not string;
    }

    private static IEnumerable<(object Key, object Value)> CheckKeysWithValues(IEnumerable<object> keys, IEnumerable<object> values, string onErrors)
    {
        Debug.Assert(keys.Count() == values.Count());

        onErrors = onErrors.ToLowerInvariant();
        if (onErrors != "drop" && onErrors != "error")
            throw new ArgumentException("onErrors must be one of 'drop' or 'error'");

        var keyValueEnumerable = keys.Zip(values);
        if (onErrors == "error")
        {
            if (keyValueEnumerable.Any(kv => kv.First is null || ParamTypeConverter.IsMissingOrError(kv.First)))
                throw new ArgumentException("keys cannot contain missing or invalid items");
        }
        else
        {
            keyValueEnumerable = keyValueEnumerable.Where(kv => kv.First is not null && !ParamTypeConverter.IsMissingOrError(kv.First));
        }

        return keyValueEnumerable;
    }

    public static XlBlockDictionary Build(XlBlockRange keys, XlBlockRange values, string onErrors)
    {
        if (keys.Count != values.Count)
            throw new ArgumentException("keys and values must have same length");

        var guessConversions = keys.ConvertToBestTypes();
        var determinedType = guessConversions.Where(x => !x.IsMissingOrError)
            .Select(x => x.ConvertedType)
            .DetermineBestType();

        var convertedKeys = guessConversions.Select(x =>
        {
            if (x.IsMissingOrError)
                return ExcelError.ExcelErrorNA;

            if (x.Success && x.ConvertedType == determinedType)
                return x.ConvertedInput ?? ExcelError.ExcelErrorNA;

            if (ParamTypeConverter.TryConvertToProvidedType(x.Input, determinedType, out var convertedInput))
                return convertedInput;

            return ExcelError.ExcelErrorNA;
        });

        var keyValues = CheckKeysWithValues(convertedKeys, values, onErrors);

        var dict = DictionaryUtilities.BuildTypedDictionary(determinedType, keyValues);
        return new XlBlockDictionary(dict, determinedType);
    }

    public static XlBlockDictionary BuildTyped(XlBlockRange keys, string keyType, XlBlockRange values, string onErrors)
    {
        var parsedType = ParamTypeConverter.StringToType(keyType) ?? throw new ArgumentException($"unknown type '{keyType}'");

        if (keys.Count != values.Count)
            throw new ArgumentException("keys and values must have same length");

        if (!_generatorDelegates.ContainsKey(parsedType))
            throw new ArgumentException($"type '{keyType}' is not a valid dictionary key type");

        var keyValues = CheckKeysWithValues(keys, values, onErrors)
            .Select(kv =>
            {
                if (!ParamTypeConverter.TryConvertToProvidedType(kv.Key, parsedType, out var converted))
                    throw new ArgumentException($"Key '{kv.Key}' cannot be converted to type '{parsedType.Name}'");

                return (Key: converted, kv.Value);
            });

        var dict = DictionaryUtilities.BuildTypedDictionary(parsedType, keyValues);
        return new XlBlockDictionary(dict, parsedType);
    }

    public static XlBlockDictionary BuildFromLists(XlBlockList keyList, XlBlockList valueList)
    {
        if (keyList.Count != valueList.Count)
            throw new ArgumentException("key list and value list must have equal number of items");

        var newDict = new XlBlockDictionary(keyList.DataType);
        for (var i = 0; i < keyList.Count; i++)
        {
            var key = keyList[i];
            if (key is not null)
                newDict._dictionary.Add(key, valueList[i]);
        }
        return newDict;
    }
}
