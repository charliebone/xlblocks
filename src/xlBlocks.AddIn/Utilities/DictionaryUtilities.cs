namespace XlBlocks.AddIn.Utilities;

using System;
using System.Collections;
using System.Collections.Generic;

internal static class DictionaryUtilities
{
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
        [typeof(float)] = DictionaryGenerator<float>(),
        [typeof(bool)] = DictionaryGenerator<bool>(),
        [typeof(DateTime)] = DictionaryGenerator<DateTime>(),
        [typeof(int)] = DictionaryGenerator<int>(),
        [typeof(uint)] = DictionaryGenerator<uint>(),
        [typeof(long)] = DictionaryGenerator<long>(),
        [typeof(short)] = DictionaryGenerator<short>(),
        [typeof(ushort)] = DictionaryGenerator<ushort>(),
        [typeof(sbyte)] = DictionaryGenerator<sbyte>(),
        [typeof(byte)] = DictionaryGenerator<byte>(),
        [typeof(char)] = DictionaryGenerator<char>(),
        [typeof(decimal)] = DictionaryGenerator<decimal>()
    };

    public static IDictionary BuildTypedDictionary(Type keyType)
    {
        return _generatorDelegates[keyType](null);
    }

    public static IDictionary BuildTypedDictionary(Type keyType, IDictionary dictionary)
    {
        return _generatorDelegates[keyType](dictionary);
    }

    public static IDictionary BuildTypedDictionary(Type keyType, IEnumerable<(object Key, object Value)> keyValues)
    {
        var dict = _generatorDelegates[keyType](null);
        foreach (var (Key, Value) in keyValues)
            dict.Add(Key, Value);

        return dict;
    }

}
