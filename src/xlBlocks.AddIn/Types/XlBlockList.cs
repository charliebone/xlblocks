namespace XlBlocks.AddIn.Types;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlBlocks.AddIn.Dna;

internal class XlBlockList : IXlBlockCopyableObject<XlBlockList>, IXlBlockArrayableObject, IXlBlockCacheableCollection, IEnumerable
{
    private readonly IList _list;

    private readonly Type _dataType;
    public Type DataType => _dataType;

    public int Count => _list.Count;

    private static Func<IList?, List<T>> ListGenerator<T>()
    {
        return (IList? list) => list is null ?
            new List<T>() :
            new List<T>(list.Cast<T>());
    }

    private static readonly Dictionary<Type, Func<IList?, IList>> _generatorDelegates = new()
    {
        [typeof(object)] = ListGenerator<object>(),
        [typeof(string)] = ListGenerator<string>(),
        [typeof(double)] = ListGenerator<double>(),
        [typeof(float)] = ListGenerator<float>(),
        [typeof(bool)] = ListGenerator<bool>(),
        [typeof(DateTime)] = ListGenerator<DateTime>(),
        [typeof(int)] = ListGenerator<int>(),
        [typeof(uint)] = ListGenerator<uint>(),
        [typeof(long)] = ListGenerator<long>(),
        [typeof(ulong)] = ListGenerator<ulong>(),
        [typeof(short)] = ListGenerator<short>(),
        [typeof(ushort)] = ListGenerator<ushort>(),
        [typeof(sbyte)] = ListGenerator<sbyte>(),
        [typeof(byte)] = ListGenerator<byte>(),
        [typeof(decimal)] = ListGenerator<decimal>()
    };

    private static Func<bool, Action<IList>> GenerateSortDelegate<T>(Comparison<T> comparison)
    {
        return (descending) => descending ?
            (IList list) => ((List<T>)list).Sort((a, b) => comparison(b, a)) :
            (IList list) => ((List<T>)list).Sort((a, b) => comparison(a, b));
    }

    private static readonly Dictionary<Type, Func<bool, Action<IList>>> _sortDelegates = new()
    {
        [typeof(object)] = GenerateSortDelegate<object>((a, b) => string.Compare(a.ToString(), b.ToString())),
        [typeof(string)] = GenerateSortDelegate<string>(Comparer<string>.Default.Compare),
        [typeof(double)] = GenerateSortDelegate<double>(Comparer<double>.Default.Compare),
        [typeof(float)] = GenerateSortDelegate<float>(Comparer<float>.Default.Compare),
        [typeof(bool)] = GenerateSortDelegate<bool>(Comparer<bool>.Default.Compare),
        [typeof(DateTime)] = GenerateSortDelegate<DateTime>(Comparer<DateTime>.Default.Compare),
        [typeof(int)] = GenerateSortDelegate<int>(Comparer<int>.Default.Compare),
        [typeof(uint)] = GenerateSortDelegate<uint>(Comparer<uint>.Default.Compare),
        [typeof(long)] = GenerateSortDelegate<long>(Comparer<long>.Default.Compare),
        [typeof(ulong)] = GenerateSortDelegate<ulong>(Comparer<ulong>.Default.Compare),
        [typeof(short)] = GenerateSortDelegate<short>((a, b) => a.CompareTo(b)),
        [typeof(ushort)] = GenerateSortDelegate<ushort>(Comparer<ushort>.Default.Compare),
        [typeof(sbyte)] = GenerateSortDelegate<sbyte>(Comparer<sbyte>.Default.Compare),
        [typeof(byte)] = GenerateSortDelegate<byte>(Comparer<byte>.Default.Compare),
        [typeof(decimal)] = GenerateSortDelegate<decimal>(Comparer<decimal>.Default.Compare)
    };

    private static Func<IList, int, IList> GenerateTakeSkipDelegate<T>()
    {
        return (IList list, int n) => n > 0 ?
            list.Cast<T>().Take(n).ToList() :
            list.Cast<T>().Skip(-n).ToList();
    }

    private static readonly Dictionary<Type, Func<IList, int, IList>> _takeSkipDelegate = new()
    {
        [typeof(object)] = GenerateTakeSkipDelegate<object>(),
        [typeof(string)] = GenerateTakeSkipDelegate<string>(),
        [typeof(double)] = GenerateTakeSkipDelegate<double>(),
        [typeof(float)] = GenerateTakeSkipDelegate<float>(),
        [typeof(bool)] = GenerateTakeSkipDelegate<bool>(),
        [typeof(DateTime)] = GenerateTakeSkipDelegate<DateTime>(),
        [typeof(int)] = GenerateTakeSkipDelegate<int>(),
        [typeof(uint)] = GenerateTakeSkipDelegate<uint>(),
        [typeof(long)] = GenerateTakeSkipDelegate<long>(),
        [typeof(ulong)] = GenerateTakeSkipDelegate<ulong>(),
        [typeof(short)] = GenerateTakeSkipDelegate<short>(),
        [typeof(ushort)] = GenerateTakeSkipDelegate<ushort>(),
        [typeof(sbyte)] = GenerateTakeSkipDelegate<sbyte>(),
        [typeof(byte)] = GenerateTakeSkipDelegate<byte>(),
        [typeof(decimal)] = GenerateTakeSkipDelegate<decimal>()
    };

    internal XlBlockList(IList list)
    {
        _list = list; // we just wrap the list we are given in lieu of copying

        var listType = list.GetType();
        if (!listType.IsGenericType)
            throw new ArgumentException("list must be a generic List<>");
        _dataType = listType.GenericTypeArguments[0];
    }

    public IEnumerator GetEnumerator() => _list.GetEnumerator();

    public object? this[int index] => _list[index];

    public bool Contains(object item) => _list.Contains(item);

    public XlBlockList Copy()
    {
        var list = _generatorDelegates[_dataType](_list);
        return new XlBlockList(list);
    }

    public override string ToString()
    {
        return $"An XlBlocks list of type '{DataType.Name}' with {Count} items";
    }

    public void Add(object item)
    {
        if (!ParamTypeConverter.TryConvertToProvidedType(item, _dataType, out var convertedItem))
            throw new ArgumentException($"Item '{item}' cannot be added to the list of type '{_dataType.Name}'");

        _list.Add(convertedItem);
    }

    public void Add(XlBlockRange items, string onErrors)
    {
        foreach (var item in items.AsCleanEnumerable(onErrors))
            Add(item);
    }

    public void Remove(object item)
    {
        _list.Remove(item);
    }

    public void Remove(XlBlockRange items)
    {
        foreach (var item in items.AsCleanEnumerable(XlBlockRange.CleanBehavior.DropErrors))
            _list.Remove(item);
    }

    public void RemoveAt(int index)
    {
        _list.RemoveAt(index);
    }

    public object[] Get() => _list.Cast<object>().ToArray();

    public object? GetAt(int index) => _list[index];

    public void Sort(bool descending)
    {
        try
        {
            _sortDelegates[_dataType](descending)(_list);
        }
        catch
        {

        }
    }

    public XlBlockList Take(int n)
    {
        if (n < 0) throw new ArgumentOutOfRangeException(nameof(n), $"{nameof(n)} must be greater than or equal to zero");

        return new XlBlockList(_takeSkipDelegate[_dataType](_list, n));
    }

    public XlBlockList Skip(int n)
    {
        if (n < 0) throw new ArgumentOutOfRangeException(nameof(n), $"{nameof(n)} must be greater than or equal to zero");

        return new XlBlockList(_takeSkipDelegate[_dataType](_list, -n));
    }

    public XlBlockList GetUniqueItems()
    {
        var uniqueList = _generatorDelegates[_dataType](null);
        var uniqueSet = new HashSet<object>(_list.Cast<object>());

        foreach (var item in uniqueSet)
            uniqueList.Add(item);

        return new XlBlockList(uniqueList);
    }

    public XlBlockList GetDuplicateItems()
    {
        var duplicateList = _generatorDelegates[_dataType](null);

        var uniqueSet = new HashSet<object>();
        var duplicateSet = new HashSet<object>();

        foreach (var item in _list)
        {
            if (!uniqueSet.Add(item) && duplicateSet.Add(item))
            {
                duplicateList.Add(item);
            }
        }

        return new XlBlockList(duplicateList);
    }

    public string ContentsToString(string separator)
    {
        return string.Join(separator, _list.Cast<object>());
    }

    public object[,] AsArray(RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        var outputArray = orientation == RangeOrientation.ByColumn ?
            new object[_list.Count, 1] : new object[1, _list.Count];

        var i = 0;
        foreach (var item in _list)
        {
            if (orientation == RangeOrientation.ByColumn)
            {
                outputArray[i, 0] = item;
            }
            else
            {
                outputArray[0, i] = item;
            }
            i++;
        }
        return outputArray;
    }

    public IXlBlockCacheableCollection CacheCollectionValues(IElementCacher elementCacher)
    {
        var cacheKeyList = new List<object>();
        for (var i = 0; i < _list.Count; i++)
        {
            var val = _list[i];
            if (val is null || !ShouldCacheElement(val))
            {
                cacheKeyList.Add(val!);
                continue;
            }

            var elementHexKey = elementCacher.CacheElement(val, i);
            cacheKeyList.Add(elementHexKey);
        }
        return new XlBlockList(cacheKeyList);
    }

    public bool ShouldCacheElement(object? element)
    {
        // don't cache items that are value types or strings
        return element is not null && !element.GetType().IsValueType && element is not string;
    }

    public static XlBlockList Build(XlBlockRange range, string onErrors)
    {
        var list = range.Clean(onErrors).ToList();
        return new XlBlockList(list);
    }

    public static XlBlockList BuildTyped(XlBlockRange range, string type, string onErrors)
    {
        var desiredType = ParamTypeConverter.StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
        var typedList = _generatorDelegates[desiredType](null);

        var list = range.Clean(onErrors).ToList();
        foreach (var result in list.ConvertToProvidedType(type))
        {
            if (!result.Success)
                throw new ArgumentException($"Item '{result.Input}' cannot be converted to type '{desiredType.Name}'");

            typedList.Add(result.ConvertedInput);
        }

        return new XlBlockList(typedList);
    }

    internal static XlBlockList BuildTyped(IEnumerable<object> items, Type type)
    {
        var typedList = _generatorDelegates[type](null);
        foreach (var item in items)
            typedList.Add(item);

        return new XlBlockList(typedList);
    }

    public static XlBlockList BuildFromString(string str, string delimiter, bool trimStrings = false, bool ignoreEmpty = true)
    {
        var list = str.Split(new[] { delimiter }, StringSplitOptions.None)
            .Where(x => !ignoreEmpty || !string.IsNullOrEmpty(x))
            .Select(x => trimStrings ? x.Trim() : x).ToList();

        return new XlBlockList(list);
    }

    public static XlBlockList MergeLists(params XlBlockList[] lists)
    {
        var types = lists.Select(x => x._dataType).Distinct().ToArray();
        if (types.Length > 1)
            throw new ArgumentException("All lists must be of the same type");

        var combinedList = _generatorDelegates[types[0]](null);
        var hashSet = new HashSet<object>();

        foreach (var list in lists)
        {
            foreach (var item in list)
            {
                if (hashSet.Add(item))
                {
                    combinedList.Add(item);
                }
            }
        }

        return new XlBlockList(combinedList);
    }

    public static XlBlockList UnifyLists(params XlBlockList[] lists)
    {
        var types = lists.Select(x => x._dataType).Distinct().ToArray();
        if (types.Length > 1)
            throw new ArgumentException("All lists must be of the same type");

        var combinedList = _generatorDelegates[types[0]](null);

        for (var i = 0; i < lists.Length; i++)
            for (var j = 0; j < lists[i].Count; j++)
                combinedList.Add(lists[i][j]);

        return new XlBlockList(combinedList);
    }

    public static XlBlockList IntersectLists(params XlBlockList[] lists)
    {
        var types = lists.Select(x => x._dataType).Distinct().ToArray();
        if (types.Length > 1)
            throw new ArgumentException("All lists must be of the same type");

        var combinedList = _generatorDelegates[types[0]](null);
        var sortedLists = lists.OrderBy(x => x.Count).ToList();
        var intersection = new HashSet<object>(sortedLists[0].Cast<object>());

        for (var i = 1; i < sortedLists.Count; i++)
        {
            var list = sortedLists[i];
            intersection.IntersectWith(sortedLists[i].Cast<object>());
            if (intersection.Count == 0)
                return new XlBlockList(combinedList);
        }

        foreach (var item in intersection)
            combinedList.Add(item);

        return new XlBlockList(combinedList);
    }
}
