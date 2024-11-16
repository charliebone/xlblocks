namespace XlBlocks.AddIn.Types;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;

internal class XlBlockRange : IXlBlockArrayableObject, IEnumerable<object>
{
    private readonly object[,] _rangeArray;

    public int RowCount { get; }

    public int ColumnCount { get; }

    public int Count => RowCount * ColumnCount;

    public bool IsFlat => ColumnCount == 1;

    private XlBlockRange(object[,] objArray)
    {
        RowCount = objArray.GetLength(0);
        ColumnCount = objArray.GetLength(1);

        _rangeArray = objArray;
    }

    public IEnumerator<object> GetEnumerator() => AsEnumerable().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => AsEnumerable().GetEnumerator();

    public object? this[int row, int col] => _rangeArray[row, col];

    private IEnumerable<object> AsEnumerable()
    {
        return _rangeArray.AsEnumerable();
    }

    public IEnumerable<T> GetAs<T>(bool dropErrors)
    {
        return AsEnumerable().TryConvertTo<T>(dropErrors);
    }

    public IEnumerable<object> GetRow(int row)
    {
        for (var col = 0; col < ColumnCount; col++)
            yield return _rangeArray[row, col];
    }

    public IEnumerable<T> GetRowAs<T>(int row, bool dropErrors)
    {
        return GetRow(row).TryConvertTo<T>(dropErrors);
    }

    public IEnumerable<object> GetRowAsType(int row, Type type, bool dropErrors)
    {
        return GetRow(row).TryConvertToType(type, dropErrors);
    }

    public IEnumerable<object> GetColumn(int column)
    {
        for (var row = 0; row < RowCount; row++)
            yield return _rangeArray[row, column];
    }

    public IEnumerable<T> GetColumnAs<T>(int column, bool dropErrors)
    {
        return GetColumn(column).TryConvertTo<T>(dropErrors);
    }

    public IEnumerable<object> GetColumnAsType(int column, Type type, bool dropErrors)
    {
        return GetColumn(column).TryConvertToType(type, dropErrors);
    }

    public object[,] AsArray(RangeOrientation orientation)
    {
        return orientation == RangeOrientation.ByColumn ?
            _rangeArray :
            Shape(ColumnCount, RowCount, ExcelError.ExcelErrorNA)._rangeArray;
    }

    public enum CleanBehavior
    {
        KeepErrors,
        DropErrors,
        ThrowOnErrors
    }

    public IEnumerable<object> AsCleanEnumerable(CleanBehavior cleanBehavior)
    {
        var cleaned = AsEnumerable().FilterMissing();

        if (cleanBehavior == CleanBehavior.DropErrors || cleanBehavior == CleanBehavior.ThrowOnErrors)
            cleaned = cleaned.FilterErrors(cleanBehavior == CleanBehavior.ThrowOnErrors);

        return cleaned;
    }

    public IEnumerable<object> AsCleanEnumerable(string onErrors)
    {
        return AsCleanEnumerable(ParseCleanBehavior(onErrors));
    }

    private static CleanBehavior ParseCleanBehavior(string onErrors)
    {
        return onErrors.Trim().ToLowerInvariant() switch
        {
            "drop" => CleanBehavior.DropErrors,
            "keep" => CleanBehavior.KeepErrors,
            "error" => CleanBehavior.ThrowOnErrors,
            _ => throw new ArgumentException($"{nameof(onErrors)} must be one of 'drop', 'keep' or 'error'"),
        };
    }

    public XlBlockRange Clean(string onErrors)
    {
        return Clean(ParseCleanBehavior(onErrors));
    }

    public XlBlockRange Clean(CleanBehavior cleanBehavior)
    {
        return new XlBlockRange(AsCleanEnumerable(cleanBehavior).AsColumnArray());
    }

    public XlBlockRange MakeSafeForArrayFormulas()
    {
        return Count == 1 ? new XlBlockRange(new[,] { { _rangeArray[0, 0] }, { ExcelError.ExcelErrorNA } }) : this;
    }

    public XlBlockRange Shape(int? rowCount, int? columnCount, object fillWith)
    {
        if (rowCount is null && columnCount is null)
            throw new ArgumentException($"at least one of {nameof(rowCount)} or {nameof(columnCount)} must be provided");

        var rows = 0;
        var columns = 0;
        if (rowCount is not null)
        {
            if (rowCount < 1)
                throw new ArgumentException($"{nameof(rowCount)} cannot be less than 1");
            rows = rowCount.Value;
            if (columnCount is null)
                columns = Count / rows + (Count % rows > 0 ? 1 : 0);
        }

        if (columnCount is not null)
        {
            if (columnCount < 1)
                throw new ArgumentException($"{nameof(columnCount)} cannot be less than 1");
            columns = columnCount.Value;
            if (rowCount is null)
                rows = Count / columns + (Count % columns > 0 ? 1 : 0);
        }

        if ((rows * columns) < Count)
            throw new ArgumentException($"{nameof(rowCount)} * {nameof(columnCount)} is less than size of range");


        var shapedArray = new object[rows, columns];
        Array.Copy(_rangeArray, shapedArray, _rangeArray.Length);
        if (shapedArray.Length > _rangeArray.Length)
        {
            for (var i = 0; i < (shapedArray.Length - _rangeArray.Length); i++)
            {
                var row = (_rangeArray.Length + i) / columns;
                var col = (_rangeArray.Length + i) % columns;
                shapedArray[row, col] = fillWith;
            }
        }
        return new XlBlockRange(shapedArray);
    }

    public XlBlockRange GetUniqueValues()
    {
        var cleaned = AsCleanEnumerable(CleanBehavior.DropErrors);
        var uniqueSet = new HashSet<object>(cleaned).AsColumnArray();
        return new XlBlockRange(uniqueSet);
    }

    public static XlBlockRange Build(object[] objArray)
    {
        // 1d array is a single column
        return new XlBlockRange(objArray.AsColumnArray());
    }

    public static XlBlockRange Build(object[,] objArray)
    {
        // a single range given as an object, as in a parameter conversion lambda or in tests
        return new XlBlockRange(objArray);
    }

    public static XlBlockRange BuildFromMultiple(XlBlockRange[] ranges, string onErrors)
    {
        // when building from multiple ranges flatten them
        var cleanBehavior = ParseCleanBehavior(onErrors);
        var count = ranges.Sum(x => x.Count);
        var objArray = new object[count, 1];
        var index = 0;

        foreach (var range in ranges)
        {
            var cleaned = range.AsCleanEnumerable(cleanBehavior);
            foreach (var item in cleaned)
                objArray[index++, 0] = item;
        }

        if (index < objArray.Length)
        {
            var trimmedArray = new object[index, 1];
            Array.Copy(objArray, trimmedArray, index);
            return new XlBlockRange(trimmedArray);
        }

        return new XlBlockRange(objArray);
    }

    public static XlBlockRange BuildFromString(string str, string delimiter, bool trimStrings = false, bool ignoreEmpty = true)
    {
        var strings = str.Split(new[] { delimiter }, StringSplitOptions.None)
            .Select(x => trimStrings ? x.Trim() : x)
            .Where(x => !ignoreEmpty || !string.IsNullOrEmpty(x))
            .ToArray();

        var objArray = new object[strings.Length, 1];
        for (var i = 0; i < strings.Length; i++)
        {
            objArray[i, 0] = strings[i];
        }
        return new XlBlockRange(objArray);
    }
}
