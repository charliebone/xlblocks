namespace xlBlocks.AddIn.Utilities;

using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;

internal static class DataFrameColumnExtensions
{
    internal static DataFrameColumn ElmentwiseIfThenElse(this DataFrameColumn conditionalColumn, DataFrameColumn trueColumn, DataFrameColumn falseColumn)
    {
        if (conditionalColumn is not PrimitiveDataFrameColumn<bool> boolConditional)
            throw new NotSupportedException("The conditional argument must evaluate to a boolean");

        if (trueColumn.DataType != falseColumn.DataType)
            throw new NotSupportedException($"True and false columns must be the same data type, got '{trueColumn.DataType.Name}' and '{falseColumn.DataType.Name}'");

        var result = trueColumn.Clone();
        for (var i = 0; i < conditionalColumn.Length; i++)
        {
            if (boolConditional[i] == true)
                result[i] = trueColumn[i];
            else if (boolConditional[i] == false)
                result[i] = falseColumn[i];
            else
                result[i] = null;
        }

        return result;
    }

    #region String functions

    internal static DataFrameColumn ElementwiseLength(this DataFrameColumn column)
    {
        if (column is not StringDataFrameColumn strColumn)
            throw new NotSupportedException("ElementwiseLength is only valid for computing length on string columns");

        var result = new Int32DataFrameColumn($"{column.Name}_length", column.Length);
        for (var i = 0; i < column.Length; i++)
            result[i] = strColumn[i]?.Length;

        return result;
    }

    internal static string EscapeLikePattern(string input)
    {
        input = Regex.Escape(input);

        var escapedLikeWildcards = Regex.Matches(input, @"\\\\[%_]"); // escaped LIKE wildcards will have double backlashes from escaping input
        var charsToEscape = @"(?<!\\)[%_]"; // unescaped LIKE wildcards
        var interim = Regex.Replace(input, charsToEscape, match =>
            {
                // Check if the match is part of an escaped sequence
                foreach (Match escapedMatch in escapedLikeWildcards)
                {
                    if (escapedMatch.Index == match.Index - 1)
                        return match.Value.Last().ToString(); // remove escapes becaues these aren't actually special to regex
                }
                return match.Value.Last() == '%' ? ".+" : ".";
            });

        return Regex.Replace(interim, @"\\\\([%_])", "$1"); // remove unnecessary escapes
    }

    private static bool? IsWildcardMatch(string? input, string? pattern, bool escapePattern = true)
    {
        if (input is null || pattern is null)
            return null;

        // % = match zero or more, _ = match zero or one
        pattern = $"^{(escapePattern ? EscapeLikePattern(pattern) : pattern)}$";
        return Regex.IsMatch(input, pattern);
    }

    internal static DataFrameColumn ElementwiseLike(this DataFrameColumn column, string pattern)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new NotSupportedException("ElementwiseLike only valid for comparing string columns or literals");

        pattern = EscapeLikePattern(pattern);
        var result = new BooleanDataFrameColumn(column.Name, column.Length);
        for (var i = 0; i < column.Length; i++)
            result[i] = IsWildcardMatch(stringColumn[i], pattern, false);

        return result;
    }

    internal static DataFrameColumn ElementwiseLike(this DataFrameColumn column, DataFrameColumn patternColumn)
    {
        if (column is not StringDataFrameColumn stringColumn || patternColumn is not StringDataFrameColumn patternStringColumn)
            throw new NotSupportedException("ElementwiseLike only valid for comparing string columns or literals");

        if (column.Length != patternColumn.Length)
            throw new ArgumentException("Columns have mismatched length");

        var result = new BooleanDataFrameColumn(column.Name, column.Length);
        for (var i = 0; i < column.Length; i++)
            result[i] = IsWildcardMatch(stringColumn[i], patternStringColumn[i]);

        return result;
    }

    internal static DataFrameColumn ElementwiseSubstring(this DataFrameColumn column, DataFrameColumn startIndexColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new NotSupportedException("ElementwiseSubstring is only valid on string columns");

        if (!startIndexColumn.IsNumericColumn())
            throw new DataFrameExpressionException("Start index must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is not null && startIndexColumn[i] is not null)
                result[i] = stringColumn[i].Substring((int)startIndexColumn[i]);
            else
                result[i] = null;
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseSubstring(this DataFrameColumn column, DataFrameColumn startIndexColumn, DataFrameColumn lengthColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new NotSupportedException("ElementwiseSubstring is only valid on string columns");

        if (!startIndexColumn.IsNumericColumn())
            throw new DataFrameExpressionException("Start index column must be numeric");

        if (!lengthColumn.IsNumericColumn())
            throw new DataFrameExpressionException("Length column must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is not null && startIndexColumn[i] is not null && lengthColumn[i] is not null)
                result[i] = stringColumn[i].Substring((int)startIndexColumn[i], (int)lengthColumn[i]);
            else
                result[i] = null;
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseTrim(this DataFrameColumn column)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new NotSupportedException("ElementwiseTrim only valid on string columns");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0; i < column.Length; i++)
            result[i] = stringColumn[i]?.Trim();

        return result;
    }

    #endregion

    #region Sorting

    private readonly record struct IndexedColumnItem<T>(long Index, long OriginalIndex, T? Item) where T : struct;
    private readonly record struct IndexedColumnStringItem(long Index, long OriginalIndex, string Item);

    private static IEnumerable<IndexedColumnItem<T>> GetIndexedColumnItems<T>(Int64DataFrameColumn indexColumn,
        PrimitiveDataFrameColumn<T> column) where T : unmanaged
    {
        for (var i = 0L; i < column.Length; i++)
        {
            Debug.Assert(indexColumn[i] is not null);
            var origIndex = indexColumn[i] ?? i;
            yield return new IndexedColumnItem<T>(i, origIndex, column[origIndex]);
        }
    }

    private static IEnumerable<IndexedColumnStringItem> GetIndexedColumnItems(Int64DataFrameColumn indexColumn,
        StringDataFrameColumn column)
    {
        for (var i = 0L; i < column.Length; i++)
        {
            Debug.Assert(indexColumn[i] is not null);
            var origIndex = indexColumn[i] ?? i;
            yield return new IndexedColumnStringItem(i, origIndex, column[origIndex]);
        }
    }

    private static Int64DataFrameColumn GetIndexColumn(DataFrameColumn column)
    {
        var indexColumn = new Int64DataFrameColumn("indices", column.Length);
        for (var i = 0L; i < column.Length; i++)
            indexColumn[i] = i;

        return indexColumn;
    }

    internal static Int64DataFrameColumn SortColumn(this DataFrameColumn column, Int64DataFrameColumn indexColumn, bool descending, bool nullsFirst)
    {
        indexColumn ??= GetIndexColumn(column);
        if (indexColumn.Length != column.Length)
            throw new ArgumentException("index column and sort column must be same length");

        if (column is PrimitiveDataFrameColumn<bool> boolColumn)
        {
            return SortColumn(indexColumn, boolColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<double> doubleColumn)
        {
            return SortColumn(indexColumn, doubleColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<float> floatColumn)
        {
            return SortColumn(indexColumn, floatColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<decimal> decimalColumn)
        {
            return SortColumn(indexColumn, decimalColumn, descending, nullsFirst);
        }
        else if (column is StringDataFrameColumn stringColumn)
        {
            return SortColumn(indexColumn, stringColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<int> intColumn)
        {
            return SortColumn(indexColumn, intColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<uint> uintColumn)
        {
            return SortColumn(indexColumn, uintColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<long> longColumn)
        {
            return SortColumn(indexColumn, longColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<ulong> ulongColumn)
        {
            return SortColumn(indexColumn, ulongColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<short> shortColumn)
        {
            return SortColumn(indexColumn, shortColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<ushort> ushortColumn)
        {
            return SortColumn(indexColumn, ushortColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<char> charColumn)
        {
            return SortColumn(indexColumn, charColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<sbyte> sbyteColumn)
        {
            return SortColumn(indexColumn, sbyteColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<byte> byteColumn)
        {
            return SortColumn(indexColumn, byteColumn, descending, nullsFirst);
        }
        else if (column is PrimitiveDataFrameColumn<DateTime> dateTimeColumn)
        {
            return SortColumn(indexColumn, dateTimeColumn, descending, nullsFirst);
        }

        throw new NotSupportedException($"cannot sort column of type '{column.DataType.Name}'");
    }

    private static Int64DataFrameColumn SortColumn<T>(Int64DataFrameColumn indexColumn, PrimitiveDataFrameColumn<T> column, bool descending, bool nullsFirst) where T : unmanaged
    {
        var indexedItems = GetIndexedColumnItems(indexColumn, column);

        // to sort nulls consistently at bottom or top regardless of sort order, handle descending in comparer itself
        return new Int64DataFrameColumn("indices", indexedItems.OrderBy(x => x, new ComparerWithIndex<T>(descending, nullsFirst)).Select(x => x.OriginalIndex));
    }

    private static Int64DataFrameColumn SortColumn(Int64DataFrameColumn indexColumn, StringDataFrameColumn column, bool descending, bool nullsFirst)
    {
        var indexedItems = GetIndexedColumnItems(indexColumn, column);

        // to sort nulls consistently at bottom or top regardless of sort order, handle descending in comparer itself
        return new Int64DataFrameColumn("indices", indexedItems.OrderBy(x => x, new StringComparerWithIndex(descending, nullsFirst)).Select(x => x.OriginalIndex));
    }

    private sealed class ComparerWithIndex<T> : IComparer<IndexedColumnItem<T>> where T : struct
    {
        private readonly Comparer<T> _comparer;
        private readonly bool _descending;
        private readonly bool _nullsFirst;

        public ComparerWithIndex(bool descending, bool nullsFirst)
        {
            _comparer = Comparer<T>.Default;
            _descending = descending;
            _nullsFirst = nullsFirst;
        }

        public int Compare(IndexedColumnItem<T> x, IndexedColumnItem<T> y)
        {
            if (x.Item is not null && y.Item is not null)
            {
                var result = _comparer.Compare(x.Item.Value, y.Item.Value);
                return result != 0 ? result * (_descending ? -1 : 1) : x.Index.CompareTo(y.Index);
            }

            // null sorting ignores descending argument
            if (x.Item is null & y.Item is null)
                return x.Index.CompareTo(y.Index);
            else if (x.Item is null)
                return (_nullsFirst ? -1 : 1);
            else
                return (_nullsFirst ? 1 : -1);
        }
    }

    private sealed class StringComparerWithIndex : IComparer<IndexedColumnStringItem>
    {
        private readonly Comparer<string> _comparer;
        private readonly bool _descending;
        private readonly bool _nullsFirst;

        public StringComparerWithIndex(bool descending, bool nullsFirst)
        {
            _comparer = Comparer<string>.Default;
            _descending = descending;
            _nullsFirst = nullsFirst;
        }

        public int Compare(IndexedColumnStringItem x, IndexedColumnStringItem y)
        {
            if (x.Item is not null && y.Item is not null)
            {
                var result = _comparer.Compare(x.Item, y.Item);
                return (result != 0 ? result * (_descending ? -1 : 1) : x.Index.CompareTo(y.Index));
            }

            // null sorting ignores descending argument
            if (x.Item is null & y.Item is null)
                return x.Index.CompareTo(y.Index);
            else if (x.Item is null)
                return (_nullsFirst ? -1 : 1);
            else
                return (_nullsFirst ? 1 : -1);
        }
    }

    #endregion

    #region Safe logical comparisons

    /*
     * these methods are to circumvent bug in boxed DataFrameColumn boolean comparisons: https://github.com/dotnet/machinelearning/issues/7091
     * it should be fixed in v0.21.2 of Microsoft.Data.Analysis
     */

    internal static DataFrameColumn AndSafe(this DataFrameColumn left, DataFrameColumn right, bool inPlace = false)
    {
        if (left is BooleanDataFrameColumn leftBool && right is BooleanDataFrameColumn rightBool)
        {
            return leftBool.And(rightBool, inPlace);
        }
        else if (left is PrimitiveDataFrameColumn<bool> leftPrim && right is PrimitiveDataFrameColumn<bool> rightPrim)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = inPlace ? left : left.Clone();
            for (var i = 0; i < result.Length; i++)
            {
                if (leftPrim[i] == true && rightPrim[i] == true ||
                    leftPrim[i] is not null && rightPrim[i] == true)
                {
                    result[i] = true;
                }
                else if (leftPrim[i] == false && rightPrim[i] is not null ||
                    leftPrim[i] is not null && rightPrim[i] == false)
                {
                    result[i] = false;
                }
                else
                {
                    result[i] = null;
                }
            }
            return result;
        }
        return left.And(right, inPlace);
    }

    internal static DataFrameColumn OrSafe(this DataFrameColumn left, DataFrameColumn right, bool inPlace = false)
    {
        if (left is BooleanDataFrameColumn leftBool && right is BooleanDataFrameColumn rightBool)
        {
            return leftBool.Or(rightBool, inPlace);
        }
        else if (left is PrimitiveDataFrameColumn<bool> leftPrim && right is PrimitiveDataFrameColumn<bool> rightPrim)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = inPlace ? left : left.Clone();
            for (var i = 0; i < result.Length; i++)
            {
                if (leftPrim[i] == true && rightPrim[i] is not null ||
                    leftPrim[i] is not null && rightPrim[i] == true)
                {
                    result[i] = true;
                }
                else if (leftPrim[i] == false && rightPrim[i] == false)
                {
                    result[i] = false;
                }
                else
                {
                    result[i] = null;
                }
            }
            return result;
        }
        return left.Or(right, inPlace);
    }

    internal static DataFrameColumn XorSafe(this DataFrameColumn left, DataFrameColumn right, bool inPlace = false)
    {
        if (left is BooleanDataFrameColumn leftBool && right is BooleanDataFrameColumn rightBool)
        {
            return leftBool.Xor(rightBool, inPlace);
        }
        else if (left is PrimitiveDataFrameColumn<bool> leftPrim && right is PrimitiveDataFrameColumn<bool> rightPrim)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = inPlace ? left : left.Clone();
            for (var i = 0; i < result.Length; i++)
            {
                if (leftPrim[i] == true && rightPrim[i] == false ||
                    leftPrim[i] == false && rightPrim[i] == true)
                {
                    result[i] = true;
                }
                else if (leftPrim[i] == true && rightPrim[i] == true ||
                    leftPrim[i] == false && rightPrim[i] == false)
                {
                    result[i] = false;
                }
                else
                {
                    result[i] = null;
                }
            }
            return result;
        }
        return left.Xor(right, inPlace);
    }

    #endregion
}
