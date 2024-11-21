namespace XlBlocks.AddIn.Utilities;

using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Data.Analysis;

internal static class DataFrameColumnExtensions
{
    internal static IEnumerable<object> AsEnumerable(this DataFrameColumn column)
    {
        for (var i = 0L; i < column.Length; i++)
            yield return column[i];
    }

    internal static DataFrameColumn ElementwiseIfThenElse(this DataFrameColumn conditionalColumn, DataFrameColumn trueColumn, DataFrameColumn falseColumn)
    {
        if (conditionalColumn is not PrimitiveDataFrameColumn<bool> boolConditional)
            throw new ArgumentException("Conditional must be boolean");

        if ((trueColumn is not NullDataFrameColumn && conditionalColumn.Length != trueColumn.Length) ||
            (falseColumn is not NullDataFrameColumn && conditionalColumn.Length != falseColumn.Length))
            throw new ArgumentException("Input columns must be the same length");

        var result = GetOutputColumn(trueColumn, falseColumn, "ifthen");
        for (var i = 0L; i < conditionalColumn.Length; i++)
        {
            if (boolConditional[i] == true)
            {
                if (trueColumn is NullDataFrameColumn)
                    continue;
                result[i] = result.DataType == typeof(string) ? trueColumn[i]?.ToString() : trueColumn[i];
            }
            else if (boolConditional[i] == false)
            {
                if (falseColumn is NullDataFrameColumn)
                    continue;
                result[i] = result.DataType == typeof(string) ? falseColumn[i]?.ToString() : falseColumn[i];
            }
        }

        return result;
    }

    internal static DataFrameColumn IsDuplicateElement(this DataFrameColumn column)
    {
        if (column is PrimitiveDataFrameColumn<bool> boolColumn)
        {
            return IsDuplicateElement(boolColumn);
        }
        else if (column is PrimitiveDataFrameColumn<double> doubleColumn)
        {
            return IsDuplicateElement(doubleColumn);
        }
        else if (column is PrimitiveDataFrameColumn<float> floatColumn)
        {
            return IsDuplicateElement(floatColumn);
        }
        else if (column is PrimitiveDataFrameColumn<decimal> decimalColumn)
        {
            return IsDuplicateElement(decimalColumn);
        }
        else if (column is StringDataFrameColumn stringColumn)
        {
            return IsDuplicateElement(stringColumn);
        }
        else if (column is PrimitiveDataFrameColumn<int> intColumn)
        {
            return IsDuplicateElement(intColumn);
        }
        else if (column is PrimitiveDataFrameColumn<uint> uintColumn)
        {
            return IsDuplicateElement(uintColumn);
        }
        else if (column is PrimitiveDataFrameColumn<long> longColumn)
        {
            return IsDuplicateElement(longColumn);
        }
        else if (column is PrimitiveDataFrameColumn<ulong> ulongColumn)
        {
            return IsDuplicateElement(ulongColumn);
        }
        else if (column is PrimitiveDataFrameColumn<short> shortColumn)
        {
            return IsDuplicateElement(shortColumn);
        }
        else if (column is PrimitiveDataFrameColumn<ushort> ushortColumn)
        {
            return IsDuplicateElement(ushortColumn);
        }
        else if (column is PrimitiveDataFrameColumn<char> charColumn)
        {
            return IsDuplicateElement(charColumn);
        }
        else if (column is PrimitiveDataFrameColumn<sbyte> sbyteColumn)
        {
            return IsDuplicateElement(sbyteColumn);
        }
        else if (column is PrimitiveDataFrameColumn<byte> byteColumn)
        {
            return IsDuplicateElement(byteColumn);
        }
        else if (column is PrimitiveDataFrameColumn<DateTime> dateTimeColumn)
        {
            return IsDuplicateElement(dateTimeColumn);
        }

        throw new NotSupportedException($"cannot check determine for column of type '{column.DataType.Name}'");
    }

    private static DataFrameColumn IsDuplicateElement<T>(this PrimitiveDataFrameColumn<T> column) where T : unmanaged
    {
        var hashSet = new HashSet<T?>();
        var result = new BooleanDataFrameColumn("duplicated", column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (column[i] is null)
                continue;

            result[i] = !hashSet.Add(column[i]);
        }
        return result;
    }

    private static DataFrameColumn IsDuplicateElement(this StringDataFrameColumn column)
    {
        var hashSet = new HashSet<string>();
        var result = new BooleanDataFrameColumn("duplicated", column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (column[i] is null)
                continue;

            result[i] = !hashSet.Add(column[i]);
        }
        return result;
    }

    private static double? Power(double? x, double? y) => (x is not null && y is not null) ? Math.Pow(x.Value, y.Value) : null;
    internal static DataFrameColumn ElementwiseExponent(this DataFrameColumn baseColumn, DataFrameColumn exponentColumn)
    {
        if (baseColumn.Length != exponentColumn.Length)
            throw new ArgumentException("Input columns must be the same length");

        if (!baseColumn.IsNumericColumn() || !exponentColumn.IsNumericColumn())
            throw new ArgumentException("Inputs must be numeric");

        var doubleColumn = new PseudoDoubleDataFrameColumn(baseColumn);
        var doubleExponentColumn = new PseudoDoubleDataFrameColumn(exponentColumn);
        var result = new DoubleDataFrameColumn("exponent", baseColumn.Length);
        for (var i = 0L; i < baseColumn.Length; i++)
        {
            if (baseColumn[i] is null || exponentColumn[i] is null)
                continue;

            result[i] = Power(doubleColumn[i], doubleExponentColumn[i]);
        }

        return result;
    }

    private static double? Log(double? x) => x is not null ? Math.Log(x.Value) : null;
    private static double? Log(double? x, double? y) => (x is not null && y is not null) ? Math.Pow(x.Value, y.Value) : null;
    internal static DataFrameColumn ElementwiseLog(this DataFrameColumn column)
    {
        if (!column.IsNumericColumn())
            throw new ArgumentException("Input must be numeric");

        var doubleColumn = new PseudoDoubleDataFrameColumn(column);
        var result = new DoubleDataFrameColumn("log", column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (column[i] is null)
                continue;

            result[i] = Log(doubleColumn[i]);
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseLog(this DataFrameColumn column, DataFrameColumn baseColumn)
    {
        if (column.Length != baseColumn.Length)
            throw new ArgumentException("Input columns must be the same length");

        if (!column.IsNumericColumn() || !baseColumn.IsNumericColumn())
            throw new ArgumentException("Inputs must be numeric");

        var doubleColumn = new PseudoDoubleDataFrameColumn(column);
        var doubleBaseColumn = new PseudoDoubleDataFrameColumn(baseColumn);
        var result = new DoubleDataFrameColumn("log", column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (column[i] is null || baseColumn[i] is null)
                continue;

            result[i] = Log(doubleColumn[i], doubleBaseColumn[i]);
        }

        return result;
    }

    #region String functions

    internal static DataFrameColumn ElementwiseConcat(this DataFrameColumn column, DataFrameColumn otherColumn)
    {
        if (column.Length != otherColumn.Length)
            throw new ArgumentException("Columns have mismatched length");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
            result[i] = $"{column[i]}{otherColumn[i]}";

        return result;
    }

    internal static DataFrameColumn ElementwiseLength(this DataFrameColumn column)
    {
        if (column is not StringDataFrameColumn strColumn)
            throw new ArgumentException("Input must be a string");

        var result = new Int32DataFrameColumn($"{column.Name}_length", column.Length);
        for (var i = 0L; i < column.Length; i++)
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
                return match.Value.Last() == '%' ? ".*" : ".";
            });

        return Regex.Replace(interim, @"\\\\([%_])", "$1"); // remove unnecessary escapes
    }

    private static bool? IsWildcardMatch(string? input, string? pattern, bool escapePattern, bool caseInsensitive)
    {
        if (input is null || pattern is null)
            return null;

        // % = match zero or more, _ = match zero or one
        pattern = $"^{(escapePattern ? EscapeLikePattern(pattern) : pattern)}$";
        return Regex.IsMatch(input, pattern, caseInsensitive ? RegexOptions.IgnoreCase : RegexOptions.None);
    }

    internal static DataFrameColumn ElementwiseLike(this DataFrameColumn column, string pattern, bool caseInsensitive = false)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        pattern = EscapeLikePattern(pattern);
        var result = new BooleanDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
            result[i] = IsWildcardMatch(stringColumn[i], pattern, false, caseInsensitive);

        return result;
    }

    internal static DataFrameColumn ElementwiseLike(this DataFrameColumn column, DataFrameColumn patternColumn, bool caseInsensitive = false)
    {
        if (column is not StringDataFrameColumn stringColumn || patternColumn is not StringDataFrameColumn patternStringColumn)
            throw new ArgumentException("Input and pattern must be strings");

        if (column.Length != patternColumn.Length)
            throw new ArgumentException("Columns have mismatched length");

        var result = new BooleanDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
            result[i] = IsWildcardMatch(stringColumn[i], patternStringColumn[i], true, caseInsensitive);

        return result;
    }

    internal static DataFrameColumn ElementwiseSubstring(this DataFrameColumn column, DataFrameColumn startIndexColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        if (!startIndexColumn.IsNumericColumn())
            throw new ArgumentException("Start index must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is null || startIndexColumn[i] is null)
                continue;

            if ((int)startIndexColumn[i] >= stringColumn[i].Length)
                result[i] = string.Empty;
            else
                result[i] = stringColumn[i][(int)startIndexColumn[i]..];
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseSubstring(this DataFrameColumn column, DataFrameColumn startIndexColumn, DataFrameColumn lengthColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        if (!startIndexColumn.IsNumericColumn())
            throw new ArgumentException("Start index must be numeric");

        if (!lengthColumn.IsNumericColumn())
            throw new ArgumentException("Length must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is null || startIndexColumn[i] is null || lengthColumn[i] is null)
                continue;

            if ((int)lengthColumn[i] == 0)
                result[i] = string.Empty;
            else if ((int)startIndexColumn[i] >= stringColumn[i].Length)
                result[i] = string.Empty;
            else
                result[i] = stringColumn[i].Substring((int)startIndexColumn[i], Math.Min((int)lengthColumn[i], stringColumn[i].Length - (int)startIndexColumn[i]));
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseLeft(this DataFrameColumn column, DataFrameColumn lengthColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        if (!lengthColumn.IsNumericColumn())
            throw new ArgumentException("Length must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is null || lengthColumn[i] is null)
                continue;

            if ((int)lengthColumn[i] == 0)
                result[i] = string.Empty;
            else if ((int)lengthColumn[i] >= stringColumn[i].Length)
                result[i] = stringColumn[i];
            else
                result[i] = stringColumn[i][..(int)lengthColumn[i]];
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseRight(this DataFrameColumn column, DataFrameColumn lengthColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        if (!lengthColumn.IsNumericColumn())
            throw new ArgumentException("Length must be numeric");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < stringColumn.Length; i++)
        {
            if (stringColumn[i] is null || lengthColumn[i] is null)
                continue;

            if ((int)lengthColumn[i] == 0)
                result[i] = string.Empty;
            else if ((int)lengthColumn[i] >= stringColumn[i].Length)
                result[i] = stringColumn[i];
            else
                result[i] = stringColumn[i][(stringColumn[i].Length - (int)lengthColumn[i])..];
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseTrim(this DataFrameColumn column)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
            result[i] = stringColumn[i]?.Trim();

        return result;
    }

    internal static DataFrameColumn ElementwiseReplace(this DataFrameColumn column, DataFrameColumn oldValueColumn, DataFrameColumn newValueColumn, DataFrameColumn? caseSensitiveColumn)
    {
        if (column is not StringDataFrameColumn stringColumn)
            throw new ArgumentException("Input must be a string");

        if (oldValueColumn is not StringDataFrameColumn stringOldValueColumn)
            throw new ArgumentException("Old value must be a string");

        if (newValueColumn is not StringDataFrameColumn stringNewValueColumn)
            throw new ArgumentException("New Value must be a string");

        caseSensitiveColumn ??= DataFrameUtilities.CreateConstantDataFrameColumn(true, column.Length);
        if (caseSensitiveColumn is not PrimitiveDataFrameColumn<bool> boolCaseSensitiveColumn)
            throw new ArgumentException("Is case sensitive must be boolean");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < stringColumn.Length; i++)
        {
            if (stringOldValueColumn[i] is null)
                result[i] = stringColumn[i];
            else
                result[i] = stringColumn[i]?.Replace(stringOldValueColumn[i], stringNewValueColumn[i], boolCaseSensitiveColumn[i] == true ? StringComparison.InvariantCultureIgnoreCase : StringComparison.InvariantCulture);
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseRegexTest(this DataFrameColumn column, DataFrameColumn patternColumn, DataFrameColumn? caseSensitiveColumn)
    {
        if (column is not StringDataFrameColumn stringColumn || patternColumn is not StringDataFrameColumn stringPatternColumn)
            throw new ArgumentException("Input and pattern must be strings");

        caseSensitiveColumn ??= DataFrameUtilities.CreateConstantDataFrameColumn(true, column.Length);
        if (caseSensitiveColumn is not PrimitiveDataFrameColumn<bool> boolCaseSensitiveColumn)
            throw new ArgumentException("Is case sensitive must be boolean");

        var result = new BooleanDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (stringColumn[i] is null || stringPatternColumn[i] is null)
                continue;

            result[i] = Regex.IsMatch(stringColumn[i], stringPatternColumn[i], boolCaseSensitiveColumn[i] == false ? RegexOptions.IgnoreCase : RegexOptions.None);
        }

        return result;
    }

    internal static DataFrameColumn ElementwiseRegexReplace(this DataFrameColumn column, DataFrameColumn patternColumn, DataFrameColumn replacementColumn)
    {
        if (column is not StringDataFrameColumn stringColumn || patternColumn is not StringDataFrameColumn stringPatternColumn || replacementColumn is not StringDataFrameColumn stringReplacementColumn)
            throw new ArgumentException("Input, pattern and replacement columns must be strings");

        var result = new StringDataFrameColumn(column.Name, column.Length);
        for (var i = 0L; i < column.Length; i++)
        {
            if (stringColumn[i] is null || stringPatternColumn[i] is null)
                continue;

            result[i] = Regex.Replace(stringColumn[i], stringPatternColumn[i], stringReplacementColumn[i]);
        }

        return result;
    }

    internal static IEnumerable<string?> ToStringEnumerable(this DataFrameColumn column)
    {
        for (var i = 0L; i < column.Length; i++)
            yield return column[i]?.ToString();
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

    #region Type helpers

    private static DataFrameColumn GetOutputColumn(DataFrameColumn left, DataFrameColumn right, string columnName)
    {
        var outputType = DetermineOutputColumnType(left, right);
        return DataFrameUtilities.CreateDataFrameColumn(outputType, columnName, left.Length);
    }

    private static bool IsIntegerType(Type type)
    {
        return type == typeof(ulong) || type == typeof(long) || type == typeof(uint) || type == typeof(int) ||
            type == typeof(ushort) || type == typeof(short) || type == typeof(byte) || type == typeof(sbyte);
    }

    private static Type DetermineOutputColumnType(DataFrameColumn left, DataFrameColumn right)
    {
        if (left.DataType == right.DataType)
            return left.DataType;

        // basically just go for output column that both left and right can be implicitly cast to
        if (left.DataType == typeof(string) || right.DataType == typeof(string))
            return typeof(string);

        if (left.DataType == typeof(bool) || right.DataType == typeof(bool))
            throw new ArgumentException($"Cannot determine output column type for operation between '{left.DataType.Name}' and '{right.DataType.Name}' columns");

        if (left.DataType == typeof(DateTime) || right.DataType == typeof(DateTime))
            throw new ArgumentException($"Cannot determine output column type for operation between '{left.DataType.Name}' and '{right.DataType.Name}' columns");

        if (left.DataType == typeof(decimal) || right.DataType == typeof(decimal))
            throw new ArgumentException($"Cannot determine output column type for operation between '{left.DataType.Name}' and '{right.DataType.Name}' columns");

        if (IsIntegerType(left.DataType) && IsIntegerType(right.DataType))
        {
            if (left.DataType == typeof(ulong) || right.DataType == typeof(ulong) || left.DataType == typeof(long) || right.DataType == typeof(long))
                return typeof(double);

            if (left.DataType == typeof(int) || right.DataType == typeof(int) || left.DataType == typeof(uint) || right.DataType == typeof(uint))
                return typeof(long);

            if (left.DataType == typeof(short) || right.DataType == typeof(short) || left.DataType == typeof(ushort) || right.DataType == typeof(ushort))
                return typeof(int);

            if (left.DataType == typeof(byte) || right.DataType == typeof(byte) || left.DataType == typeof(sbyte) || right.DataType == typeof(sbyte))
                return typeof(short);
        }

        return typeof(double);
    }

    #endregion

    #region DateTime-aware comparisons
    /*
    ElementwiseEquals(right),
    DataFrameExpressionToken.COMP_NOTEQUALS => left.ElementwiseNotEquals(right),
    DataFrameExpressionToken.COMP_LT => left.ElementwiseLessThan(right),
    DataFrameExpressionToken.COMP_GT => left.ElementwiseGreaterThan(right),
    DataFrameExpressionToken.COMP_LTE => left.ElementwiseLessThanOrEqual(right),
    DataFrameExpressionToken.COMP_GTE => left.ElementwiseGreaterThanOrEqual(right),
    */

    internal static DataFrameColumn ElementwiseEqualsDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = leftDateTime[i].Equals(rightDateTime[i]);
            }
            return result;
        }
        return left.ElementwiseEquals(right);
    }

    internal static DataFrameColumn ElementwiseNotEqualsDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = !leftDateTime[i].Equals(rightDateTime[i]);
            }
            return result;
        }
        return left.ElementwiseNotEquals(right);
    }

    internal static DataFrameColumn ElementwiseLessThanDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = leftDateTime[i] < rightDateTime[i];
            }
            return result;
        }
        return left.ElementwiseLessThan(right);
    }

    internal static DataFrameColumn ElementwiseGreaterThanDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = leftDateTime[i] > rightDateTime[i];
            }
            return result;
        }
        return left.ElementwiseGreaterThan(right);
    }

    internal static DataFrameColumn ElementwiseLessThanOrEqualDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = leftDateTime[i] <= rightDateTime[i];
            }
            return result;
        }
        return left.ElementwiseLessThanOrEqual(right);
    }

    internal static DataFrameColumn ElementwiseGreaterThanOrEqualDateAware(this DataFrameColumn left, DataFrameColumn right)
    {
        if (left is PrimitiveDataFrameColumn<DateTime> leftDateTime && right is PrimitiveDataFrameColumn<DateTime> rightDateTime)
        {
            if (left.Length != right.Length)
                throw new ArgumentException("Columns have mismatched length");

            var result = new BooleanDataFrameColumn("equals", left.Length);
            for (var i = 0; i < result.Length; i++)
            {
                if (leftDateTime[i] is null || rightDateTime[i] is null)
                    continue;

                result[i] = leftDateTime[i] >= rightDateTime[i];
            }
            return result;
        }
        return left.ElementwiseGreaterThanOrEqual(right);
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
                if (leftPrim[i] == true && rightPrim[i] == true)
                {
                    result[i] = true;
                }
                else if ((leftPrim[i] == false && rightPrim[i] is not null) ||
                    (leftPrim[i] is not null && rightPrim[i] == false))
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
                if ((leftPrim[i] == true && rightPrim[i] is not null) ||
                    (leftPrim[i] is not null && rightPrim[i] == true))
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
                if ((leftPrim[i] == true && rightPrim[i] == false) ||
                    (leftPrim[i] == false && rightPrim[i] == true))
                {
                    result[i] = true;
                }
                else if ((leftPrim[i] == true && rightPrim[i] == true) ||
                    (leftPrim[i] == false && rightPrim[i] == false))
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
