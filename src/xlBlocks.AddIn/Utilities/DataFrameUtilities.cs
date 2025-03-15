namespace XlBlocks.AddIn.Utilities;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataFrameUtilities
{
    internal static DataFrame ToDataFrame(object[,] objArray, Type[] types)
    {
        if (objArray.GetLength(1) != types.Length)
            throw new ArgumentException($"{nameof(types)} array must have length equal to number of columns in {nameof(objArray)}");

        var columnInfos = Enumerable.Range(0, objArray.GetLength(1))
            .Select(col => (objArray[0, col].ToString(), types[col]))
            .ToList();

        var rowLists = Enumerable.Range(1, objArray.GetLength(0) - 1)
            .Select(row => Enumerable.Range(0, objArray.GetLength(1))
                .Select(col => objArray[row, col]).ToList())
            .ToList();

        return DataFrame.LoadFrom(rowLists, columnInfos);
    }

    private static IEnumerable<TResult> RepeatLong<TResult>(TResult element, long count)
    {
        for (var i = 0L; i < count; i++) yield return element;
    }

    public static DataFrameColumn CreateConstantDataFrameColumn(object obj, string type, long length, string columnName = "constant")
    {
        var colType = ParamTypeConverter.StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
        return CreateConstantDataFrameColumn(obj, colType, length, columnName);
    }

    public static DataFrameColumn CreateConstantDataFrameColumn(object obj, Type type, long length, string columnName = "constant")
    {
        return CreateDataFrameColumn(RepeatLong(obj, length), type, columnName);
    }

    public static DataFrameColumn CreateConstantDataFrameColumn<T>(T obj, long length, string columnName = "constant")
    {
        return CreateDataFrameColumn(RepeatLong(obj, length), columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn<T>(IEnumerable<T?> typedItems, string columnName = "constant")
    {
        return CreateDataFrameColumn(typedItems, typeof(T), columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn(object input, long length, string? type = null, string columnName = "constant")
    {
        return CreateDataFrameColumn(RepeatLong(input, length), type, columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn(IEnumerable<object> input, string? type = null, string columnName = "constant")
    {
        if (type != null)
        {
            var columnType = ParamTypeConverter.StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
            var convertedValues = input.ConvertToProvidedType(type)
                .Select(x =>
                {
                    if (x.Input is null)
                        return null;

                    if (!x.IsMissingOrError && x.Success)
                        return x.ConvertedInput;

                    throw new ArgumentException($"Item '{x.Input}' cannot be converted to type '{columnType.Name}'");
                });

            return CreateDataFrameColumn(convertedValues, columnType, columnName);
        }
        else
        {
            var guessConversions = input.ConvertToBestTypes();
            var determinedType = guessConversions.Where(x => !x.IsMissingOrError)
                .Select(x => x.ConvertedType)
                .DetermineBestType();

            var convertedValues = guessConversions.Select(x =>
            {
                if (x.IsMissingOrError || x.Input is null)
                    return null;

                if (x.Success && x.ConvertedType == determinedType)
                    return x.ConvertedInput;

                if (ParamTypeConverter.TryConvertToProvidedType(x.Input, determinedType, out var convertedInput))
                    return convertedInput;

                throw new ArgumentException($"Item '{x.Input}' cannot be converted to type '{determinedType.Name}'");
            });

            return CreateDataFrameColumn(convertedValues, determinedType, columnName);
        }
    }

    public static DataFrameColumn CreateDataFrameColumn(IEnumerable typedItems, Type type, string columnName = "constant")
    {
        if (type == typeof(bool))
        {
            return new BooleanDataFrameColumn(columnName, typedItems.Cast<bool?>());
        }
        else if (type == typeof(double))
        {
            return new DoubleDataFrameColumn(columnName, typedItems.Cast<double?>());
        }
        else if (type == typeof(float))
        {
            return new SingleDataFrameColumn(columnName, typedItems.Cast<float?>());
        }
        else if (type == typeof(decimal))
        {
            return new DecimalDataFrameColumn(columnName, typedItems.Cast<decimal?>());
        }
        else if (type == typeof(string))
        {
            return new StringDataFrameColumn(columnName, typedItems.Cast<string?>());
        }
        else if (type == typeof(int))
        {
            return new Int32DataFrameColumn(columnName, typedItems.Cast<int?>());
        }
        else if (type == typeof(uint))
        {
            return new UInt32DataFrameColumn(columnName, typedItems.Cast<uint?>());
        }
        else if (type == typeof(long))
        {
            return new Int64DataFrameColumn(columnName, typedItems.Cast<long?>());
        }
        else if (type == typeof(ulong))
        {
            return new UInt64DataFrameColumn(columnName, typedItems.Cast<ulong?>());
        }
        else if (type == typeof(short))
        {
            return new Int16DataFrameColumn(columnName, typedItems.Cast<short?>());
        }
        else if (type == typeof(ushort))
        {
            return new UInt16DataFrameColumn(columnName, typedItems.Cast<ushort?>());
        }
        else if (type == typeof(char))
        {
            return new CharDataFrameColumn(columnName, typedItems.Cast<char?>());
        }
        else if (type == typeof(sbyte))
        {
            return new SByteDataFrameColumn(columnName, typedItems.Cast<sbyte?>());
        }
        else if (type == typeof(byte))
        {
            return new ByteDataFrameColumn(columnName, typedItems.Cast<byte?>());
        }
        else if (type == typeof(DateTime))
        {
            return new DateTimeDataFrameColumn(columnName, typedItems.Cast<DateTime?>());
        }

        throw new NotSupportedException($"cannot build data frame column of type '{type.Name}'");
    }

    public static DataFrameColumn CreateDataFrameColumn(Type type, string columnName, long length = 0)
    {
        if (type == typeof(bool))
        {
            return new BooleanDataFrameColumn(columnName, length);
        }
        else if (type == typeof(double))
        {
            return new DoubleDataFrameColumn(columnName, length);
        }
        else if (type == typeof(float))
        {
            return new SingleDataFrameColumn(columnName, length);
        }
        else if (type == typeof(decimal))
        {
            return new DecimalDataFrameColumn(columnName, length);
        }
        else if (type == typeof(string))
        {
            return new StringDataFrameColumn(columnName, length);
        }
        else if (type == typeof(int))
        {
            return new Int32DataFrameColumn(columnName, length);
        }
        else if (type == typeof(uint))
        {
            return new UInt32DataFrameColumn(columnName, length);
        }
        else if (type == typeof(long))
        {
            return new Int64DataFrameColumn(columnName, length);
        }
        else if (type == typeof(ulong))
        {
            return new UInt64DataFrameColumn(columnName, length);
        }
        else if (type == typeof(short))
        {
            return new Int16DataFrameColumn(columnName, length);
        }
        else if (type == typeof(ushort))
        {
            return new UInt16DataFrameColumn(columnName, length);
        }
        else if (type == typeof(char))
        {
            return new CharDataFrameColumn(columnName, length);
        }
        else if (type == typeof(sbyte))
        {
            return new SByteDataFrameColumn(columnName, length);
        }
        else if (type == typeof(byte))
        {
            return new ByteDataFrameColumn(columnName, length);
        }
        else if (type == typeof(DateTime))
        {
            return new DateTimeDataFrameColumn(columnName, length);
        }

        throw new NotSupportedException($"cannot build data frame column of type '{type.Name}'");
    }

    public static void SetTypedValue(this DataFrame dataFrame, string columnName, long index, object? value)
    {
        var column = dataFrame[columnName];
        if (value is null)
        {
            column[index] = null;
        }
        else if (column.DataType == typeof(bool) && value is bool boolValue)
        {
            column[index] = boolValue;
        }
        else if (column.DataType == typeof(double) && value is double doubleValue)
        {
            column[index] = doubleValue;
        }
        else if (column.DataType == typeof(float) && value is float floatValue)
        {
            column[index] = floatValue;
        }
        else if (column.DataType == typeof(decimal) && value is decimal decimalValue)
        {
            column[index] = decimalValue;
        }
        else if (column.DataType == typeof(string) && value is string stringValue)
        {
            column[index] = stringValue;
        }
        else if (column.DataType == typeof(int) && value is int intValue)
        {
            column[index] = intValue;
        }
        else if (column.DataType == typeof(uint) && value is uint uintValue)
        {
            column[index] = uintValue;
        }
        else if (column.DataType == typeof(long) && value is long longValue)
        {
            column[index] = longValue;
        }
        else if (column.DataType == typeof(ulong) && value is ulong ulongValue)
        {
            column[index] = ulongValue;
        }
        else if (column.DataType == typeof(short) && value is short shortValue)
        {
            column[index] = shortValue;
        }
        else if (column.DataType == typeof(ushort) && value is ushort ushortValue)
        {
            column[index] = ushortValue;
        }
        else if (column.DataType == typeof(char) && value is char charValue)
        {
            column[index] = charValue;
        }
        else if (column.DataType == typeof(sbyte) && value is sbyte sbyteValue)
        {
            column[index] = sbyteValue;
        }
        else if (column.DataType == typeof(byte) && value is byte byteValue)
        {
            column[index] = byteValue;
        }
        else if (column.DataType == typeof(DateTime) && value is DateTime dateTimeValue)
        {
            column[index] = dateTimeValue;
        }
        else
        {
            throw new NotSupportedException($"value type '{value.GetType().Name}' does not match column type '{column.DataType.Name}'");
        }
    }

    internal static DataFrame TrimNullRows(this DataFrame dataFrame)
    {
        if (dataFrame.Rows.Count == 0)
            return dataFrame;

        var nonNull = false;
        var index = new BooleanDataFrameColumn("index", RepeatLong(true, dataFrame.Rows.Count));
        for (var row = dataFrame.Rows.Count - 1; row >= 0; row--)
        {
            foreach (var column in dataFrame.Columns)
            {
                if (column[row] != null)
                {
                    nonNull = true;
                    break;
                }
            }
            if (nonNull)
                break;

            index[row] = false;
        }

        if (index[dataFrame.Rows.Count - 1] != true)
            dataFrame = dataFrame[index];

        return dataFrame;
    }

    public static bool IsValidColumnType(Type type)
    {
        return type == typeof(bool) ||
            type == typeof(double) ||
            type == typeof(float) ||
            type == typeof(decimal) ||
            type == typeof(string) ||
            type == typeof(int) ||
            type == typeof(uint) ||
            type == typeof(long) ||
            type == typeof(ulong) ||
            type == typeof(short) ||
            type == typeof(ushort) ||
            type == typeof(char) ||
            type == typeof(sbyte) ||
            type == typeof(byte) ||
            type == typeof(DateTime);
    }

    public static JoinAlgorithm ParseJoinType(string joinType)
    {
        return joinType.ToLowerInvariant() switch
        {
            "inner" => JoinAlgorithm.Inner,
            "full" or "outer" or "full outer" => JoinAlgorithm.FullOuter,
            "left" => JoinAlgorithm.Left,
            "right" => JoinAlgorithm.Right,
            _ => throw new ArgumentException($"unknown join type '{joinType}', must be one of 'inner', 'outer' 'right' or 'left'")
        };
    }

    public static DropNullOptions ParseDropNullOption(string dropNullBehavior)
    {
        return dropNullBehavior.ToLowerInvariant() switch
        {
            "all" => DropNullOptions.All,
            "any" => DropNullOptions.Any,
            _ => throw new ArgumentException($"unknown drop null behavior '{dropNullBehavior}', must be either 'all' or 'any'")
        };
    }

    public enum DuplicateKeyBehavior
    {
        Error,
        TakeFirst,
        TakeLast
    }

    public static DuplicateKeyBehavior ParseDuplicateKeyBehavior(string onDuplicateKeys)
    {
        return onDuplicateKeys.ToLowerInvariant() switch
        {
            "error" => DuplicateKeyBehavior.Error,
            "first" => DuplicateKeyBehavior.TakeFirst,
            "last" => DuplicateKeyBehavior.TakeLast,
            _ => throw new ArgumentException($"unknown duplicate key behavior '{onDuplicateKeys}', must be one of 'error', 'first' or 'last'")
        };
    }

    public static DataFrame ComputeGroupAggregations(DataFrame dataFrame, List<string> groupColumnNames, List<string> groupByOperations, List<string> aggregationColumnNames, List<string> newColumnNames)
    {
        if (groupByOperations.Count != aggregationColumnNames.Count || groupByOperations.Count != newColumnNames.Count)
            throw new ArgumentException("group by operations, aggregation columns and new column names must have same lengths");

        var dataFrameWithComposite = dataFrame.Clone();
        var compositeKeyCol = dataFrame.MakeCompositeColumn(groupColumnNames);
        dataFrameWithComposite.Columns.Add(compositeKeyCol);

        var newColumns = new List<DataFrameColumn>();
        foreach (var columnName in groupColumnNames)
            newColumns.Add(dataFrameWithComposite[columnName]);

        var groupBy = dataFrameWithComposite.GroupBy(compositeKeyCol.Name);
        var groupBySets = groupByOperations.Zip(aggregationColumnNames, newColumnNames).ToLookup(x => x.First, x => new { AggColumn = x.Second, OutputName = x.Third });
        foreach (var groupBySet in groupBySets)
        {
            var groupedDataFrame = GroupByEnhanced(dataFrameWithComposite, groupBySet.Key, groupBySet.Select(x => x.AggColumn).Distinct().ToArray(), groupBy);
            groupedDataFrame.Columns[0].SetName(compositeKeyCol.Name);

            var joined = dataFrameWithComposite.Merge(groupedDataFrame, new[] { compositeKeyCol.Name }, new[] { compositeKeyCol.Name }, "_left", "_right", JoinAlgorithm.Left);
            var joinedColumnNames = joined.Columns.Select(x => x.Name).ToList();

            foreach (var columnSet in groupBySet)
            {
                var aggregationColumn = joined[$"{columnSet.AggColumn}_right"].Clone();
                aggregationColumn.SetName(newColumns.GetSafeColumnName(columnSet.OutputName));
                newColumns.Add(aggregationColumn);
            }
        }

        var newDataFrame = new DataFrame(newColumns);
        newDataFrame = newDataFrame.Filter(compositeKeyCol.IsDuplicateElement().ElementwiseEquals(false));
        return newDataFrame;
    }

    private static DataFrame GroupByEnhanced(this DataFrame dataFrame, string groupByOperation, string[] columnNames, GroupBy groupBy)
    {
        return groupByOperation.ToLower() switch
        {
            "sum" => groupBy.ToGroupByEnhanced(dataFrame).Sum(columnNames),
            "product" or "prod" => groupBy.ToGroupByEnhanced(dataFrame).Product(columnNames),
            "min" or "minimum" => groupBy.ToGroupByEnhanced(dataFrame).Min(columnNames),
            "max" or "maximum" => groupBy.ToGroupByEnhanced(dataFrame).Max(columnNames),
            "mean" or "average" or "avg" => groupBy.ToGroupByEnhanced(dataFrame).Mean(columnNames),
            "med" or "median" => groupBy.ToGroupByEnhanced(dataFrame).Median(columnNames),
            "count" => groupBy.ToGroupByEnhanced(dataFrame).Count(columnNames, false),
            "counta" => groupBy.ToGroupByEnhanced(dataFrame).Count(columnNames, true),
            "first" => groupBy.ToGroupByEnhanced(dataFrame).First(columnNames, false),
            "firsta" => groupBy.ToGroupByEnhanced(dataFrame).First(columnNames, true),
            "last" => groupBy.ToGroupByEnhanced(dataFrame).Last(columnNames, false),
            "lasta" => groupBy.ToGroupByEnhanced(dataFrame).Last(columnNames, true),
            "all" => groupBy.ToGroupByEnhanced(dataFrame).All(columnNames, false),
            "alla" => groupBy.ToGroupByEnhanced(dataFrame).All(columnNames, true),
            "stddev" => groupBy.ToGroupByEnhanced(dataFrame).StdDev(columnNames, true),
            "stddevp" => groupBy.ToGroupByEnhanced(dataFrame).StdDev(columnNames, false),
            "var" or "variance" => groupBy.ToGroupByEnhanced(dataFrame).Variance(columnNames, true),
            "varp" or "variancep" => groupBy.ToGroupByEnhanced(dataFrame).Variance(columnNames, false),
            "skew" => groupBy.ToGroupByEnhanced(dataFrame).Skew(columnNames, true),
            "skewp" => groupBy.ToGroupByEnhanced(dataFrame).Skew(columnNames, false),
            "kurt" or "kurtosis" => groupBy.ToGroupByEnhanced(dataFrame).Kurtosis(columnNames, true),
            "kurtp" or "kurtosisp" => groupBy.ToGroupByEnhanced(dataFrame).Kurtosis(columnNames, false),
            _ => throw new ArgumentException($"unknown group by operation'{groupByOperation}', must be one of 'sum', 'product', 'min' 'max' 'mean', 'median', 'count', 'counta', 'first', 'firsta', 'last', 'lasta', 'stddev', 'stddevp', 'var', 'varp', 'skew', 'skewp', 'kurt' or 'kurp'")
        };
    }

    private static IGroupByEnhanced ToGroupByEnhanced(this GroupBy groupBy, DataFrame dataFrame)
    {
        if (groupBy is GroupBy<bool> boolGroupBy)
        {
            return new GroupByEnhanced<bool>(boolGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<double> doubleGroupBy)
        {
            return new GroupByEnhanced<double>(doubleGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<float> floatGroupBy)
        {
            return new GroupByEnhanced<float>(floatGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<decimal> decimalGroupBy)
        {
            return new GroupByEnhanced<decimal>(decimalGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<string> stringGroupBy)
        {
            return new GroupByEnhanced<string>(stringGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<int> intGroupBy)
        {
            return new GroupByEnhanced<int>(intGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<uint> uintGroupBy)
        {
            return new GroupByEnhanced<uint>(uintGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<long> longGroupBy)
        {
            return new GroupByEnhanced<long>(longGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<ulong> ulongGroupBy)
        {
            return new GroupByEnhanced<ulong>(ulongGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<short> shortGroupBy)
        {
            return new GroupByEnhanced<short>(shortGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<ushort> ushortGroupBy)
        {
            return new GroupByEnhanced<ushort>(ushortGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<char> charGroupBy)
        {
            return new GroupByEnhanced<char>(charGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<sbyte> sbyteGroupBy)
        {
            return new GroupByEnhanced<sbyte>(sbyteGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<byte> byteGroupBy)
        {
            return new GroupByEnhanced<byte>(byteGroupBy, dataFrame);
        }
        else if (groupBy is GroupBy<DateTime> dateTimeGroupBy)
        {
            return new GroupByEnhanced<DateTime>(dateTimeGroupBy, dataFrame);
        }

        throw new NotSupportedException($"cannot group by column of type");
    }

    public static DataFrame DictionaryToDataFrame(XlBlockDictionary dictionary, string keyColumnName, string valueColumnName, string? valueType = null)
    {
        var keyColumn = CreateDataFrameColumn(dictionary.Keys, dictionary.KeyType, keyColumnName);
        var valueColumn = CreateDataFrameColumn(dictionary.Values, valueType, valueColumnName);

        return new DataFrame(keyColumn, valueColumn);
    }

    private static string GetSafeColumnName(this DataFrame dataFrame, string baseName)
    {
        var columnName = baseName;
        var tail = 0;
        while (dataFrame.Columns.IndexOf(columnName) >= 0)
            columnName = $"{baseName}.{tail++}";
        return columnName;
    }

    private static string GetSafeColumnName(this IList<DataFrameColumn> columns, string baseName)
    {
        var columnName = baseName;
        var tail = 0;
        var columnNames = columns.Select(x => x.Name).ToHashSet();
        while (columnNames.Contains(columnName))
            columnName = $"{baseName}.{tail++}";
        return columnName;
    }

    public static DataFrameColumn MakeCompositeColumn(this DataFrame dataFrame, List<string>? includeColumns = null)
    {
        var compositeIndex = dataFrame.Columns
            .Where(x => includeColumns is null || includeColumns.Contains(x.Name))
            .Select(x => x.ToStringEnumerable())
            .Aggregate((x, y) => x.Zip(y).Select(z => $"{z.First ?? "null"}_{z.Second ?? "null"}"));

        return new StringDataFrameColumn(dataFrame.GetSafeColumnName("composite"), compositeIndex);
    }

    public static DataFrame Merge(DataFrame left, DataFrame right, string[] leftJoinColumns, string[] rightJoinColumns, string? leftSuffix, string? rightSuffix, JoinAlgorithm joinAlgorithm)
    {
        if (joinAlgorithm == JoinAlgorithm.Left || joinAlgorithm == JoinAlgorithm.Right)
        {
            // ml.net left and right joins drop rows where the value being joined on is null, which is different than standard sql behavior
            var retained = joinAlgorithm == JoinAlgorithm.Left ? left.Clone() : right.Clone();
            var retainedJoinColumns = joinAlgorithm == JoinAlgorithm.Left ? leftJoinColumns : rightJoinColumns;
            var retainedSuffix = joinAlgorithm == JoinAlgorithm.Left ? leftSuffix : rightSuffix;

            var other = joinAlgorithm == JoinAlgorithm.Left ? right.Clone() : left.Clone();
            var otherJoinColumns = joinAlgorithm == JoinAlgorithm.Left ? rightJoinColumns : leftJoinColumns;
            var otherSuffix = joinAlgorithm == JoinAlgorithm.Left ? rightSuffix : leftSuffix;

            var retainedIndexColumn = GetIndexColumn(retained);
            retained.Columns.Add(retainedIndexColumn);
            var retainedIndexColumnName = retainedIndexColumn.Name;

            var retainedNullIndex = new BooleanDataFrameColumn("nullIndex", RepeatLong(false, retained.Rows.Count));
            for (var i = 0L; i < retainedJoinColumns.Length; i++)
            {
                retainedNullIndex = PopulateNullIndices(retained[retainedJoinColumns[i]], retainedNullIndex);
            }
            var retainedNullRows = retained.Filter(retainedNullIndex);
            var retainedToJoin = retained.Filter(retainedNullIndex.ElementwiseEquals(false));
            DataFrame merged;
            if (joinAlgorithm == JoinAlgorithm.Left)
            {
                merged = retainedToJoin.Merge(other, retainedJoinColumns, otherJoinColumns, retainedSuffix, otherSuffix, JoinAlgorithm.Left);
                for (var i = 0; i < retainedNullRows.Columns.Count; i++)
                    retainedNullRows.Columns[i].SetName(merged.Columns[i].Name);
            }
            else
            {
                merged = other.Merge(retainedToJoin, otherJoinColumns, retainedJoinColumns, otherSuffix, retainedSuffix, JoinAlgorithm.Right);
                for (var i = 0; i < retainedNullRows.Columns.Count; i++)
                    retainedNullRows.Columns[i].SetName(merged.Columns[i + other.Columns.Count].Name);
            }
            merged.Append(retainedNullRows.Rows, true);

            var sortedIndex = merged[retainedIndexColumnName].GetSortIndices();
            var newColumns = new List<DataFrameColumn>(retained.Columns.Count - 1);
            for (var i = 0; i < merged.Columns.Count; i++)
            {
                var oldColumn = merged.Columns[i];
                if (oldColumn.Name == retainedIndexColumnName)
                    continue;

                var newColumn = oldColumn.Clone(sortedIndex, false, 0);
                newColumns.Add(newColumn);
            }

            return new DataFrame(newColumns);
        }
        return left.Merge(right, leftJoinColumns, rightJoinColumns, leftSuffix, rightSuffix, joinAlgorithm);
    }

    private static BooleanDataFrameColumn PopulateNullIndices(DataFrameColumn column, BooleanDataFrameColumn nullIndexColumn)
    {
        for (var i = 0L; i < column.Length; i++)
            if (column[i] is null)
                nullIndexColumn[i] = true;

        return nullIndexColumn;
    }

    internal static Int64DataFrameColumn GetIndexColumn(DataFrame dataFrame)
    {
        if (dataFrame is null || dataFrame.Columns.Count == 0)
            throw new ArgumentException("Dataframe must have at least one column");

        var column = dataFrame.Columns[0].GetIndexColumn();
        column.SetName(dataFrame.GetSafeColumnName("rowIndex"));
        return column;
    }
}
