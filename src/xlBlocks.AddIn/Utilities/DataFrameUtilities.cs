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
                if (x.IsMissingOrError)
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

    public delegate DataFrame GroupByOperation(GroupBy groupBy, string[] columnNames);
    public static GroupByOperation ParseGroupByOperation(string groupByOperation)
    {
        return groupByOperation.ToLowerInvariant() switch
        {
            "sum" => (GroupBy g, string[] n) => g.Sum(n),
            "product" or "prod" => (GroupBy g, string[] n) => g.Product(n),
            "min" or "minimum" => (GroupBy g, string[] n) => g.Min(n),
            "max" or "maximum" => (GroupBy g, string[] n) => g.Max(n),
            "mean" or "average" or "avg" => (GroupBy g, string[] n) => g.Mean(n),
            "count" => (GroupBy g, string[] n) => g.Count(n),
            _ => throw new ArgumentException($"unknown group by operation'{groupByOperation}', must be one of 'sum', 'product', 'min' 'max' 'mean' or 'count'")
        };
    }

    public static DataFrame DictionaryToDataFrame(XlBlockDictionary dictionary, string keyColumnName, string valueColumnName, string? valueType = null)
    {
        var keyColumn = CreateDataFrameColumn(dictionary.Keys, dictionary.KeyType, keyColumnName);
        var valueColumn = CreateDataFrameColumn(dictionary.Values, valueType, valueColumnName);

        return new DataFrame(keyColumn, valueColumn);
    }

    public static string GetSafeColumnName(this DataFrame dataFrame, string baseName)
    {
        var columnName = baseName;
        while (dataFrame.Columns.IndexOf(columnName) >= 0)
            columnName = $"{columnName}";
        return columnName;
    }

    public static StringDataFrameColumn MakeCompositeColumn(this DataFrame dataFrame, List<string>? includeColumns = null)
    {
        var compositeIndex = dataFrame.Columns
            .Where(x => includeColumns is null || includeColumns.Contains(x.Name))
            .Select(x => x.ToStringEnumerable())
            .Aggregate((x, y) => x.Zip(y).Select(z => $"{z.First ?? "null"}_{z.Second ?? "null"}"));

        return new StringDataFrameColumn(dataFrame.GetSafeColumnName("composite"), compositeIndex);
    }
}
