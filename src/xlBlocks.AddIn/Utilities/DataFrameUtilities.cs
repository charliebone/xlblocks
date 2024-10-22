namespace XlBlocks.AddIn.Utilities;

using System;
using System.Collections;
using System.Linq;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Dna;

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

    public static DataFrameColumn CreateDataFrameColumn<T>(T obj, long length, string columnName = "constant")
    {
        return CreateDataFrameColumn(Enumerable.Repeat(obj, (int)length), columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn<T>(IEnumerable<T?> items, string columnName = "constant")
    {
        return CreateDataFrameColumn(items, typeof(T), columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn(IEnumerable items, string type, string columnName = "constant")
    {
        var columnType = ParamTypeConverter.StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
        return CreateDataFrameColumn(items, columnType, columnName);
    }

    public static DataFrameColumn CreateDataFrameColumn(IEnumerable items, Type type, string columnName = "constant")
    {
        if (type == typeof(bool))
        {
            return new BooleanDataFrameColumn(columnName, items.Cast<bool?>());
        }
        else if (type == typeof(double))
        {
            return new DoubleDataFrameColumn(columnName, items.Cast<double?>());
        }
        else if (type == typeof(float))
        {
            return new SingleDataFrameColumn(columnName, items.Cast<float?>());
        }
        else if (type == typeof(decimal))
        {
            return new DecimalDataFrameColumn(columnName, items.Cast<decimal?>());
        }
        else if (type == typeof(string))
        {
            return new StringDataFrameColumn(columnName, items.Cast<string?>());
        }
        else if (type == typeof(int))
        {
            return new Int32DataFrameColumn(columnName, items.Cast<int?>());
        }
        else if (type == typeof(uint))
        {
            return new UInt32DataFrameColumn(columnName, items.Cast<uint?>());
        }
        else if (type == typeof(long))
        {
            return new Int64DataFrameColumn(columnName, items.Cast<long?>());
        }
        else if (type == typeof(ulong))
        {
            return new UInt64DataFrameColumn(columnName, items.Cast<ulong?>());
        }
        else if (type == typeof(short))
        {
            return new Int16DataFrameColumn(columnName, items.Cast<short?>());
        }
        else if (type == typeof(ushort))
        {
            return new UInt16DataFrameColumn(columnName, items.Cast<ushort?>());
        }
        else if (type == typeof(char))
        {
            return new CharDataFrameColumn(columnName, items.Cast<char?>());
        }
        else if (type == typeof(sbyte))
        {
            return new SByteDataFrameColumn(columnName, items.Cast<sbyte?>());
        }
        else if (type == typeof(byte))
        {
            return new ByteDataFrameColumn(columnName, items.Cast<byte?>());
        }
        else if (type == typeof(DateTime))
        {
            return new DateTimeDataFrameColumn(columnName, items.Cast<DateTime?>());
        }

        throw new NotSupportedException($"cannot build data frame column of type '{type.Name}'");
    }

    public static DataFrameColumn CreateDataFrameColumn<T>(string columnName, long length = 0)
    {
        return CreateDataFrameColumn(typeof(T), columnName, length);
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
}
