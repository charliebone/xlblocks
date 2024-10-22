namespace XlBlocks.AddIn.Excel.DataTypes;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ReaLTaiizor.Controls;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataTypes_Table
{
    [return: CacheContents]
    [ExcelFunction(Description = "Build a table from a range", IsThreadSafe = true)]
    public static XlBlockTable? xbTable_Build(
        [ExcelArgument(Description = "A range of data to use for the table")] XlBlockRange dataRange)
    {
        return XlBlockTable.Build(dataRange);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a table from a range, specifying the column types and names", IsThreadSafe = true)]
    public static XlBlockTable? xbTable_BuildWithTypes(
        [ExcelArgument(Description = "A range of data to use for the table")] XlBlockRange dataRange,
        [ExcelArgument(Description = "A range of column types")] XlBlockRange columnTypeRange,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange columnNameRange)
    {
        return XlBlockTable.BuildWithTypes(dataRange, columnTypeRange, columnNameRange);
    }

    [ExcelFunction(Description = "Get a table", IsThreadSafe = true)]
    public static object[,] xbTable_Get(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Include header in output (TRUE)")] bool includeHeader = true)
    {
        return table.AsArray(includeHeader);
    }

    [ExcelFunction(Description = "Get a column in a table", IsThreadSafe = true)]
    public static object[,] xbTable_GetColumn(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The column name or number")] object columnNameOrNumber,
        [ExcelArgument(Description = "Include header in output (TRUE)")] bool includeHeader = true)
    {
        if (columnNameOrNumber is double columnNumber)
            return table.AsArray(includeHeader, columnNumber: (int)columnNumber);

        return table.AsArray(includeHeader, columnName: columnNameOrNumber.ToString());
    }

    [ExcelFunction(Description = "Get the column names of a table", IsThreadSafe = true)]
    public static XlBlockRange xbTable_ColumnNames(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return XlBlockRange.Build(table.ColumnNames);
    }

    [ExcelFunction(Description = "Get the column types of a table", IsThreadSafe = true)]
    public static XlBlockRange xbTable_ColumnTypes(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return XlBlockRange.Build(table.ColumnNames);
    }

    [ExcelFunction(Description = "Get the number of rows in a table", IsThreadSafe = true)]
    public static long xbTable_RowCount(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return table.RowCount;
    }

    [ExcelFunction(Description = "Get the number of columns in a table", IsThreadSafe = true)]
    public static int xbTable_ColumnCount(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return table.ColumnCount;
    }

    [ExcelFunction(Description = "Get the dimensions of a table as a string", IsThreadSafe = true)]
    public static string xbTable_Dim(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return $"{table.RowCount:N0}x{table.ColumnCount:N0}";
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Join two tables", IsThreadSafe = true)]
    public static XlBlockTable xbTable_Join(
        [ExcelArgument(Description = "The left table to join"), CacheContents(AsReference = true)] XlBlockTable leftTable,
        [ExcelArgument(Description = "The right table to join"), CacheContents(AsReference = true)] XlBlockTable rightTable,
        [ExcelArgument(Description = "The join type, one of 'full', 'inner', 'right' or 'left'")] string joinType,
        [ExcelArgument(Description = "The keys to join on. Optional, defaults to all common columns'")] XlBlockRange? joinOn,
        [ExcelArgument(Description = "The suffix to apply to shared non-key columns from the left table. Optional ('.left')'")] string? leftSuffix = ".left",
        [ExcelArgument(Description = "The suffix to apply to shared non-key columns from the right table. Optional ('.left')")] string? rightSuffix = ".right",
        [ExcelArgument(Description = "Include both sets of identical joined columns in output (FALSE)")] bool includeDuplicateJoinColumns = false)
    {
        return XlBlockTable.Join(leftTable, rightTable, joinType, joinOn, leftSuffix, rightSuffix, includeDuplicateJoinColumns);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Drop rows containing nulls from a table", IsThreadSafe = true)]
    public static XlBlockTable xbTable_DropNulls(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Drop behavior, 'any' to drop rows containing any nulls, 'all' to drop rows containing only nulls")] string dropNullBehavior)
    {
        return table.DropNulls(dropNullBehavior);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Sort a table by one or more columns", IsThreadSafe = true)]
    public static XlBlockTable xbTable_Sort(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Columns to sort on")] XlBlockRange sortColumns,
        [ExcelArgument(Description = "Optional range to sort in descending order. (FALSE)"), Optional] XlBlockRange? descending,
        [ExcelArgument(Description = "Optional flag to sort null values first, (FALSE)"), Optional] XlBlockRange? nullsFirst)
    {
        return table.Sort(sortColumns, descending, nullsFirst);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Filter a table by values in a column", IsThreadSafe = true)]
    public static XlBlockTable xbTable_Filter(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Column to filter on")] string filterColumn,
        [ExcelArgument(Description = "The value to filter on")] object filterValue,
        [ExcelArgument(Description = "Optional flag indicating whether filter is inclusive (TRUE)")] bool inclusive = true)
    {
        return table.Filter(filterColumn, filterValue, inclusive);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Filter a table with an expression", IsThreadSafe = true)]
    public static XlBlockTable xbTable_FilterWith(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A filter expression")] string filterExpression)
    {
        return table.Filter(filterExpression);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Create a dictionary from two table columns", IsThreadSafe = true)]
    public static XlBlockDictionary xbTable_ToDict(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The name of the column to use for keys")] string keyColumnName,
        [ExcelArgument(Description = "The name of the column to use for values")] string valueColumnName)
    {
        return table.ToDictionary(keyColumnName, valueColumnName);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Append a column to a table from a dictionary", IsThreadSafe = true)]
    public static XlBlockTable xbTable_AppendColumnFromDict(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dictionary,
        [ExcelArgument(Description = "The name of the column used to match the dictionary keys")] string keyColumnName,
        [ExcelArgument(Description = "The name to use for the appended value column")] string valueColumnName,
        [ExcelArgument(Description = "Optional, the type to use for the value column")] string? valueColumnType = null)
    {
        return table.AppendColumnFromDictionary(dictionary, keyColumnName, valueColumnName, valueColumnType);
    }

    [return: CacheContents(CacheCollectionElements = true)]
    [ExcelFunction(Description = "Append a column to a table from a dictionary", IsThreadSafe = true)]
    public static XlBlockDictionary xbTable_ToDictofDicts(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The name of the key column")] string keyColumnName)
    {
        return table.ToDictionaryOfDictionaries(keyColumnName);
    }
}
