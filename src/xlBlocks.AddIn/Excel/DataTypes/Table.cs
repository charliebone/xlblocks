namespace XlBlocks.AddIn.Excel.DataTypes;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataTypes_Table
{
    [return: CacheContents]
    [ExcelFunction(Description = "Build a table from a range", IsThreadSafe = true)]
    public static XlBlockTable? XBTable_Build(
        [ExcelArgument(Description = "A range of data to use for the table")] XlBlockRange dataRange)
    {
        return XlBlockTable.Build(dataRange);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a table from a range, specifying the column types and names", IsThreadSafe = true)]
    public static XlBlockTable? XBTable_BuildWithTypes(
        [ExcelArgument(Description = "A range of data to use for the table")] XlBlockRange dataRange,
        [ExcelArgument(Description = "A range of column types")] XlBlockRange columnTypeRange,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange columnNameRange)
    {
        return XlBlockTable.BuildWithTypes(dataRange, columnTypeRange, columnNameRange);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a table from a delimited file", IsThreadSafe = true)]
    public static XlBlockTable? XBTable_BuildFromCsv(
        [ExcelArgument(Description = "A filepath")] string filePath,
        [ExcelArgument(Description = "The delimiter to use")] string delimiter = ",",
        [ExcelArgument(Description = "Flag indicating whether the csv has a header row")] bool hasHeader = true,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange? columnNameRange = null,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange? columnTypeRange = null,
        [ExcelArgument(Description = "The file encoding")] string encoding = "utf-8")
    {
        return XlBlockTable.BuildFromCsv(filePath, delimiter, hasHeader, columnNameRange, columnTypeRange, encoding);
    }

    [ExcelFunction(Description = "Save a table to a delimited file", IsThreadSafe = true)]
    public static bool XBTable_SaveToCsv(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A filepath")] string filePath,
        [ExcelArgument(Description = "The delimiter to use")] string delimiter = ",",
        [ExcelArgument(Description = "Include a header row")] bool includeHeader = true,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange? columnNameRange = null,
        [ExcelArgument(Description = "A range of column names")] XlBlockRange? columnTypeRange = null,
        [ExcelArgument(Description = "The file encoding")] string encoding = "utf-8",
        [ExcelArgument(Description = "A relative or absolute path to a directory in which to save a timestamped copy of an existing original file")] string? archivePath = null)
    {
        table.SaveToCsv(filePath, delimiter, includeHeader, encoding, archivePath);
        return true;
    }

    [ExcelFunction(Description = "Get a table", IsThreadSafe = true)]
    public static object[,] XBTable_Get(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Include header in output")] bool includeHeader = true)
    {
        return table.AsArray(includeHeader);
    }

    [ExcelFunction(Description = "Get a column in a table", IsThreadSafe = true)]
    public static object[,] XBTable_GetColumn(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The column name or number")] object columnNameOrNumber,
        [ExcelArgument(Description = "Include header in output")] bool includeHeader = false)
    {
        if (columnNameOrNumber is double columnNumber)
            return table.AsArray(includeHeader, columnNumber: (int)columnNumber);

        return table.AsArray(includeHeader, columnName: columnNameOrNumber.ToString());
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Get a column in a table as a list", IsThreadSafe = true)]
    public static XlBlockList XBTable_GetColumnAsList(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The column name or number")] object columnNameOrNumber)
    {
        if (columnNameOrNumber is double columnNumber)
            return table.GetColumnAsList((int)columnNumber);

        return table.GetColumnAsList(columnNameOrNumber.ToString());
    }

    [ExcelFunction(Description = "Get the column names of a table", IsThreadSafe = true)]
    public static XlBlockRange XBTable_ColumnNames(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return XlBlockRange.Build(table.ColumnNames);
    }

    [ExcelFunction(Description = "Get the column types of a table", IsThreadSafe = true)]
    public static XlBlockRange XBTable_ColumnTypes(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return XlBlockRange.Build(table.ColumnTypes.Select(x => x.Name).ToArray());
    }

    [ExcelFunction(Description = "Get the number of rows in a table", IsThreadSafe = true)]
    public static long XBTable_RowCount(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return table.RowCount;
    }

    [ExcelFunction(Description = "Get the number of columns in a table", IsThreadSafe = true)]
    public static int XBTable_ColumnCount(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return table.ColumnCount;
    }

    [ExcelFunction(Description = "Get the dimensions of a table as a string", IsThreadSafe = true)]
    public static string XBTable_Dim(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table)
    {
        return $"{table.RowCount:N0}x{table.ColumnCount:N0}";
    }

    [ExcelFunction(Description = "Lookup a value in a table", IsThreadSafe = true)]
    public static object XBTable_LookupValue(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The name of the column in which to look up the value")] string lookupColumn,
        [ExcelArgument(Description = "The value to search for")] object lookupValue,
        [ExcelArgument(Description = "The name of the column in the matching row to return a value from")] string valueColumn,
        [ExcelArgument(Description = "Behavior on multiple matching rows, one of 'error', 'first' or 'last'")] string onMultipleMatches = "error")
    {
        return table.LookupValue(lookupColumn, lookupValue, valueColumn, onMultipleMatches);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Join two tables", IsThreadSafe = true)]
    public static XlBlockTable XBTable_Join(
        [ExcelArgument(Description = "The left table to join"), CacheContents(AsReference = true)] XlBlockTable leftTable,
        [ExcelArgument(Description = "The right table to join"), CacheContents(AsReference = true)] XlBlockTable rightTable,
        [ExcelArgument(Description = "The join type, one of 'full', 'inner', 'right' or 'left'")] string joinType,
        [ExcelArgument(Description = "The keys to join on, defaults to all common columns'")] XlBlockRange? joinOn = null,
        [ExcelArgument(Description = "The suffix to apply to shared non-key columns from the left table'")] string? leftSuffix = ".left",
        [ExcelArgument(Description = "The suffix to apply to shared non-key columns from the right table")] string? rightSuffix = ".right",
        [ExcelArgument(Description = "Include both sets of identical joined columns in output")] bool includeDuplicateJoinColumns = false)
    {
        return XlBlockTable.Join(leftTable, rightTable, joinType, joinOn, leftSuffix, rightSuffix, includeDuplicateJoinColumns);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Combine rows of multiple matching tables", IsThreadSafe = true)]
    public static XlBlockTable XBTable_UnionAll(
        [ExcelArgument(Name = "table", Description = "A table"), CacheContents(AsReference = true)] params XlBlockTable[] tables)
    {
        return XlBlockTable.UnionAll(tables);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Combine rows of multiple matching tables, removing duplicates", IsThreadSafe = true)]
    public static XlBlockTable XBTable_Union(
        [ExcelArgument(Name = "table", Description = "A table"), CacheContents(AsReference = true)] params XlBlockTable[] tables)
    {
        return XlBlockTable.Union(tables);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Combine multiple tables, returning a table with the superset of all columns", IsThreadSafe = true)]
    public static XlBlockTable XBTable_UnionSuperset(
        [ExcelArgument(Name = "table", Description = "A table"), CacheContents(AsReference = true)] params XlBlockTable[] tables)
    {
        return XlBlockTable.UnionSuperset(tables);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Drop rows containing nulls from a table", IsThreadSafe = true)]
    public static XlBlockTable XBTable_DropNulls(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Drop behavior, 'any' to drop rows containing any nulls, 'all' to drop rows containing only nulls")] string dropNullBehavior)
    {
        return table.DropNulls(dropNullBehavior);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Sort a table by one or more columns", IsThreadSafe = true)]
    public static XlBlockTable XBTable_Sort(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Columns to sort on")] XlBlockRange sortColumns,
        [ExcelArgument(Description = "Sort in descending order"), Optional] XlBlockRange? descending,
        [ExcelArgument(Description = "Sort null values first"), Optional] XlBlockRange? nullsFirst)
    {
        return table.Sort(sortColumns, descending, nullsFirst);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Filter a table by values in a column", IsThreadSafe = true)]
    public static XlBlockTable XBTable_Filter(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Column to filter on")] string filterColumn,
        [ExcelArgument(Description = "The value to filter on")] object filterValue,
        [ExcelArgument(Description = "Use inclusive filtering")] bool inclusive = true)
    {
        return table.Filter(filterColumn, filterValue, inclusive);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Filter a table with an expression", IsThreadSafe = true)]
    public static XlBlockTable XBTable_FilterWith(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A filter expression")] string filterExpression)
    {
        return table.Filter(filterExpression);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Filter a table with an expression", IsThreadSafe = true)]
    public static XlBlockTable XBTable_AppendColumnsWith(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Column names")] XlBlockRange columnNames,
        [ExcelArgument(Description = "Column expressions")] XlBlockRange columnExpressions)
    {
        return table.AppendColumnsWith(columnNames, columnExpressions);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Append a column to a table from a dictionary", IsThreadSafe = true)]
    public static XlBlockTable XBTable_AppendColumnFromList(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list,
        [ExcelArgument(Description = "The name to use for the new column")] string columnName,
        [ExcelArgument(Description = "The type to use for the new column column")] string? columnType = null)
    {
        return table.AppendColumnFromList(list, columnName, columnType);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Append a column to a table from a dictionary", IsThreadSafe = true)]
    public static XlBlockTable XBTable_AppendColumnFromDict(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dictionary,
        [ExcelArgument(Description = "The name of the column used to match the dictionary keys")] string keyColumnName,
        [ExcelArgument(Description = "The name to use for the value column")] string valueColumnName,
        [ExcelArgument(Description = "The type to use for the value column")] string? valueColumnType = null,
        [ExcelArgument(Description = "The value to use for keys not found in dictionary")] object? valueOnMissing = null)
    {
        return table.AppendColumnFromDictionary(dictionary, keyColumnName, valueColumnName, valueColumnType, valueOnMissing);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Create a dictionary from two table columns", IsThreadSafe = true)]
    public static XlBlockDictionary XBTable_ToDict(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The name of the column to use for keys")] string keyColumnName,
        [ExcelArgument(Description = "The name of the column to use for values")] string valueColumnName,
        [ExcelArgument(Description = "Behavior on duplicate keys, one of 'error', 'first' or 'last'")] string onDuplicateKeys = "error")
    {
        return table.ToDictionary(keyColumnName, valueColumnName, onDuplicateKeys);
    }

    [return: CacheContents(CacheCollectionElements = true)]
    [ExcelFunction(Description = "Append a column to a table from a dictionary", IsThreadSafe = true)]
    public static XlBlockDictionary XBTable_ToDictofDicts(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "The name of the key column")] string keyColumnName,
        [ExcelArgument(Description = "Behavior on duplicate keys, one of 'error', 'first' or 'last'")] string onDuplicateKeys = "error")
    {
        return table.ToDictionaryOfDictionaries(keyColumnName, onDuplicateKeys);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Project a table onto a new table, optionally changing column names and types", IsThreadSafe = true)]
    public static XlBlockTable XBTable_Projection(
        [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
        [ExcelArgument(Description = "Names of columns to include in the projection")] XlBlockRange currentColumnNames,
        [ExcelArgument(Description = "Names to use to rename the columns in the projection"), Optional] XlBlockRange? newColumnNames,
        [ExcelArgument(Description = "Types to use to convert the columns in the projection"), Optional] XlBlockRange? newColumnTypes)
    {
        return table.Project(currentColumnNames, newColumnNames, newColumnTypes);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Project a table onto a new table, optionally changing column names and types", IsThreadSafe = true)]
    public static XlBlockTable XBTable_GroupBy(
       [ExcelArgument(Description = "A table"), CacheContents(AsReference = true)] XlBlockTable table,
       [ExcelArgument(Description = "Names of columns to group by")] XlBlockRange groupByColumns,
       [ExcelArgument(Description = "Group by operation(s), one or more of 'sum', 'product', 'min' 'max' 'mean', 'median', 'count', 'counta', 'first', 'firsta', 'last', 'lasta', 'stddev', 'stddevp', 'var', 'varp', 'skew', 'skewp', 'kurt' or 'kurp'"), Optional] XlBlockRange groupByOperation,
       [ExcelArgument(Description = "Names of columns to aggregate, defaults to all numeric non-group columns"), Optional] XlBlockRange? aggregateColumns,
       [ExcelArgument(Description = "Names to use for the aggregate output columns, defaults to '[ColumnName].[Operation]'"), Optional] XlBlockRange? newColumnNames)
    {
        return table.GroupBy(groupByColumns, groupByOperation, aggregateColumns, newColumnNames);
    }
}
