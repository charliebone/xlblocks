namespace XlBlocks.AddIn.Types;

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal class XlBlockTable : IXlBlockCopyableObject<XlBlockTable>, IXlBlockArrayableObject, IEnumerable<DataFrameRow>, IEnumerable
{
    private readonly DataFrame _dataFrame;

    public long RowCount => _dataFrame.Rows.Count;

    public int ColumnCount => _dataFrame.Columns.Count;

    public Type[] ColumnTypes => _dataFrame.Columns.Select(x => x.DataType).ToArray();

    public string[] ColumnNames => _dataFrame.Columns.Select(x => x.Name).ToArray();

    internal XlBlockTable(DataFrame dataFrame)
    {
        _dataFrame = dataFrame;
    }

    public IEnumerator<DataFrameRow> GetEnumerator() => _dataFrame.Rows.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => _dataFrame.Rows.AsEnumerable().GetEnumerator();

    public object? this[int row, int column] => _dataFrame.Rows[row][column];

    public object? this[int row, string columnName] => _dataFrame.Rows[row][columnName];

    public bool ContainsColumn(string columnName) => _dataFrame.Columns.IndexOf(columnName) >= 0;

    private void AssertColumnExists(string columnName)
    {
        if (!ContainsColumn(columnName))
            throw new ArgumentException($"column '{columnName}' does not exist in table");
    }

    private void AssertColumnNotExists(string columnName)
    {
        if (ContainsColumn(columnName))
            throw new ArgumentException($"column '{columnName}' already exists in table");
    }

    public XlBlockTable Copy()
    {
        return new XlBlockTable(_dataFrame.Clone());
    }

    public override string ToString()
    {
        return $"An XlBlocks table with {RowCount} rows and {ColumnCount} columns";
    }

    public XlBlockList GetColumnAsList(string? columnName)
    {
        if (columnName == null)
            throw new ArgumentNullException(nameof(columnName));

        var column = _dataFrame.Columns[columnName];
        return XlBlockList.BuildTyped(column.AsEnumerable(), column.DataType);
    }

    public XlBlockList GetColumnAsList(int columnNumber)
    {
        var column = _dataFrame.Columns[columnNumber];
        return XlBlockList.BuildTyped(column.AsEnumerable(), column.DataType);
    }

    public object[,] AsArray(RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        return AsArray(true, null, null, orientation);
    }

    public object[,] AsArray(bool includeHeader, string? columnName = null, int? columnNumber = null, RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        if (columnName is not null && columnNumber is not null)
            throw new ArgumentException($"only one of '{nameof(columnName)}' and '{nameof(columnNumber)}' may be provided");

        if (columnName is not null)
            AssertColumnExists(columnName);

        var outputColumns = _dataFrame.Columns.Where((x, i) => (columnName is null || columnName == x.Name) && (columnNumber is null || columnNumber == (i + 1))).ToList();

        var outputArray = orientation == RangeOrientation.ByColumn ?
            new object[RowCount + (includeHeader ? 1 : 0), outputColumns.Count] :
            new object[outputColumns.Count, RowCount + (includeHeader ? 1 : 0)];

        var col = 0;
        foreach (var column in outputColumns)
        {
            if (includeHeader)
            {
                if (orientation == RangeOrientation.ByColumn)
                    outputArray[0, col] = column.Name;
                else
                    outputArray[col, 0] = column.Name;
            }

            var offset = includeHeader ? 1 : 0;
            for (var row = 0; row < _dataFrame.Rows.Count; row++)
            {
                if (orientation == RangeOrientation.ByColumn)
                {
                    outputArray[row + offset, col] = column[row] ?? ExcelError.ExcelErrorNA;
                }
                else
                {
                    outputArray[col, row + offset] = column[row] ?? ExcelError.ExcelErrorNA;
                }
            }
            col++;
        }

        return outputArray;
    }

    public static XlBlockTable Build(XlBlockRange dataRange)
    {
        var columnNames = dataRange.GetRowAs<string>(0, false).ToList();

        var columns = new List<DataFrameColumn>();
        for (var col = 0; col < dataRange.ColumnCount; col++)
        {
            var guessConversions = dataRange.GetColumn(col).Skip(1).ConvertToBestTypes();
            var determinedType = guessConversions.Where(x => !x.IsMissingOrError)
                .Select(x => x.ConvertedType)
                .DetermineBestType();

            var converted = guessConversions.Select(x =>
            {
                if (x.IsMissingOrError)
                    return null;

                if (x.Success && x.ConvertedType == determinedType)
                    return x.ConvertedInput;

                if (ParamTypeConverter.TryConvertToProvidedType(x.Input, determinedType, out var convertedInput))
                    return convertedInput;

                return null;
            });

            var dataColumn = DataFrameUtilities.CreateDataFrameColumn(converted, determinedType, columnNames[col]);
            columns.Add(dataColumn);
        }

        var dataFrame = new DataFrame(columns.ToArray()).TrimNullRows();
        return new XlBlockTable(dataFrame);
    }

    public static XlBlockTable BuildWithTypes(XlBlockRange dataRange, XlBlockRange columnTypeRange, XlBlockRange columnNameRange)
    {
        if (dataRange.ColumnCount != columnTypeRange.Count)
            throw new ArgumentException($"{nameof(columnTypeRange)} length must equal number of {nameof(dataRange)} columns");

        if (dataRange.ColumnCount != columnNameRange.Count)
            throw new ArgumentException($"{nameof(columnNameRange)} length must equal number of {nameof(dataRange)} columns");

        var columnNames = columnNameRange.GetAs<string>(false).ToList();
        var columnTypes = columnTypeRange.GetAs<string>(false)
            .Select(x => ParamTypeConverter.StringToType(x) ?? throw new ArgumentException($"specified column type '{x}' is not valid"))
            .ToList();

        var columns = new List<DataFrameColumn>();
        for (var col = 0; col < dataRange.ColumnCount; col++)
        {
            var converted = dataRange.GetColumn(col).ConvertToProvidedType(columnTypes[col])
                .Select((x, row) =>
                {
                    if (x.IsMissingOrError)
                        return null;

                    if (!x.Success)
                        throw new ArgumentException($"cannot convert value '{x.Input}' into type '{columnTypes[col].Name}' [column '{columnNames[col]}', row {row + 1}]");

                    return x.ConvertedInput;
                });

            var dataColumn = DataFrameUtilities.CreateDataFrameColumn(converted, columnTypes[col], columnNames[col]);
            columns.Add(dataColumn);
        }

        var dataFrame = new DataFrame(columns.ToArray()).TrimNullRows();
        return new XlBlockTable(dataFrame);
    }

    public static XlBlockTable BuildFromDictionary(XlBlockDictionary dictionary, string keyColumnName, string valueColumnName, string? valueType = null)
    {
        var dataFrame = DataFrameUtilities.DictionaryToDataFrame(dictionary, keyColumnName, valueColumnName, valueType).TrimNullRows();
        return new XlBlockTable(dataFrame);
    }

    private static Type GuessTypeFunction(IEnumerable<string> columnValues)
    {
        var guessConversions = columnValues
            .Where(x => !string.IsNullOrEmpty(x) && !string.Equals(x, "null", StringComparison.OrdinalIgnoreCase))
            .ConvertToBestTypes();

        var determinedType = guessConversions.Where(x => !x.IsMissingOrError)
            .Select(x => x.ConvertedType)
            .DetermineBestType();

        return determinedType;
    }

    public static XlBlockTable BuildFromCsv(string csvPath, string separator, bool hasHeader, XlBlockRange? columnNameRange = null,
        XlBlockRange? columnTypeRange = null, string? encoding = null)
    {
        if (!File.Exists(csvPath))
            throw new FileNotFoundException($"file '{csvPath}' was not found");

        if (separator.Length != 1)
            throw new ArgumentException("separator must be a single character");

        var columnNames = columnNameRange?.GetAs<string>(false).ToArray();
        var columnTypes = columnTypeRange?.GetAs<string>(false).Select(ParamTypeConverter.StringToType).ToArray();

        if (columnNames is not null && columnTypes is not null && columnNames.Length != columnTypes.Length)
            throw new ArgumentException("Column names and column types must be same length");

        using var csvStream = new FileStream(csvPath, FileMode.Open);
        var dataFrame = DataFrame.LoadCsv(csvStream, separator[0], hasHeader, columnNames, columnTypes,
            addIndexColumn: false,
            encoding: !string.IsNullOrEmpty(encoding) ? Encoding.GetEncoding(encoding) : Encoding.UTF8,
            guessTypeFunction: GuessTypeFunction);
        return new XlBlockTable(dataFrame);
    }

    public void SaveToCsv(string csvPath, string separator, bool includeHeader, string? encoding = null)
    {
        using var csvStream = new FileStream(csvPath, FileMode.Create);
        SaveToCsv(csvStream, separator, includeHeader, encoding);
    }

    internal void SaveToCsv(Stream stream, string separator, bool includeHeader, string? encoding = null)
    {
        if (separator.Length != 1)
            throw new ArgumentException("separator must be a single character");

        DataFrame.SaveCsv(_dataFrame, stream, separator[0], includeHeader, encoding: !string.IsNullOrEmpty(encoding) ? Encoding.GetEncoding(encoding) : Encoding.UTF8);
    }

    public static XlBlockTable Join(XlBlockTable left, XlBlockTable right, string joinType, XlBlockRange? joinOn,
        string? leftSuffix = ".left", string? rightSuffix = ".right", bool includeDuplicateJoinColumns = false)
    {
        // dataframe has a join method which joins on a numeric index column, and a merge method which joins on the column values as one would expect
        var joinAlgorithm = DataFrameUtilities.ParseJoinType(joinType);
        if (leftSuffix == rightSuffix)
            throw new ArgumentException("left and right suffixes must be unique");

        if (joinOn != null)
        {
            if (joinOn.ColumnCount > 2)
                throw new ArgumentException("joinOn must have either one or two columns");

            if (joinOn.ColumnCount == 2)
            {
                var leftColumns = joinOn.GetColumnAs<string>(0, false).ToArray();
                var rightColumns = joinOn.GetColumnAs<string>(1, false).ToArray();
                var merged = DataFrameUtilities.Merge(left._dataFrame, right._dataFrame, leftColumns, rightColumns, leftSuffix, rightSuffix, joinAlgorithm);
                return new XlBlockTable(merged);
            }
            else
            {
                var joinColumns = joinOn.GetAs<string>(false).ToArray();
                var merged = DataFrameUtilities.Merge(left._dataFrame, right._dataFrame, joinColumns, joinColumns, leftSuffix, rightSuffix, joinAlgorithm);
                if (!includeDuplicateJoinColumns && joinAlgorithm != JoinAlgorithm.FullOuter)
                {
                    foreach (var duplicatedColumn in joinColumns)
                    {
                        merged.Columns.Remove($"{duplicatedColumn}{(joinAlgorithm == JoinAlgorithm.Right ? leftSuffix : rightSuffix)}");
                        merged.Columns[$"{duplicatedColumn}{(joinAlgorithm == JoinAlgorithm.Right ? rightSuffix : leftSuffix)}"].SetName(duplicatedColumn);
                    }
                }
                return new XlBlockTable(merged);
            }
        }
        else
        {
            var commonColumns = left.ColumnNames.Intersect(right.ColumnNames).ToArray();
            if (commonColumns.Length == 0)
                throw new ArgumentException("cannot find common columns to join on and no join columns specified");

            var merged = DataFrameUtilities.Merge(left._dataFrame, right._dataFrame, commonColumns, commonColumns, leftSuffix, rightSuffix, joinAlgorithm);
            if (!includeDuplicateJoinColumns && joinAlgorithm != JoinAlgorithm.FullOuter)
            {
                foreach (var duplicatedColumn in commonColumns)
                {
                    merged.Columns.Remove($"{duplicatedColumn}{(joinAlgorithm == JoinAlgorithm.Right ? leftSuffix : rightSuffix)}");
                    merged.Columns[$"{duplicatedColumn}{(joinAlgorithm == JoinAlgorithm.Right ? rightSuffix : leftSuffix)}"].SetName(duplicatedColumn);
                }
            }
            return new XlBlockTable(merged);
        }
    }

    public static XlBlockTable UnionAll(params XlBlockTable[] tables)
    {
        if (tables.Length == 0)
            throw new ArgumentException("Must provide at least one table to append rows from");

        var dataFrame = tables[0]._dataFrame.Clone();
        var columnNamesSet = dataFrame.Columns.Select(x => x.Name).ToHashSet();
        foreach (var table in tables.Skip(1))
        {
            foreach (var columnName in table.ColumnNames)
                if (!columnNamesSet.Contains(columnName))
                    throw new ArgumentException($"{columnName} does not exist in every table being unioned");

            dataFrame.Append(table, true);
        }

        return new XlBlockTable(dataFrame);
    }

    public static XlBlockTable Union(params XlBlockTable[] tables)
    {
        var unioned = UnionAll(tables);

        var duplicates = unioned._dataFrame.MakeCompositeColumn().IsDuplicateElement();
        return new XlBlockTable(unioned._dataFrame.Filter(duplicates.ElementwiseEquals(false)));
    }

    public XlBlockTable DropNulls(string dropNullBehavior = "all")
    {
        var dropNullOptions = DataFrameUtilities.ParseDropNullOption(dropNullBehavior);
        return new XlBlockTable(_dataFrame.DropNulls(dropNullOptions));
    }

    public object LookupValue(string lookupColumnName, object lookupValue, string valueColumnName, string onMultipleMatches = "error")
    {
        AssertColumnExists(lookupColumnName);
        AssertColumnExists(valueColumnName);

        var lookupColumn = _dataFrame[lookupColumnName];
        if (!ParamTypeConverter.TryConvertToProvidedType(lookupValue, lookupColumn.DataType, out var converted))
            throw new ArgumentException($"could not convert value '{lookupValue}' to lookup column type '{lookupColumn.DataType.Name}'");

        var valueColumn = DataFrameUtilities.CreateConstantDataFrameColumn(converted, lookupColumn.DataType, lookupColumn.Length);
        var matchIndex = lookupColumn.ElementwiseEquals(valueColumn);
        var matchingData = _dataFrame[matchIndex];
        if (matchingData.Rows.Count > 1)
        {
            var multipleMatchBehavior = DataFrameUtilities.ParseDuplicateKeyBehavior(onMultipleMatches);
            return multipleMatchBehavior switch
            {
                DataFrameUtilities.DuplicateKeyBehavior.TakeFirst => matchingData[valueColumnName][0],
                DataFrameUtilities.DuplicateKeyBehavior.TakeLast => matchingData[valueColumnName][matchingData.Rows.Count - 1],
                _ => throw new ArgumentException("multiple matches found"),
            };
        }
        else if (matchingData.Rows.Count == 1)
        {
            return matchingData[valueColumnName][0];
        }
        else
        {
            // should we add further behavior here ala excel MATCH and XLOOKUP?
            throw new ArgumentException("no matching rows found");
        }
    }

    public XlBlockTable Sort(XlBlockRange sortColumnRange, XlBlockRange? isDescendingRange, XlBlockRange? nullsFirstRange)
    {
        // dataframe sorting is not stable, see https://github.com/dotnet/machinelearning/issues/6443
        // because of this, we can't use for iterative multi-column sorting
        // so we must use our own sort. it will certainly be slower, so hopefully ml.net addresses this at some point

        if (isDescendingRange is not null && sortColumnRange.Count != isDescendingRange.Count)
            throw new ArgumentException("isDescending range must be same length of sort column range");

        if (nullsFirstRange is not null && sortColumnRange.Count != nullsFirstRange.Count)
            throw new ArgumentException("nullsFirst range must be same length of sort column range");

        var sortColumns = sortColumnRange.GetAs<string>(false);
        var isDescending = isDescendingRange?.GetAs<bool>(false) ?? Enumerable.Repeat(false, sortColumnRange.Count);
        var isNullsFirst = nullsFirstRange?.GetAs<bool>(false) ?? Enumerable.Repeat(false, sortColumnRange.Count);

        Int64DataFrameColumn indexColumn = null!;
        foreach (var (columnName, descending, nullsFirst) in sortColumns.Zip(isDescending, isNullsFirst).Reverse())
        {
            AssertColumnExists(columnName);

            var column = _dataFrame.Columns[columnName];
            indexColumn = column.SortColumn(indexColumn, descending, nullsFirst);
        }

        var newColumns = new List<DataFrameColumn>(_dataFrame.Columns.Count);
        for (var i = 0; i < _dataFrame.Columns.Count; i++)
        {
            var oldColumn = _dataFrame.Columns[i];
            var newColumn = oldColumn.Clone(indexColumn, false, 0);

            newColumns.Add(newColumn);
        }
        return new XlBlockTable(new DataFrame(newColumns));
    }

    public XlBlockTable Filter(string expression)
    {
        if (string.IsNullOrEmpty(expression))
            return this;

        var parsed = DataFrameExpressionParser.Parser.Parse(expression);
        if (parsed.IsError)
            throw new ArgumentException($"error parsing filter expression: {string.Join("|", parsed.Errors.Select(x => x.ErrorMessage))}");

        var context = new DataFrameContext(_dataFrame);
        var resultColumn = parsed.Result.Evaluate(context);
        if (resultColumn is not PrimitiveDataFrameColumn<bool> filter)
            throw new Exception("filter expression must evaluate to a boolean");

        return new XlBlockTable(_dataFrame.Filter(filter));
    }

    public XlBlockTable Filter(string columnName, object? value, bool inclusive = true)
    {
        if (value is null)
            throw new ArgumentNullException(nameof(value));

        AssertColumnExists(columnName);

        var column = _dataFrame[columnName];
        if (!ParamTypeConverter.TryConvertToProvidedType(value, column.DataType, out var converted))
            throw new ArgumentException($"could not convert value '{value}' to type '{column.DataType.Name}'");

        var valueColumn = DataFrameUtilities.CreateConstantDataFrameColumn(converted, column.DataType, column.Length);
        var filter = inclusive ? column.ElementwiseEquals(valueColumn) : column.ElementwiseNotEquals(valueColumn);
        return new XlBlockTable(_dataFrame.Filter(filter));
    }

    public XlBlockTable AppendColumnsWith(XlBlockRange columnNamesRange, XlBlockRange columnExpressionsRange)
    {
        var columnNames = columnNamesRange.GetAs<string>(false).ToList();
        var columnExpressions = columnExpressionsRange.GetAs<string>(false).ToList();

        if (columnNames.Count != columnExpressions.Count)
            throw new ArgumentException("column names range must be same length of column expressions range");

        var dataFrame = _dataFrame.Clone();
        var context = new DataFrameContext(dataFrame);
        for (var i = 0; i < columnNames.Count; i++)
        {
            AssertColumnNotExists(columnNames[i]);

            var parsed = DataFrameExpressionParser.Parser.Parse(columnExpressions[i]);
            if (parsed.IsError)
                throw new ArgumentException($"error parsing expression for column '{columnNames[i]}': {string.Join("|", parsed.Errors.Select(x => x.ErrorMessage))}");

            var resultColumn = parsed.Result.Evaluate(context);
            if (resultColumn is NullDataFrameColumn)
                resultColumn = new BooleanDataFrameColumn(columnNames[i], dataFrame.Rows.Count);
            else
                resultColumn.SetName(columnNames[i]);

            dataFrame.Columns.Add(resultColumn);
        }

        return new XlBlockTable(dataFrame);
    }

    public XlBlockTable AppendColumnFromList(XlBlockList list, string columnName, string? columnType = null)
    {
        AssertColumnNotExists(columnName);
        if (list.Count != RowCount)
            throw new ArgumentException("list must have same number of rows as table");

        var dataFrame = _dataFrame.Clone();
        var listColumn = DataFrameUtilities.CreateDataFrameColumn(list.Cast<object>(), columnType, columnName);
        dataFrame.Columns.Add(listColumn);
        return new XlBlockTable(dataFrame);
    }

    public XlBlockTable AppendColumnFromDictionary(XlBlockDictionary dictionary, string keyColumnName, string valueColumnName,
        string? valueType = null, object? valueOnMissing = null)
    {
        AssertColumnExists(keyColumnName);
        AssertColumnNotExists(valueColumnName);

        var rightDataFrame = DataFrameUtilities.DictionaryToDataFrame(dictionary, keyColumnName, valueColumnName, valueType);

        var joinColumns = new[] { keyColumnName };
        var merged = DataFrameUtilities.Merge(_dataFrame, rightDataFrame, joinColumns, joinColumns, "_left", "_right", JoinAlgorithm.Left);
        merged.Columns.Remove($"{keyColumnName}_right");
        merged.Columns[$"{keyColumnName}_left"].SetName(keyColumnName);

        if (valueOnMissing != null)
        {
            var valueOnMissingColumn = DataFrameUtilities.CreateDataFrameColumn(valueOnMissing, merged.Rows.Count, valueType);
            var newColumn = merged.Columns[valueColumnName].ElementwiseIsNotNull().ElementwiseIfThenElse(merged.Columns[valueColumnName], valueOnMissingColumn);
            newColumn.SetName(valueColumnName);
            merged.Columns.Remove(valueColumnName);
            merged.Columns.Add(newColumn);
        }

        return new XlBlockTable(merged);
    }

    public XlBlockDictionary ToDictionary(string keyColumnName, string valueColumnName, string onDuplicateKeys = "error")
    {
        AssertColumnExists(keyColumnName);
        AssertColumnExists(valueColumnName);

        var duplicateKeyBehavior = DataFrameUtilities.ParseDuplicateKeyBehavior(onDuplicateKeys);
        var keyColumn = _dataFrame[keyColumnName];
        var valueColumn = _dataFrame[valueColumnName];

        var dict = DictionaryUtilities.BuildTypedDictionary(keyColumn.DataType);
        for (var i = 0L; i < keyColumn.Length; i++)
        {
            if (keyColumn[i] is null)
                continue;

            switch (duplicateKeyBehavior)
            {
                case DataFrameUtilities.DuplicateKeyBehavior.TakeFirst:
                    if (!dict.Contains(keyColumn[i]))
                        dict.Add(keyColumn[i], valueColumn[i]);
                    break;
                case DataFrameUtilities.DuplicateKeyBehavior.TakeLast:
                    dict[keyColumn[i]] = valueColumn[i];
                    break;
                case DataFrameUtilities.DuplicateKeyBehavior.Error:
                default:
                    dict.Add(keyColumn[i], valueColumn[i]);
                    break;
            }
        }

        return new XlBlockDictionary(dict, keyColumn.DataType);
    }

    public XlBlockDictionary ToDictionaryOfDictionaries(string keyColumnName, string onDuplicateKeys = "error")
    {
        AssertColumnExists(keyColumnName);

        var duplicateKeyBehavior = DataFrameUtilities.ParseDuplicateKeyBehavior(onDuplicateKeys);
        var keyColumn = _dataFrame[keyColumnName];

        var dict = DictionaryUtilities.BuildTypedDictionary(keyColumn.DataType);
        for (var row = 0L; row < keyColumn.Length; row++)
        {
            if (keyColumn[row] is null)
                continue;

            if (dict.Contains(keyColumn[row]))
            {
                if (duplicateKeyBehavior == DataFrameUtilities.DuplicateKeyBehavior.TakeFirst)
                    continue;

                if (duplicateKeyBehavior == DataFrameUtilities.DuplicateKeyBehavior.Error)
                    throw new ArgumentException($"duplicate key in column {keyColumnName}: '{keyColumn[row]}'");
            }

            var rowDict = new Dictionary<string, object>();
            foreach (var column in _dataFrame.Columns)
                rowDict[column.Name] = column[row];

            dict[keyColumn[row]] = new XlBlockDictionary(rowDict, typeof(string));
        }

        return new XlBlockDictionary(dict, keyColumn.DataType);
    }

    public XlBlockTable Project(XlBlockRange currentColumnNamesRange, XlBlockRange? newColumnNamesRange, XlBlockRange? newColumnTypesRange)
    {
        if (newColumnNamesRange is not null && currentColumnNamesRange.Count != newColumnNamesRange.Count)
            throw new ArgumentException("new column names range must be same length of current column names range");

        if (newColumnTypesRange is not null && currentColumnNamesRange.Count != newColumnTypesRange.Count)
            throw new ArgumentException("new column types range must be same length of current column names range");

        var currentColumnNames = currentColumnNamesRange.GetAs<string>(false);
        var newColumnNames = newColumnNamesRange?.GetAs<string>(false) ?? currentColumnNames;
        var columnTypes = newColumnTypesRange?.GetAs<string>(false) ?? Enumerable.Repeat(string.Empty, currentColumnNamesRange.Count);

        var columns = new List<DataFrameColumn>();
        foreach (var (currentColumnName, newColumnName, columnType) in currentColumnNames.Zip(newColumnNames, columnTypes))
        {
            AssertColumnExists(currentColumnName);

            var column = _dataFrame[currentColumnName];
            var type = (string.IsNullOrEmpty(columnType) ? column.DataType : ParamTypeConverter.StringToType(columnType)) ?? throw new ArgumentException($"unknown type '{columnType}'");

            var newColumn = column.ConvertColumnType(type);
            newColumn.SetName(newColumnName);
            columns.Add(newColumn);
        }

        var dataFrame = new DataFrame(columns);
        return new XlBlockTable(dataFrame);
    }

    public XlBlockTable GroupBy(XlBlockRange groupColumnNamesRange, string groupByOperation, XlBlockRange? aggregationColumnNamesRange, XlBlockRange? newColumnNamesRange)
    {
        return GroupBy(groupColumnNamesRange, new List<string>() { groupByOperation }, aggregationColumnNamesRange, newColumnNamesRange);
    }

    public XlBlockTable GroupBy(XlBlockRange groupColumnNamesRange, XlBlockRange groupByOperations, XlBlockRange? aggregationColumnNamesRange, XlBlockRange? newColumnNamesRange)
    {
        if (groupByOperations.Count == 0)
            throw new ArgumentException("at least one group by operation must be specified");

        return GroupBy(groupColumnNamesRange, groupByOperations.GetAs<string>(false).ToList(), aggregationColumnNamesRange, newColumnNamesRange);
    }

    private XlBlockTable GroupBy(XlBlockRange groupColumnNamesRange, List<string> groupByOperations, XlBlockRange? aggregationColumnNamesRange, XlBlockRange? newColumnNamesRange)
    {
        var groupColumnNames = groupColumnNamesRange.GetAs<string>(false).ToList();
        if (groupColumnNames.Count == 0)
            throw new ArgumentException("at least one group column must be specified");

        foreach (var column in groupColumnNames)
            AssertColumnExists(column);

        var aggregationColumnNames = aggregationColumnNamesRange?.GetAs<string>(false)?.ToList();
        if (aggregationColumnNames is null || !aggregationColumnNames.Any())
        {
            if (groupByOperations.Count > 1)
                throw new ArgumentException("multiple group by operations only allowed if aggregation columns are specified");

            aggregationColumnNames = new List<string>();
            foreach (var column in _dataFrame.Columns)
            {
                if (groupColumnNames.Contains(column.Name))
                    continue;

                if (column.IsNumericColumn())
                    aggregationColumnNames.Add(column.Name);
            }
            groupByOperations = Enumerable.Repeat(groupByOperations[0], aggregationColumnNames.Count).ToList();
        }
        else
        {
            if (groupByOperations.Count == 1)
                groupByOperations = Enumerable.Repeat(groupByOperations[0], aggregationColumnNames.Count).ToList();

            if (groupByOperations.Count != aggregationColumnNames.Count)
                throw new ArgumentException("group by operations list must have same length as aggregate columns list");

            foreach (var column in aggregationColumnNames)
                AssertColumnExists(column);
        }

        var newColumnNames = newColumnNamesRange?.GetAs<string>(false)?.ToList() ?? aggregationColumnNames.Zip(groupByOperations).Select(x => $"{x.First}.{x.Second}").ToList();
        if (newColumnNames.Any() && newColumnNames.Count != aggregationColumnNames.Count)
            throw new ArgumentException($"new column names list count ({newColumnNames.Count}) does not match aggregation column count ({aggregationColumnNames.Count})");

        var newDataFrame = DataFrameUtilities.ComputeGroupAggregations(_dataFrame, groupColumnNames, groupByOperations, aggregationColumnNames, newColumnNames);
        return new XlBlockTable(newDataFrame);
    }
}
