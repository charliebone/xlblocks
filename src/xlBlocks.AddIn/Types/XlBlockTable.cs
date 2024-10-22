namespace XlBlocks.AddIn.Types;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Data.Analysis;
using xlBlocks.AddIn.Utilities;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal class XlBlockTable : IXlBlockCopyableObject<XlBlockTable>, IXlBlockArrayableObject, IEnumerable<DataFrameRow>, IEnumerable
{
    // TODO: consider implementation that uses Microsoft.DataAnalytics.DataFrame instead
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

    private void CheckColumnExists(string columnName)
    {
        if (!ContainsColumn(columnName))
            throw new ArgumentException($"column '{columnName}' does not exist in table");
    }

    private void CheckColumnNotExists(string columnName)
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

    public object[,] AsArray(RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        return AsArray(true, null, null, orientation);
    }

    public object[,] AsArray(bool includeHeader, string? columnName = null, int? columnNumber = null, RangeOrientation orientation = RangeOrientation.ByColumn)
    {
        if (columnName is not null && columnNumber is not null)
            throw new ArgumentException($"only one of '{nameof(columnName)}' and '{nameof(columnNumber)}' may be provided");

        if (columnName is not null)
            CheckColumnExists(columnName);

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

        var dataFrame = new DataFrame(columns.ToArray());
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
                        throw new ArgumentException($"cannot convert value '{x.Input}' into type '{columnTypes[col].Name}' [row {row + 1}, col {col + 1}]");

                    return x.ConvertedInput;
                });

            var dataColumn = DataFrameUtilities.CreateDataFrameColumn(converted, columnTypes[col], columnNames[col]);
            columns.Add(dataColumn);
        }

        var dataFrame = new DataFrame(columns.ToArray());
        return new XlBlockTable(dataFrame);
    }

    public static XlBlockTable MergeRows(params XlBlockTable[] tables)
    {
        // how should we check rows here?
        /*var first = tables.;
        if (first ==)
        var dataTable = tables.First().CopyToDataTable();
        foreach (var table in tables.Skip(1))
            dataTable.Merge(table._dataFrame);

        return new XlBlockTable(dataTable);*/
        return null!;
    }

    public static XlBlockTable Join(XlBlockTable left, XlBlockTable right, string joinType, XlBlockRange? joinOn,
        string? leftSuffix = ".left", string? rightSuffix = ".right", bool includeDuplicateJoinColumns = false)
    {
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
                return new XlBlockTable(left._dataFrame.Merge(right._dataFrame, leftColumns, rightColumns, leftSuffix, rightSuffix, joinAlgorithm));
            }
            else
            {
                var joinColumns = joinOn.GetAs<string>(false).ToArray();
                var merged = left._dataFrame.Merge(right._dataFrame, joinColumns, joinColumns, leftSuffix, rightSuffix, joinAlgorithm);
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

            var merged = left._dataFrame.Merge(right._dataFrame, commonColumns, commonColumns, leftSuffix, rightSuffix, joinAlgorithm);
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

    public XlBlockTable DropNulls(string dropNullBehavior = "all")
    {
        var dropNullOptions = DataFrameUtilities.ParseDropNullOption(dropNullBehavior);
        return new XlBlockTable(_dataFrame.DropNulls(dropNullOptions));
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
            CheckColumnExists(columnName);

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
            throw new ArgumentException("expression must be a valid string");

        var parsed = DataFrameExpressionParser.Parser.Parse(expression);
        if (parsed.IsError)
            throw new Exception(string.Join("|", parsed.Errors.Select(x => x.ErrorMessage)));

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

        CheckColumnExists(columnName);

        var column = _dataFrame[columnName];
        if (!ParamTypeConverter.TryConvertToProvidedType(value, column.DataType, out var converted))
            throw new ArgumentException($"could not convert value '{value}' to type '{column.DataType.Name}'");

        var filter = inclusive ? column.ElementwiseEquals(converted) : column.ElementwiseNotEquals(converted);
        return new XlBlockTable(_dataFrame.Filter(filter));
    }

    public XlBlockTable AppendColumnFromDictionary(XlBlockDictionary dictionary, string keyColumnName, string valueColumnName, string? valueType = null)
    {
        CheckColumnExists(keyColumnName);
        CheckColumnNotExists(valueColumnName);

        var keyColumn = DataFrameUtilities.CreateDataFrameColumn(dictionary.Keys, dictionary.KeyType, keyColumnName);
        DataFrameColumn valueColumn;
        if (valueType != null)
        {
            var columnType = ParamTypeConverter.StringToType(valueType) ?? throw new ArgumentException($"unknown type '{valueType}'");
            var convertedValues = dictionary.Values.ConvertToProvidedType(valueType)
                .Select(x =>
                {
                    if (!x.IsMissingOrError && x.Success)
                        return x.ConvertedInput;

                    return null;
                });
            
            valueColumn = DataFrameUtilities.CreateDataFrameColumn(convertedValues, columnType, valueColumnName);
        }
        else
        {
            var guessConversions = dictionary.Values.ConvertToBestTypes();
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

                return null;
            });
            valueColumn = DataFrameUtilities.CreateDataFrameColumn(convertedValues, determinedType, valueColumnName);

        }
        var rightDataFrame = new DataFrame(keyColumn, valueColumn);

        var joinColumns = new[] { keyColumnName };
        var merged = _dataFrame.Merge(rightDataFrame, joinColumns, joinColumns, "_left", "_right", JoinAlgorithm.Left);
        merged.Columns.Remove($"{keyColumnName}_right");
        merged.Columns[$"{keyColumnName}_left"].SetName(keyColumnName);

        return new XlBlockTable(merged);
    }

    public XlBlockDictionary ToDictionary(string keyColumnName, string valueColumnName)
    {
        CheckColumnExists(keyColumnName);
        CheckColumnExists(valueColumnName);

        var keyColumn = _dataFrame[keyColumnName];
        var valueColumn = _dataFrame[valueColumnName];

        var dict = DictionaryUtilities.BuildTypedDictionary(keyColumn.DataType);
        for (var i = 0L; i < keyColumn.Length; i++)
        {
            if (keyColumn[i] is null)
                continue;

            dict.Add(keyColumn[i], valueColumn[i]);
        }

        return new XlBlockDictionary(dict, keyColumn.DataType);
    }

    public XlBlockDictionary ToDictionaryOfDictionaries(string keyColumnName)
    {
        CheckColumnExists(keyColumnName);

        var keyColumn = _dataFrame[keyColumnName];

        var dict = DictionaryUtilities.BuildTypedDictionary(keyColumn.DataType);
        for (var row = 0L; row < keyColumn.Length; row++)
        {
            if (keyColumn[row] is null)
                continue;

            var rowDict = new Dictionary<string, object>();
            foreach (var column in _dataFrame.Columns)
            {
                if (column.Name == keyColumnName)
                    continue;

                rowDict[column.Name] = column[row];
            }
            dict.Add(keyColumn[row], new XlBlockDictionary(rowDict, typeof(string)));
        }

        return new XlBlockDictionary(dict, keyColumn.DataType);
    }
}
