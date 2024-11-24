namespace XlBlocks.AddIn.Utilities;

using System.Collections;
using Microsoft.Data.Analysis;

internal class PseudoDoubleDataFrameColumn : PrimitiveDataFrameColumn<double>
{
    // consider inheriting from DataFrameColumn instead once ml.net v22 issue is resolved: https://github.com/dotnet/machinelearning/issues/7323
    private readonly PrimitiveDataFrameColumn<double>? _doubleColumn;
    private readonly PrimitiveDataFrameColumn<float>? _floatColumn;
    private readonly PrimitiveDataFrameColumn<decimal>? _decimalColumn;
    private readonly PrimitiveDataFrameColumn<ulong>? _ulongColumn;
    private readonly PrimitiveDataFrameColumn<long>? _longColumn;
    private readonly PrimitiveDataFrameColumn<uint>? _uintColumn;
    private readonly PrimitiveDataFrameColumn<int>? _intColumn;
    private readonly PrimitiveDataFrameColumn<ushort>? _ushortColumn;
    private readonly PrimitiveDataFrameColumn<short>? _shortColumn;
    private readonly PrimitiveDataFrameColumn<sbyte>? _sbyteColumn;
    private readonly PrimitiveDataFrameColumn<byte>? _byteColumn;

    private readonly Type _actualType;

    private readonly Func<long> _getNullCountDelegate;
    private readonly Func<long, double?> _getValueDelegate;
    private readonly Func<IEnumerator> _getEnumeratorDelegate;
    private readonly Func<DataFrameColumn, (HashSet<long>, Dictionary<long, ICollection<long>>)> _getGroupedOccurancesDelegate;

    public override long NullCount => _getNullCountDelegate();
    public Type ActualType => _actualType;

    public new double? this[long rowIndex] => _getValueDelegate(rowIndex);

    public PseudoDoubleDataFrameColumn(DataFrameColumn column) : base(column.Name, column.Length)
    {
        _actualType = column.DataType;

        if (column is PrimitiveDataFrameColumn<double> doubleColumn)
        {
            _doubleColumn = doubleColumn;
            _getNullCountDelegate = () => _doubleColumn.NullCount;
            _getValueDelegate = (rowIndex) => _doubleColumn[rowIndex];
            _getEnumeratorDelegate = () => _doubleColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _doubleColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<float> floatColumn)
        {
            _floatColumn = floatColumn;
            _getNullCountDelegate = () => _floatColumn.NullCount;
            _getValueDelegate = (rowIndex) => _floatColumn[rowIndex];
            _getEnumeratorDelegate = () => _floatColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _floatColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<decimal> decimalColumn)
        {
            _decimalColumn = decimalColumn;
            _getNullCountDelegate = () => _decimalColumn.NullCount;
            _getValueDelegate = (rowIndex) =>
            {
                var decVal = _decimalColumn[rowIndex];
                return decVal.HasValue ? (double)decVal.Value : null;
            };
            _getEnumeratorDelegate = () => _decimalColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _decimalColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<ulong> ulongColumn)
        {
            _ulongColumn = ulongColumn;
            _getNullCountDelegate = () => _ulongColumn.NullCount;
            _getValueDelegate = (rowIndex) => _ulongColumn[rowIndex];
            _getEnumeratorDelegate = () => _ulongColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _ulongColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<long> longColumn)
        {
            _longColumn = longColumn;
            _getNullCountDelegate = () => _longColumn.NullCount;
            _getValueDelegate = (rowIndex) => _longColumn[rowIndex];
            _getEnumeratorDelegate = () => _longColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _longColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<uint> uintColumn)
        {
            _uintColumn = uintColumn;
            _getNullCountDelegate = () => _uintColumn.NullCount;
            _getValueDelegate = (rowIndex) => _uintColumn[rowIndex];
            _getEnumeratorDelegate = () => _uintColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _uintColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<int> intColumn)
        {
            _intColumn = intColumn;
            _getNullCountDelegate = () => _intColumn.NullCount;
            _getValueDelegate = (rowIndex) => _intColumn[rowIndex];
            _getEnumeratorDelegate = () => _intColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _intColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<ushort> ushortColumn)
        {
            _ushortColumn = ushortColumn;
            _getNullCountDelegate = () => _ushortColumn.NullCount;
            _getValueDelegate = (rowIndex) => _ushortColumn[rowIndex];
            _getEnumeratorDelegate = () => _ushortColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _ushortColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<short> shortColumn)
        {
            _shortColumn = shortColumn;
            _getNullCountDelegate = () => _shortColumn.NullCount;
            _getValueDelegate = (rowIndex) => _shortColumn[rowIndex];
            _getEnumeratorDelegate = () => _shortColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _shortColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<byte> byteColumn)
        {
            _byteColumn = byteColumn;
            _getNullCountDelegate = () => _byteColumn.NullCount;
            _getValueDelegate = (rowIndex) => _byteColumn[rowIndex];
            _getEnumeratorDelegate = () => _byteColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _byteColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else if (column is PrimitiveDataFrameColumn<sbyte> sbyteColumn)
        {
            _sbyteColumn = sbyteColumn;
            _getNullCountDelegate = () => _sbyteColumn.NullCount;
            _getValueDelegate = (rowIndex) => _sbyteColumn[rowIndex];
            _getEnumeratorDelegate = () => _sbyteColumn.GetEnumerator();
            _getGroupedOccurancesDelegate = (other) =>
            {
                var groupedOccurances = _sbyteColumn.GetGroupedOccurrences(other, out var otherColumnNullIndices);
                return (otherColumnNullIndices, groupedOccurances);
            };
        }
        else
        {
            throw new ArgumentException($"Cannot wrap column of type '{column.DataType.Name}' in a pseudo double column");
        }
    }

    protected override object GetValue(long rowIndex) => _getValueDelegate(rowIndex)!;

    public double? GetDoubleValue(long rowIndex) => _getValueDelegate(rowIndex);

    protected override IReadOnlyList<object> GetValues(long startIndex, int length)
    {
        if (startIndex >= Length)
            throw new ArgumentOutOfRangeException(nameof(startIndex));

        var ret = new List<object>(length);
        long endIndex = Math.Min(Length, startIndex + length);
        for (var i = startIndex; i < endIndex; i++)
        {
            ret.Add(GetValue(i));
        }
        return ret;
    }

    protected override void SetValue(long rowIndex, object value)
    {
        throw new InvalidOperationException("Cannot set values on pseudo double column");
    }

    protected override IEnumerator GetEnumeratorCore() => _getEnumeratorDelegate();

    public override Dictionary<long, ICollection<long>> GetGroupedOccurrences(DataFrameColumn other, out HashSet<long> otherColumnNullIndices)
    {
        var groupedOccurancesResult = _getGroupedOccurancesDelegate(other);
        otherColumnNullIndices = groupedOccurancesResult.Item1;
        return groupedOccurancesResult.Item2;
    }
}

