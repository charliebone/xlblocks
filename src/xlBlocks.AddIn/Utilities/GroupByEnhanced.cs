namespace XlBlocks.AddIn.Utilities;

using System;
using System.Linq;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Dna;

public interface IGroupByEnhanced
{
    DataFrame Sum(string[] columnNames);
    DataFrame Product(string[] columnNames);
    DataFrame Min(string[] columnNames);
    DataFrame Max(string[] columnNames);
    DataFrame Median(string[] columnNames);
    DataFrame First(string[] columnNames, bool includeNulls);
    DataFrame Last(string[] columnNames, bool includeNulls);
    DataFrame Count(string[] columnNames, bool includeNulls);
    DataFrame Mean(string[] columnNames);
    DataFrame Variance(string[] columnNames, bool isSample);
    DataFrame StdDev(string[] columnNames, bool isSample);
    DataFrame Skew(string[] columnNames, bool isSample);
    DataFrame Kurtosis(string[] columnNames, bool isSample);
}

public class GroupByEnhanced<TKey> : IGroupByEnhanced
{
    private readonly GroupBy<TKey> _groupBy;
    private readonly DataFrame _dataFrame;

    private class MomentAggregate
    {
        private long _n = 0;
        private double _m1 = 0;
        private double _m2 = 0;
        private double _m3 = 0;
        private double _m4 = 0;

        public double? Mean() => _n > 0 ? _m1 : null;

        public double? Variance(bool isSample) => _n > 0 ? _m2 / (isSample ? (_n - 1) : _n) : null;

        public double? StdDev(bool isSample)
        {
            if (_n == 0)
                return null;

            var variance = Variance(isSample);
            return variance is null ? null : Math.Sqrt(variance.Value);
        }

        public double? Skew(bool isSample)
        {
            if (_n == 0)
                return null;

            return _m2 == 0d ? double.NaN : (isSample ? (_n * Math.Sqrt(_n - 1) / (_n - 2)) : Math.Sqrt(_n)) * _m3 / Math.Pow(_m2, 1.5);
        }

        public double? Kurtosis(bool isSample)
        {
            if (_n == 0)
                return null;

            return _m2 == 0d ? double.NaN :
                (isSample ? (_n * (_n + 1) * (_n - 1) / ((_n - 2) * (_n - 3))) : _n) * _m4 / (_m2 * _m2) - (isSample ? (3 * Math.Pow(_n - 1, 2) / ((_n - 2) * (_n - 3))) : 3);
        }

        private void AddValue(double value)
        {
            _n++;
            var delta = value - _m1;
            var delta_n = delta / _n;
            var delta_n2 = Math.Pow(delta_n, 2);
            var term1 = delta * delta_n * (_n - 1);

            _m1 += delta_n;
            _m4 += term1 * delta_n2 * (Math.Pow(_n, 2) - 3 * _n + 3) + 6 * delta_n2 * _m2 - 4 * delta_n * _m3;
            _m3 += term1 * delta_n * (_n - 2) - 3 * delta_n * _m2;
            _m2 += term1;
        }

        private MomentAggregate() { }

        public static MomentAggregate Aggregate(string columnName, IEnumerable<DataFrameRow> rows)
        {
            var aggregate = new MomentAggregate();
            foreach (var row in rows)
            {
                var objValue = row[columnName];
                if (objValue is null)
                    continue;

                if (!ParamTypeConverter.TryConvertTo<double>(objValue, out var doubleValue))
                    throw new Exception($"Column '{columnName}' is not numeric");

                aggregate.AddValue(doubleValue);
            }
            return aggregate;
        }
    }

    public GroupByEnhanced(GroupBy<TKey> groupBy, DataFrame dataFrame)
    {
        _groupBy = groupBy;
        _dataFrame = dataFrame;
    }

    private DataFrame EnumerateAndAggregate<T>(string[] columnNames, Func<string, IEnumerable<DataFrameRow>, T?> statistic)
    {
        var groupingsList = _groupBy.Groupings.Select((x, i) => new { Key = x.Key, Rows = x.AsEnumerable(), Index = i }).ToList();
        var groupKeyColumn = DataFrameUtilities.CreateDataFrameColumn(groupingsList.Select(x => x.Key), "key");
        var dataFrame = new DataFrame(groupKeyColumn);
        foreach (var grouping in groupingsList)
        {
            var rows = grouping.Rows;
            foreach (var columnName in columnNames)
            {
                if (grouping.Index == 0)
                {
                    var columnType = typeof(T) == typeof(object) ? _dataFrame.Columns[columnName].DataType : typeof(T);
                    var underlyingType = Nullable.GetUnderlyingType(columnType);
                    if (underlyingType is not null)
                        columnType = underlyingType;

                    var column = DataFrameUtilities.CreateDataFrameColumn(columnType, columnName, groupingsList.Count);
                    dataFrame.Columns.Add(column);
                }

                dataFrame.SetTypedValue(columnName, grouping.Index, statistic(columnName, rows));
            }
        }
        return dataFrame;
    }

    public DataFrame Sum(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Sum(columnName, rows));
    }

    public DataFrame Product(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Product(columnName, rows));
    }

    public DataFrame Min(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Min(columnName, rows));
    }

    public DataFrame Max(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Max(columnName, rows));
    }

    public DataFrame Median(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Median(columnName, rows));
    }

    public DataFrame First(string[] columnNames, bool includeNulls)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => First(columnName, rows, includeNulls));
    }

    public DataFrame Last(string[] columnNames, bool includeNulls)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Last(columnName, rows, includeNulls));
    }

    public DataFrame Count(string[] columnNames, bool includeNulls)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Count(columnName, rows, includeNulls));
    }

    public DataFrame Mean(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => MomentAggregate.Aggregate(columnName, rows).Mean());
    }

    public DataFrame Variance(string[] columnNames, bool isSample)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => MomentAggregate.Aggregate(columnName, rows).Variance(isSample));
    }

    public DataFrame StdDev(string[] columnNames, bool isSample)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => MomentAggregate.Aggregate(columnName, rows).StdDev(isSample));
    }

    public DataFrame Skew(string[] columnNames, bool isSample)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => MomentAggregate.Aggregate(columnName, rows).Skew(isSample));
    }

    public DataFrame Kurtosis(string[] columnNames, bool isSample)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => MomentAggregate.Aggregate(columnName, rows).Kurtosis(isSample));
    }

    private IEnumerable<double?> GetColumnValuesAsDouble(string columnName, IEnumerable<DataFrameRow> rows, bool includeNulls = false)
    {
        foreach (var row in rows)
        {
            var objValue = row[columnName];
            if (objValue is null)
            {
                if (!includeNulls)
                    continue;
                yield return null;
            }
            else
            {
                if (!ParamTypeConverter.TryConvertTo<double>(objValue, out var doubleValue))
                    throw new Exception($"Column '{columnName}' is not numeric");

                yield return doubleValue;
            }
        }
    }

    private IEnumerable<object> GetColumnValuesAsColumnType(string columnName, IEnumerable<DataFrameRow> rows, bool includeNulls = false)
    {
        var colType = _dataFrame.Columns[columnName].DataType;
        foreach (var row in rows)
        {
            var objValue = row[columnName];
            if (objValue is null)
            {
                if (!includeNulls)
                    continue;
                yield return null!;
            }
            else
            {
                if (!ParamTypeConverter.TryConvertToProvidedType(objValue, colType, out var doubleValue))
                    throw new Exception($"Column '{columnName}' is not numeric");

                yield return doubleValue;
            }
        }
    }

    private double? Sum(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValuesAsDouble(columnName, rows);
        if (!vals.Any())
            return null;

        double? sum = null;
        foreach (var val in vals)
            if (val is not null)
                sum = sum is null ? val : sum + val;
        return sum;
    }

    private double? Product(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValuesAsDouble(columnName, rows);
        if (!vals.Any())
            return null;

        double? product = null;
        foreach (var val in vals)
            if (val is not null)
                product = product is null ? val : product * val;
        return product;
    }

    private double? Median(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValuesAsDouble(columnName, rows).OrderBy(x => x).ToList();
        if (vals.Count == 0)
            return null;

        return vals.Count % 2 == 0 ?
            (vals[vals.Count / 2 - 1] + vals[vals.Count / 2]) / 2d :
            vals[vals.Count / 2];
    }

    private object? Min(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValuesAsColumnType(columnName, rows);
        return vals.Any() ? vals.Min() : null;
    }

    private object? Max(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValuesAsColumnType(columnName, rows);
        return vals.Any() ? vals.Max() : null;
    }

    private object? First(string columnName, IEnumerable<DataFrameRow> rows, bool includeNulls)
    {
        var vals = GetColumnValuesAsColumnType(columnName, rows, includeNulls);
        return vals.FirstOrDefault();
    }

    private object? Last(string columnName, IEnumerable<DataFrameRow> rows, bool includeNulls)
    {
        var vals = GetColumnValuesAsColumnType(columnName, rows, includeNulls);
        return vals.LastOrDefault();
    }

    private int? Count(string columnName, IEnumerable<DataFrameRow> rows, bool includeNulls)
    {
        var vals = GetColumnValuesAsColumnType(columnName, rows, includeNulls);
        return vals.Count();
    }
}
