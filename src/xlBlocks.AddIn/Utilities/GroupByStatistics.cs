namespace XlBlocks.AddIn.Utilities;

using System;
using System.Linq;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Dna;

public interface IGroupByStatistics
{
    DataFrame Mean(string[] columnNames);
    DataFrame Variance(string[] columnNames, bool isSample);
    DataFrame StdDev(string[] columnNames, bool isSample);
    DataFrame Skew(string[] columnNames, bool isSample);
    DataFrame Kurtosis(string[] columnNames, bool isSample);
    DataFrame Median(string[] columnNames);
}

public class GroupByStatistics<TKey> : IGroupByStatistics
{
    private readonly GroupBy<TKey> _groupBy;

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

    public GroupByStatistics(GroupBy<TKey> groupBy)
    {
        _groupBy = groupBy;
    }

    private DataFrame EnumerateAndAggregate(string[] columnNames, Func<string, IEnumerable<DataFrameRow>, double?> statistic)
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
                    var column = new DoubleDataFrameColumn(columnName, groupingsList.Count);
                    dataFrame.Columns.Add(column);
                }

                dataFrame.Columns[columnName][grouping.Index] = statistic(columnName, rows);
            }
        }
        return dataFrame;
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

    public DataFrame Median(string[] columnNames)
    {
        return EnumerateAndAggregate(columnNames, (columnName, rows) => Median(columnName, rows));
    }

    private static IEnumerable<double> GetColumnValues(string columnName, IEnumerable<DataFrameRow> rows)
    {
        foreach (var row in rows)
        {
            var objValue = row[columnName];
            if (objValue is null)
                continue;

            if (!ParamTypeConverter.TryConvertTo<double>(objValue, out var doubleValue))
                throw new Exception($"Column '{columnName}' is not numeric");

            yield return doubleValue;
        }
    }

    private static double? Median(string columnName, IEnumerable<DataFrameRow> rows)
    {
        var vals = GetColumnValues(columnName, rows).OrderBy(x => x).ToList();
        if (vals.Count == 0)
            return null;

        return vals.Count % 2 == 0 ?
            (vals[vals.Count / 2 - 1] + vals[vals.Count / 2]) / 2d :
            vals[vals.Count / 2];
    }
}
