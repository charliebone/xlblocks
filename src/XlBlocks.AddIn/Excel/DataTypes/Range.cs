namespace XlBlocks.AddIn.Excel.DataTypes;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

internal static class DataTypes_Range
{
    [ExcelFunction(Description = "Clean a range: normalize its shape, remove missing values and optionally handle error values", IsThreadSafe = true)]
    public static XlBlockRange XBRange_Clean(
        [ExcelArgument(Description = "A range")] XlBlockRange range,
        [ExcelArgument(Description = "Error handling ('drop', 'keep', or 'error'). Default is 'drop'")] string onErrors = "drop")
    {
        return range.Clean(onErrors);
    }

    [ExcelFunction(Description = "Get the number of cells in a range", IsThreadSafe = true)]
    public static int XBRange_Count(
        [ExcelArgument(Description = "A range")] XlBlockRange range)
    {
        return range.Count;
    }

    [ExcelFunction(Description = "Get the number of rows in a range", IsThreadSafe = true)]
    public static int XBRange_RowCount(
        [ExcelArgument(Description = "A range")] XlBlockRange range)
    {
        return range.RowCount;
    }

    [ExcelFunction(Description = "Get the number of columns in a range", IsThreadSafe = true)]
    public static int XBRange_ColumnCount(
        [ExcelArgument(Description = "A range")] XlBlockRange range)
    {
        return range.ColumnCount;
    }

    [ExcelFunction(Description = "Shape a range", IsThreadSafe = true)]
    public static XlBlockRange XBRange_Shape(
        [ExcelArgument(Description = "A range")] XlBlockRange range,
        [ExcelArgument(Description = "Number of rows in new range")] int? rowCount = null,
        [ExcelArgument(Description = "Number of columns in new range")] int? columnCount = null,
        [ExcelArgument(Description = "Padding value for any added elements (NA)")] object? fillWith = null)
    {
        fillWith ??= ExcelError.ExcelErrorNA;
        return range.Shape(rowCount, columnCount, fillWith);
    }

    [ExcelFunction(Description = "Clean and combine ranges", IsThreadSafe = true)]
    public static XlBlockRange XBRange_Gather(
        [ExcelArgument(Description = "A range")] params XlBlockRange[] range)
    {
        return XlBlockRange.BuildFromMultiple(range, "drop");
    }

    [ExcelFunction(Description = "Reduce a range to only unique values", IsThreadSafe = true)]
    public static XlBlockRange XBRange_GetUniqueValues(
        [ExcelArgument(Description = "A range")] XlBlockRange range)
    {
        return range.GetUniqueValues();
    }

    [ExcelFunction(Description = "Concatenate cells of a range into a single string", IsThreadSafe = true)]
    public static string XBRange_Concatenate(
        [ExcelArgument(Description = "A range")] XlBlockRange range,
        [ExcelArgument(Description = "A delimiter. Default is ','")] string delimiter = ",",
        [ExcelArgument(Description = "A prefix to prepend at the beginning of the returned string. Default is nothing")] string prefix = "",
        [ExcelArgument(Description = "A suffix to append to the end of the returned string. Default is nothing")] string suffix = "")
    {
        return $"{prefix}{range.GetAs<string>(true).Aggregate((x, y) => $"{x}{delimiter}{y}")}{suffix}";
    }

    [ExcelFunction(Description = "SumIf, handling errors", IsThreadSafe = true)]
    public static XlBlockRange XBRange_SumIf(
        [ExcelArgument(Description = "A range")] XlBlockRange range,
        [ExcelArgument(Description = "A range")] XlBlockRange range2)
    {
        throw new NotImplementedException();
    }


}
