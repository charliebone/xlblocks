namespace XlBlocks.AddIn.Tests.Types;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

public class XlBlockRangeTests
{

    #region Build tests
    [Fact]
    public void Build_Success()
    {
        object[,] array =
        {
            { "a1", "b1", "c1" },
            { "a2", "b2", "c2" },
            { "a3", "b3", "c3" }
        };

        var range = XlBlockRange.Build(array);

        Assert.NotNull(range);
        Assert.Equal(3, range.RowCount);
        Assert.Equal(3, range.ColumnCount);
        Assert.Equal(9, range.Count);
    }

    [Fact]
    public void NonSquareArray_RowAndColumnCount_Success()
    {
        object[,] array =
        {
            { "a1", "b1" },
            { "a2", "b2" },
            { "a3", "b3" },
            { "a4", "b4" }
        };

        var range = XlBlockRange.Build(array);

        Assert.Equal(4, range.RowCount);
        Assert.Equal(2, range.ColumnCount);
        Assert.Equal(8, range.Count);
    }

    [Fact]
    public void BuildFromMultiple_Success_KeepErrors()
    {
        object[,] array1 =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        object[,] array2 =
        {
            { 4, 5, 6 },
            { ExcelEmpty.Value, 8, ExcelError.ExcelErrorNA }
        };

        var range1 = XlBlockRange.Build(array1);
        var range2 = XlBlockRange.Build(array2);
        var ranges = new[] { range1, range2 };

        var combined = XlBlockRange.BuildFromMultiple(ranges, "keep");

        Assert.Equal(9, combined.Count);
        Assert.Equal(1, combined[0, 0]);
        Assert.Equal(2, combined[1, 0]);
        Assert.Equal(3, combined[2, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, combined[3, 0]);
        Assert.Equal(4, combined[4, 0]);
        Assert.Equal(5, combined[5, 0]);
        Assert.Equal(6, combined[6, 0]);
        Assert.Equal(8, combined[7, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, combined[8, 0]);
    }

    [Fact]
    public void BuildFromMultiple_Success_ThrowOnErrors()
    {
        object[,] array1 =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        object[,] array2 =
        {
            { 4, 5, 6 },
            { ExcelEmpty.Value, 8, ExcelError.ExcelErrorNA }
        };

        var range1 = XlBlockRange.Build(array1);
        var range2 = XlBlockRange.Build(array2);
        var ranges = new[] { range1, range2 };

        Assert.Throws<ArgumentException>(() => XlBlockRange.BuildFromMultiple(ranges, "error"));
    }

    [Fact]
    public void BuildFromMultiple_Success_DropErrors()
    {
        object[,] array1 =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        object[,] array2 =
        {
            { 4, 5, 6 },
            { ExcelEmpty.Value, 8, ExcelError.ExcelErrorNA }
        };

        var range1 = XlBlockRange.Build(array1);
        var range2 = XlBlockRange.Build(array2);
        var ranges = new[] { range1, range2 };

        var combined = XlBlockRange.BuildFromMultiple(ranges, "drop");

        Assert.Equal(7, combined.Count);
        Assert.Equal(1, combined[0, 0]);
        Assert.Equal(2, combined[1, 0]);
        Assert.Equal(3, combined[2, 0]);
        Assert.Equal(4, combined[3, 0]);
        Assert.Equal(5, combined[4, 0]);
        Assert.Equal(6, combined[5, 0]);
        Assert.Equal(8, combined[6, 0]);
    }

    [Fact]
    public void BuildFromString_DefaultParameters()
    {
        string str = " a, b , c, , d ";
        string delimiter = ",";

        var range = XlBlockRange.BuildFromString(str, delimiter);

        Assert.Equal(5, range.RowCount);
        Assert.Equal(1, range.ColumnCount);
        Assert.Equal(" a", range[0, 0]);
        Assert.Equal(" b ", range[1, 0]);
        Assert.Equal(" c", range[2, 0]);
        Assert.Equal(" ", range[3, 0]);
        Assert.Equal(" d ", range[4, 0]);
    }

    [Fact]
    public void BuildFromString_TrimStrings()
    {
        string str = " a, b , c, , d ";
        string delimiter = ",";

        var range = XlBlockRange.BuildFromString(str, delimiter, true);

        Assert.Equal(4, range.RowCount);
        Assert.Equal(1, range.ColumnCount);
        Assert.Equal("a", range[0, 0]);
        Assert.Equal("b", range[1, 0]);
        Assert.Equal("c", range[2, 0]);
        Assert.Equal("d", range[3, 0]);
    }

    [Fact]
    public void BuildFromString_IgnoreEmpty()
    {
        string str = " a, b , c, , d ";
        string delimiter = ",";

        var range = XlBlockRange.BuildFromString(str, delimiter, true, true);

        Assert.Equal(4, range.RowCount);
        Assert.Equal(1, range.ColumnCount);
        Assert.Equal("a", range[0, 0]);
        Assert.Equal("b", range[1, 0]);
        Assert.Equal("c", range[2, 0]);
        Assert.Equal("d", range[3, 0]);
    }

    [Fact]
    public void BuildFromString_CustomDelimiter()
    {
        string str = "a|b|c|d";
        string delimiter = "|";

        var range = XlBlockRange.BuildFromString(str, delimiter);

        Assert.Equal(4, range.RowCount);
        Assert.Equal(1, range.ColumnCount);
        Assert.Equal("a", range[0, 0]);
        Assert.Equal("b", range[1, 0]);
        Assert.Equal("c", range[2, 0]);
        Assert.Equal("d", range[3, 0]);
    }

    #endregion

    #region Row and column tests

    [Fact]
    public void GetRow_Success()
    {
        object[,] array =
        {
            { "a1", "b1", "c1" },
            { "a2", "b2", "c2" },
            { "a3", "b3", "c3" }
        };

        var range = XlBlockRange.Build(array);
        var row = range.GetRow(1).ToArray();

        Assert.Equal(3, row.Length);
        Assert.Equal("a2", row[0]);
        Assert.Equal("b2", row[1]);
        Assert.Equal("c2", row[2]);
    }

    [Fact]
    public void GetRow_AsInt_Success()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { "4", "5", "6" },
            { 7, 8, 9 }
        };

        var range = XlBlockRange.Build(array);
        var row = range.GetRowAs<int>(1, false).ToArray();

        Assert.Equal(3, row.Length);
        Assert.Equal(4, row[0]);
        Assert.Equal(5, row[1]);
        Assert.Equal(6, row[2]);
    }

    [Fact]
    public void GetRowAs_ConversionFails_ThrowsException()
    {
        object[,] array =
        {
            { 1, "two", 3 },
            { "invalid", 5, 6 },
            { 7, 8, 9 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.GetRowAs<int>(1, false).ToArray());
    }

    [Fact]
    public void GetRowAs_ConversionFails_SuccessIgnoreErrors()
    {
        object[,] array =
        {
            { 1, "two", 3 },
            { "invalid", 5, 6 },
            { 7, 8, 9 }
        };

        var range = XlBlockRange.Build(array);
        var row = range.GetRowAs<int>(1, true).ToArray();

        Assert.Equal(2, row.Length);
        Assert.Equal(5, row[0]);
        Assert.Equal(6, row[1]);
    }

    [Fact]
    public void GetRowAs_WithExcelSpecialValues_ThrowsException()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelError.ExcelErrorNA, 5, 6 },
            { ExcelMissing.Value, 8, 9 },
            { ExcelEmpty.Value, 11, 12 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.GetRowAs<int>(1, false).ToArray());
    }

    [Fact]
    public void GetRowAs_WithExcelSpecialValues_SuccessIgnoreErrors()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelError.ExcelErrorNA, 5, 6 },
            { ExcelMissing.Value, 8, 9 },
            { ExcelEmpty.Value, 11, 12 }
        };

        var range = XlBlockRange.Build(array);
        var row = range.GetRowAs<int>(1, true).ToArray();

        Assert.Equal(2, row.Length);
        Assert.Equal(5, row[0]);
        Assert.Equal(6, row[1]);
    }

    [Fact]
    public void GetColumn_Success()
    {
        object[,] array =
        {
            { "a1", "b1", "c1" },
            { "a2", "b2", "c2" },
            { "a3", "b3", "c3" }
        };

        var range = XlBlockRange.Build(array);
        var column = range.GetColumn(1).ToArray();

        Assert.Equal(3, column.Length);
        Assert.Equal("b1", column[0]);
        Assert.Equal("b2", column[1]);
        Assert.Equal("b3", column[2]);
    }

    [Fact]
    public void GetColumn_AsInt_Success()
    {
        object[,] array =
        {
            { 1, "2", 3 },
            { 4, "5", 6 },
            { 7, "8", 9 }
        };

        var range = XlBlockRange.Build(array);
        var column = range.GetColumnAs<int>(1, false).ToArray();

        Assert.Equal(3, column.Length);
        Assert.Equal(2, column[0]);
        Assert.Equal(5, column[1]);
        Assert.Equal(8, column[2]);
    }

    [Fact]
    public void GetColumnAs_ConversionFails_ThrowsException()
    {
        object[,] array =
        {
            { 1, 2, "three" },
            { 4, 5, "invalid" },
            { 7, 8, 9 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.GetColumnAs<int>(2, false).ToArray());
    }

    [Fact]
    public void GetColumnAs_ConversionFails_SuccessIgnoreErrors()
    {
        object[,] array =
        {
            { 1, 2, "three" },
            { 4, 5, "invalid" },
            { 7, 8, 9 }
        };

        var range = XlBlockRange.Build(array);
        var column = range.GetColumnAs<int>(2, true).ToArray();

        Assert.Single(column);
        Assert.Equal(9, column[0]);
    }

    [Fact]
    public void GetColumnAs_WithExcelSpecialValues_ThrowsException()
    {
        object[,] array =
        {
            { 1, ExcelError.ExcelErrorNA, 3 },
            { 4, ExcelMissing.Value, 6 },
            { 7, ExcelEmpty.Value, 9 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.GetColumnAs<int>(1, false).ToArray());
    }

    [Fact]
    public void GetColumnAs_WithExcelSpecialValues_SuccessIgnoreErrors()
    {
        object[,] array =
        {
            { 1, ExcelError.ExcelErrorNA, 3 },
            { 4, ExcelMissing.Value, 6 },
            { 7, ExcelEmpty.Value, 9 }
        };

        var range = XlBlockRange.Build(array);
        var column = range.GetColumnAs<int>(1, true).ToArray();

        Assert.Empty(column);
    }

    #endregion

    #region Clean tests

    [Fact]
    public void Clean_Success_KeepErrors()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        var range = XlBlockRange.Build(array);
        var cleaned = range.Clean(XlBlockRange.CleanBehavior.KeepErrors);

        Assert.Equal(4, cleaned.RowCount);
        Assert.Equal(1, cleaned.ColumnCount);
        Assert.Equal(1, cleaned[0, 0]);
        Assert.Equal(2, cleaned[1, 0]);
        Assert.Equal(3, cleaned[2, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, cleaned[3, 0]);
    }

    [Fact]
    public void Clean_Success_DropErrors()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        var range = XlBlockRange.Build(array);
        var cleaned = range.Clean(XlBlockRange.CleanBehavior.DropErrors);

        Assert.Equal(3, cleaned.RowCount);
        Assert.Equal(1, cleaned.ColumnCount);
        Assert.Equal(1, cleaned[0, 0]);
        Assert.Equal(2, cleaned[1, 0]);
        Assert.Equal(3, cleaned[2, 0]);
    }

    [Fact]
    public void Clean_Success_ThrowOnErrors()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        var range = XlBlockRange.Build(array);
        Assert.Throws<ArgumentException>(() => range.Clean(XlBlockRange.CleanBehavior.ThrowOnErrors));
    }

    [Fact]
    public void Clean_MakeArraySafe_SingleValue()
    {
        object[,] array =
        {
            { ExcelMissing.Value, 2 },
            { ExcelEmpty.Value, ExcelError.ExcelErrorNA }
        };

        var range = XlBlockRange.Build(array);
        var cleaned = range.Clean(XlBlockRange.CleanBehavior.DropErrors);

        Assert.Equal(1, cleaned.RowCount);
        Assert.Equal(1, cleaned.ColumnCount);
        Assert.Equal(2, cleaned[0, 0]);

        var madeSafe = cleaned.MakeSafeForArrayFormulas();

        Assert.Equal(2, madeSafe.RowCount);
        Assert.Equal(1, madeSafe.ColumnCount);
        Assert.Equal(2, madeSafe[0, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, madeSafe[1, 0]);
    }

    [Fact]
    public void Clean_MakeArraySafe_MultipleValues()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { ExcelEmpty.Value, ExcelMissing.Value, ExcelError.ExcelErrorNA }
        };

        var range = XlBlockRange.Build(array);
        var cleaned = range.Clean(XlBlockRange.CleanBehavior.DropErrors);
        var madeSafe = cleaned.MakeSafeForArrayFormulas();

        Assert.Equal(3, madeSafe.RowCount);
        Assert.Equal(1, madeSafe.ColumnCount);
        Assert.Equal(1, madeSafe[0, 0]);
        Assert.Equal(2, madeSafe[1, 0]);
        Assert.Equal(3, madeSafe[2, 0]);
    }

    #endregion

    #region Shape tests

    [Fact]
    public void Shape_Success_ChangeShape()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 4, 5, 6 }
        };

        var range = XlBlockRange.Build(array);
        var reshaped = range.Shape(3, 2, 0); // New shape: 3 rows, 2 columns, fill with 0

        Assert.Equal(3, reshaped.RowCount);
        Assert.Equal(2, reshaped.ColumnCount);
        Assert.Equal(1, reshaped[0, 0]);
        Assert.Equal(2, reshaped[0, 1]);
        Assert.Equal(3, reshaped[1, 0]);
        Assert.Equal(4, reshaped[1, 1]);
        Assert.Equal(5, reshaped[2, 0]);
        Assert.Equal(6, reshaped[2, 1]);
    }

    [Fact]
    public void Shape_ThrowsArgumentException_TooFewCells()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 4, 5, 6 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.Shape(2, 2, 0));  // 2x2 < 6 cells, should throw exception
    }

    [Fact]
    public void Shape_ThrowsArgumentException_InvalidRowCount()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 4, 5, 6 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.Shape(0, 2, 0));  // Row count < 1, should throw exception
    }

    [Fact]
    public void Shape_ThrowsArgumentException_InvalidColumnCount()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 4, 5, 6 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<ArgumentException>(() => range.Shape(2, 0, 0));  // Column count < 1, should throw exception
    }

    [Fact]
    public void Shape_Success_FillWithValue()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 4, 5, 6 }
        };

        var range = XlBlockRange.Build(array);
        var reshaped = range.Shape(3, 3, -1); // New shape: 3 rows, 3 columns, fill with -1

        Assert.Equal(3, reshaped.RowCount);
        Assert.Equal(3, reshaped.ColumnCount);
        Assert.Equal(1, reshaped[0, 0]);
        Assert.Equal(2, reshaped[0, 1]);
        Assert.Equal(3, reshaped[0, 2]);
        Assert.Equal(4, reshaped[1, 0]);
        Assert.Equal(5, reshaped[1, 1]);
        Assert.Equal(6, reshaped[1, 2]);
        Assert.Equal(-1, reshaped[2, 0]);  // Filled with -1
        Assert.Equal(-1, reshaped[2, 1]);  // Filled with -1
        Assert.Equal(-1, reshaped[2, 2]);  // Filled with -1
    }

    #endregion

    #region Misc tests
    [Fact]
    public void IsFlat_Property_True()
    {
        object[,] array =
        {
            { 1 },
            { 2 },
            { 3 }
        };

        var range = XlBlockRange.Build(array);

        Assert.True(range.IsFlat);
    }

    [Fact]
    public void IsFlat_Property_False()
    {
        object[,] array =
        {
            { 1, 2 },
            { 3, 4 }
        };

        var range = XlBlockRange.Build(array);

        Assert.False(range.IsFlat);
    }

    [Fact]
    public void IEnumerable_Implementation_Success()
    {
        object[,] array =
        {
            { 1, 2 },
            { 3, 4 },
            { 5, 6 }
        };

        var range = XlBlockRange.Build(array);

        int count = 0;
        foreach (var item in range)
        {
            count++;
        }

        Assert.Equal(6, count);  // There are 6 elements in the 2D array
    }

    [Fact]
    public void IndexOperator_ValidIndices_Success()
    {
        object[,] array =
        {
            { 1, 2 },
            { 3, 4 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Equal(1, range[0, 0]);
        Assert.Equal(2, range[0, 1]);
        Assert.Equal(3, range[1, 0]);
        Assert.Equal(4, range[1, 1]);
    }

    [Fact]
    public void IndexOperator_InvalidIndices_ThrowsException()
    {
        object[,] array =
        {
            { 1, 2 },
            { 3, 4 }
        };

        var range = XlBlockRange.Build(array);

        Assert.Throws<IndexOutOfRangeException>(() => range[2, 0]); // Invalid row index
        Assert.Throws<IndexOutOfRangeException>(() => range[0, 2]); // Invalid column index
        Assert.Throws<IndexOutOfRangeException>(() => range[-1, 0]); // Negative row index
        Assert.Throws<IndexOutOfRangeException>(() => range[0, -1]); // Negative column index
    }

    [Fact]
    public void AsArray_ByColumn_Success_NonSquare()
    {
        object[,] array =
        {
            { "a1", "b1" },
            { "a2", "b2" },
            { "a3", "b3" },
            { "a4", "b4" }
        };

        var range = XlBlockRange.Build(array);
        var result = range.AsArray(RangeOrientation.ByColumn);

        Assert.Equal(4, result.GetLength(0));  // Number of rows
        Assert.Equal(2, result.GetLength(1));  // Number of columns
        Assert.Equal("a1", result[0, 0]);
        Assert.Equal("b1", result[0, 1]);
        Assert.Equal("a2", result[1, 0]);
        Assert.Equal("b2", result[1, 1]);
        Assert.Equal("a3", result[2, 0]);
        Assert.Equal("b3", result[2, 1]);
        Assert.Equal("a4", result[3, 0]);
        Assert.Equal("b4", result[3, 1]);
    }

    [Fact]
    public void AsArray_ByRow_Success_NonSquare()
    {
        object[,] array =
        {
            { "a1", "b1" },
            { "a2", "b2" },
            { "a3", "b3" },
            { "a4", "b4" }
        };

        var range = XlBlockRange.Build(array);
        var result = range.AsArray(RangeOrientation.ByRow);

        Assert.Equal(2, result.GetLength(0));  // Number of rows
        Assert.Equal(4, result.GetLength(1));  // Number of columns
        Assert.Equal("a1", result[0, 0]);
        Assert.Equal("b1", result[0, 1]);
        Assert.Equal("a2", result[0, 2]);
        Assert.Equal("b2", result[0, 3]);
        Assert.Equal("a3", result[1, 0]);
        Assert.Equal("b3", result[1, 1]);
        Assert.Equal("a4", result[1, 2]);
        Assert.Equal("b4", result[1, 3]);
    }

    [Fact]
    public void GetUniqueValues_Success()
    {
        object[,] array =
        {
            { 1, 2, 3 },
            { 1, 4, 2 },
            { 5, 6, 3 }
        };

        var range = XlBlockRange.Build(array);
        var uniqueValues = range.GetUniqueValues();

        Assert.Equal(6, uniqueValues.RowCount);
        Assert.Equal(1, uniqueValues.ColumnCount);
        Assert.Equal(1, uniqueValues[0, 0]);
        Assert.Equal(2, uniqueValues[1, 0]);
        Assert.Equal(3, uniqueValues[2, 0]);
        Assert.Equal(4, uniqueValues[3, 0]);
        Assert.Equal(5, uniqueValues[4, 0]);
        Assert.Equal(6, uniqueValues[5, 0]);
    }

    [Fact]
    public void GetUniqueValues_WithSpecialValues()
    {
        object[,] array =
        {
            { 1, ExcelError.ExcelErrorNA, 3 },
            { 1, ExcelEmpty.Value, ExcelMissing.Value },
            { ExcelError.ExcelErrorNA, 6, 3 }
        };

        var range = XlBlockRange.Build(array);
        var uniqueValues = range.GetUniqueValues();

        Assert.Equal(3, uniqueValues.RowCount);
        Assert.Equal(1, uniqueValues.ColumnCount);
        Assert.Equal(1, uniqueValues[0, 0]);
        Assert.Equal(3, uniqueValues[1, 0]);
        Assert.Equal(6, uniqueValues[2, 0]);
    }

    #endregion
}
