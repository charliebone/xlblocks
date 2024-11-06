namespace XlBlocks.AddIn.Tests.Parser;

using Microsoft.Data.Analysis;

internal static class DataFrameTestHelpers
{
    public static void AssertDataColumnsEqual(DataFrameColumn? expected, DataFrameColumn? actual)
    {
        Assert.NotNull(expected);
        Assert.NotNull(actual);
        Assert.Equal(expected.Length, actual.Length);

        var equals = expected.ElementwiseEquals(actual);
        for (var i = 0; i < equals.Length; i++)
            Assert.True(equals[i], $"column values differ, expected '{expected[i]}', got '{actual[i]}' at row {i}");
    }
}
