namespace XlBlocks.AddIn.Tests.Parser;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Utilities;
using Xunit;

public class DataFrameColumnExtensionsTests
{
    DataFrameColumn? result, expected;

    [Theory]
    [InlineData(@"test", @"test")]
    [InlineData(@"another_string", @"another.string")]
    [InlineData(@"wild%card", @"wild.*card")]
    [InlineData(@"escaped\_char", @"escaped_char")]
    [InlineData(@"$#^_%", @"\$\#\^..*")]
    [InlineData(@"A mix%of \_escaped \and $ special.chars", @"A\ mix.*of\ _escaped\ \\and\ \$\ special\.chars")]
    public void EscapeLikePattern_Works(string input, string expected)
    {
        Assert.Equal(expected, DataFrameColumnExtensions.EscapeLikePattern(input));
    }

    [Fact]
    public void Elementwise_Like_Tests()
    {
        var _testData = DataFrameUtilities.ToDataFrame(
            new object[,]
            {
                { "Id", "Name", "Age", "Nickname", "Email" },
                { 1, "Alice", 30, "Liz", "Liz_O@notarealdomain.org" },
                { 2, "Bob", 25, null!, "Bob$t*r@somerandom.tld" },
                { 3, "Charlie", 35, "Chuck", "percent%test@notarealdomain.org" }
            },
            new[] { typeof(int), typeof(string), typeof(int), typeof(string), typeof(string) });

        result = _testData["Name"].ElementwiseLike("Bob");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, true, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("Sally");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("Al%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("al%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("al%", true);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("%li%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, false, true });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("%li_");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, true });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseLike("li");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike("%$%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, true, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike("%$_*%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, true, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike("%.org");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, false, true });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike(@"\%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike(@"%\%%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, true });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Email"].ElementwiseLike(@"Liz\_O%");
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }


}
