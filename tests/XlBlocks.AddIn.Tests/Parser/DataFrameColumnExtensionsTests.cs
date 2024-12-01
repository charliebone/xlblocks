namespace XlBlocks.AddIn.Tests.Parser;

using ExcelDna.Integration;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Utilities;
using Xunit;

public class DataFrameColumnExtensionsTests
{

    DataFrameColumn? result, expected;

    private static readonly DataFrame _testData = DataFrameUtilities.ToDataFrame(
            new object[,]
            {
                { "Id", "Name", "Age", "Nickname", "Email", "Comment", "Score" },
                { 1, "Alice", 30, "Liz", "Liz_O@notarealdomain.org", "This is a comment ", 78.92 },
                { 2, "Bob", 25, null!, "Bob$t*r@somerandom.tld", " I have a leading space", 27.34 },
                { 3, "Charlie", 35, "Chuck", "percent%test@notarealdomain.org", null!, 98.765 }
            },
            new[] { typeof(int), typeof(string), typeof(int), typeof(string), typeof(string), typeof(string), typeof(double) });

    private static DataFrameColumn ConstantColumn<T>(T obj) => DataFrameUtilities.CreateConstantDataFrameColumn(obj, _testData.Rows.Count);

    private static readonly DataFrame _logDataTable = DataFrameUtilities.ToDataFrame(
        new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 0, "Trace", ExcelError.ExcelErrorNA, 38.83 },
            { 1, "Warning", 25, ExcelError.ExcelErrorNA },
            { 2, ExcelError.ExcelErrorNA, 21, 83.45 },
            { 3, "Critical", 2, 1.77 },
            { 4, "Debug", 62, 53.67 },
            { 5, ExcelError.ExcelErrorNA, 22, ExcelError.ExcelErrorNA },
            { 6, "Warning", 11, 33.32 },
            { 7, "Critical", 45, 0.82 },
            { 8, "Info", 0, ExcelError.ExcelErrorNA },
            { 9, "Info", 101, ExcelError.ExcelErrorNA },
            { 10, "Debug", 62, 6.34 },
            { 11, "Warning", 45, ExcelError.ExcelErrorNA }
        },
        new[] { typeof(int), typeof(string), typeof(int), typeof(double) });

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
    public void Elementwise_BinaryOps()
    {
        // check that https://github.com/dotnet/machinelearning/issues/7091 is resolved
        var dfTest = new DataFrame(
            new BooleanDataFrameColumn("col1", new[] { true, false, true, false }),
            new BooleanDataFrameColumn("col2", new[] { false, true, true, false }));
        result = dfTest["col1"].And(dfTest["col2"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { false, false, true, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        dfTest = new DataFrame(
            new BooleanDataFrameColumn("col1", new[] { true, false, true, false }),
            new BooleanDataFrameColumn("col2", new[] { false, true, true, false }));
        result = dfTest["col1"].Or(dfTest["col2"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, true, true, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        dfTest = new DataFrame(
            new BooleanDataFrameColumn("col1", new[] { true, false, true, false }),
            new BooleanDataFrameColumn("col2", new[] { false, true, true, false }));
        result = dfTest["col1"].Xor(dfTest["col2"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { true, true, false, false });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Like_Tests()
    {
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

    [Fact]
    public void Elementwise_Len_Tests()
    {
        result = _testData["Name"].ElementwiseLength();
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { 5, 3, 7 }, typeof(int));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseLength();
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { 3, null!, 5 }, typeof(int));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_ExpLog_Tests()
    {
        result = _testData["Age"].ElementwiseExponent(DataFrameUtilities.CreateConstantDataFrameColumn(2, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { Math.Pow(30, 2), Math.Pow(25, 2), Math.Pow(35, 2) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Age"].ElementwiseLog();
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { Math.Log(30), Math.Log(25), Math.Log(35) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Age"].ElementwiseLog(DataFrameUtilities.CreateConstantDataFrameColumn(10, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { Math.Log(30, 10), Math.Log(25, 10), Math.Log(35, 10) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Round_Tests()
    {
        result = _testData["Score"].ElementwiseRound();
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { Math.Round(78.92), Math.Round(27.34), Math.Round(98.765) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Score"].ElementwiseRound(DataFrameUtilities.CreateConstantDataFrameColumn(1, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { Math.Round(78.92, 1), Math.Round(27.34, 1), Math.Round(98.765, 1) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Score"].ElementwiseRound(DataFrameUtilities.CreateConstantDataFrameColumn(-1, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { 80, 30, 100 });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Substring_Tests()
    {
        result = _testData["Nickname"].ElementwiseSubstring(ConstantColumn(2));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "z", null!, "uck" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseSubstring(ConstantColumn(4), ConstantColumn(2));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "", null!, "k" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseSubstring(ConstantColumn(0), ConstantColumn(2));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Li", null!, "Ch" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Left_Right_Tests()
    {
        result = _testData["Nickname"].ElementwiseLeft(ConstantColumn(2));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Li", null!, "Ch" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseLeft(ConstantColumn(4));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Liz", null!, "Chuc" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseRight(ConstantColumn(2));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "iz", null!, "ck" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseRight(ConstantColumn(4));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Liz", null!, "huck" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Trim_Tests()
    {
        result = _testData["Nickname"].ElementwiseTrim();
        expected = _testData["Nickname"];
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Comment"].ElementwiseTrim();
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "This is a comment", "I have a leading space", null! });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Replace_Tests()
    {
        result = _testData["Name"].ElementwiseReplace(
            DataFrameUtilities.CreateConstantDataFrameColumn("Al", _testData.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn("Bc", _testData.Rows.Count),
            null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Bcice", "Bob", "Charlie" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseReplace(
            DataFrameUtilities.CreateConstantDataFrameColumn("al", _testData.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn("Bc", _testData.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn(false, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Alice", "Bob", "Charlie" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseReplace(
            DataFrameUtilities.CreateConstantDataFrameColumn("li", _testData.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn("asdf", _testData.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn(true, _testData.Rows.Count));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "Aasdfce", "Bob", "Charasdfe" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Regex_Test_Tests()
    {
        result = _testData["Nickname"].ElementwiseRegexTest(ConstantColumn("i"), null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { true, null!, false }, typeof(bool));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseRegexTest(ConstantColumn("l"), null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { false, null!, false }, typeof(bool));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Nickname"].ElementwiseRegexTest(ConstantColumn("l"), ConstantColumn(false));
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { true, null!, false }, typeof(bool));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Regex_Find_Tests()
    {
        result = _testData["Name"].ElementwiseRegexFind(ConstantColumn("al"), null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { null!, null!, null! }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseRegexFind(ConstantColumn("al"), ConstantColumn(false));
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "Al", null!, null! }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseRegexFind(ConstantColumn("[Aa]l"), null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "Al", null!, null! }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseRegexFind(ConstantColumn("\\w+e$"), null);
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "Alice", null!, "Charlie" }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Regex_Replace_Tests()
    {
        result = _testData["Name"].ElementwiseRegexReplace(ConstantColumn("al"), ConstantColumn("_"));
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "Alice", "Bob", "Charlie" }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseRegexReplace(ConstantColumn("[Aa]l"), ConstantColumn("_"));
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "_ice", "Bob", "Charlie" }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = _testData["Name"].ElementwiseRegexReplace(ConstantColumn("\\w+e$"), ConstantColumn("ends in e"));
        expected = DataFrameUtilities.CreateDataFrameColumn(new object[] { "ends in e", "Bob", "ends in e" }, typeof(string));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_Format_Tests()
    {
        result = _testData["Age"].ElementwiseFormat(ConstantColumn("0.00"));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "30.00", "25.00", "35.00" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = DataFrameUtilities.CreateDataFrameColumn(new[] { new DateTime(2024, 1, 23), new DateTime(2024, 12, 25), new DateTime(2024, 6, 17) })
            .ElementwiseFormat(DataFrameUtilities.CreateDataFrameColumn("yyyy-MM-dd", result.Length));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { "2024-01-23", "2024-12-25", "2024-06-17" });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Elementwise_ToDateTime_Tests()
    {
        result = DataFrameUtilities.CreateDataFrameColumn(new[] { "2024-01-23", "2024-12-25", "2024-06-17" })
            .ElementwiseToDateTime(DataFrameUtilities.CreateDataFrameColumn("yyyy-MM-dd", 3));
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { new DateTime(2024, 1, 23), new DateTime(2024, 12, 25), new DateTime(2024, 6, 17) });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void CumulativeSumIf_Tests()
    {
        result = _logDataTable.Columns["Id"].CumulativeSumIf(_logDataTable.Columns["Category"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { 0d, 1d, 2d, 3d, 4d, 7d, 7d, 10d, 8d, 17d, 14d, 18d });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void CumulativeProductIf_Tests()
    {
        result = _logDataTable.Columns["Id"].CumulativeProductIf(_logDataTable.Columns["Category"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { 0d, 1d, 2d, 3d, 4d, 10d, 6d, 21d, 8d, 72d, 40d, 66d });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void CumulativeMinIf_Tests()
    {
        result = _logDataTable.Columns["Id"].CumulativeMinIf(_logDataTable.Columns["Category"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { 0d, 1d, 2d, 3d, 4d, 2d, 1d, 3d, 8d, 8d, 4d, 1d });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void CumulativeMaxIf_Tests()
    {
        result = _logDataTable.Columns["Id"].CumulativeMaxIf(_logDataTable.Columns["Category"]);
        expected = DataFrameUtilities.CreateDataFrameColumn(new[] { 0d, 1d, 2d, 3d, 4d, 5d, 6d, 7d, 8d, 9d, 10d, 11d });
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }
}
