namespace XlBlocks.AddIn.Tests.Parser;

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
}
