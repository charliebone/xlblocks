namespace XlBlocks.AddIn.Tests.Parser;

using System;
using Microsoft.Data.Analysis;
using sly.parser;
using sly.parser.generator;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Parser.Expressions;
using XlBlocks.AddIn.Utilities;
using Xunit.Abstractions;

public class ParserTests
{
    private readonly Parser<DataFrameExpressionToken, IColumnExpression> _parser;
    private readonly ITestOutputHelper _outputHelper;

    DataFrameColumn? result, expected;

    public ParserTests(ITestOutputHelper outputHelper)
    {
        _outputHelper = outputHelper;
        _parser = GetParser(outputHelper);
    }

    private static readonly DataFrame _testData1 = DataFrameUtilities.ToDataFrame(
        new object[,]
        {
            { "Id", "Name", "Age", "Nickname" },
            { 1, "Alice", 30, "Liz" },
            { 2, "Bob", 25, null! },
            { 3, "Charlie", 35, "Chuck" }
        },
        new[] { typeof(int), typeof(string), typeof(int), typeof(string) });

    private static readonly DataFrame _testData2 = DataFrameUtilities.ToDataFrame(
        new object[,]
        {
            { "Id", "Name", "Age", "Nickname" },
            { 1, "Alice", 30, "Liz" },
            { 2, "Bob", 25, "" },
            { 3, "Charlie", 35, "Chuck" }
        },
        new[] { typeof(int), typeof(string), typeof(int), typeof(string) });

    #region Helpers

    private static Parser<DataFrameExpressionToken, IColumnExpression> GetParser(ITestOutputHelper outputHelper)
    {
        var parserInstance = new DataFrameExpressionParser();
        var builder = new ParserBuilder<DataFrameExpressionToken, IColumnExpression>();
        var parser = builder.BuildParser(parserInstance, ParserType.EBNF_LL_RECURSIVE_DESCENT, $"{nameof(DataFrameExpressionParser)}_expressions");
        if (parser.IsError)
        {
            foreach (var error in parser.Errors)
                outputHelper.WriteLine(error.Message);
        }

        Assert.True(parser.IsOk);
        return parser.Result;
    }

    private DataFrameColumn? ParseWithDataFrame(string expr, DataFrame dataFrame)
    {
        var result = _parser.Parse(expr);
        if (result.IsError)
        {
            foreach (var error in result.Errors)
                _outputHelper.WriteLine(error.ErrorMessage);
        }
        Assert.True(result.IsOk);

        var context = new DataFrameContext(dataFrame);
        return result.Result.Evaluate(context);
    }

    #endregion

    #region Operator tests

    [Fact]
    public void MathBinaryOp_ColumnAndColumn_Success()
    {
        result = ParseWithDataFrame("[Id] + [Age]", _testData1);
        expected = _testData1.Columns["Id"].Add(_testData1.Columns["Age"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] - [Id]", _testData1);
        expected = _testData1.Columns["Age"].Subtract(_testData1.Columns["Id"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] * [Age]", _testData1);
        expected = _testData1.Columns["Id"].Multiply(_testData1.Columns["Age"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] / [Id]", _testData1);
        expected = _testData1.Columns["Age"].Divide(_testData1.Columns["Id"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void MathBinaryOp_ColumnAndConstant_Success()
    {
        result = ParseWithDataFrame("[Id] + 1", _testData1);
        expected = _testData1.Columns["Id"].Add(1);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] - 5.0", _testData1);
        expected = _testData1.Columns["Age"].Subtract(5d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] * 5", _testData1);
        expected = _testData1.Columns["Id"].Multiply(5);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] / 5.0", _testData1);
        expected = _testData1.Columns["Age"].Divide(5d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void MathBinaryOp_OrderOfOperations()
    {
        result = ParseWithDataFrame("([Id] * 5) - 2", _testData1);
        expected = _testData1.Columns["Id"].Multiply(5).Subtract(2);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("([Id] - 2) * 5.0", _testData1);
        expected = _testData1.Columns["Id"].Subtract(2).Multiply(5d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("(([Id]) - 2) + ((5.0 - [Id]) - 7.0)", _testData1);
        expected = _testData1.Columns["Id"].Subtract(2).Add(_testData1.Columns["Id"].ReverseSubtract(5d).Subtract(7d));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] * [Age] - 3", _testData1);
        expected = _testData1.Columns["Id"].Multiply(_testData1.Columns["Age"]).Subtract(3);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] - [Id] * 5.0", _testData1);
        expected = _testData1.Columns["Age"].Subtract(_testData1.Columns["Id"].Multiply(5d));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("([Age] - [Id]) * 5.0", _testData1);
        expected = _testData1.Columns["Age"].Subtract(_testData1.Columns["Id"]).Multiply(5d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void LogicalBinaryOp_ColumnAndColumnComparisons_Success()
    {
        result = ParseWithDataFrame("20 * [Id] < [Age]", _testData1);
        expected = _testData1.Columns["Id"].Multiply(20).ElementwiseLessThan(_testData1.Columns["Age"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] + 1 <= 10.0 * [Age]", _testData1);
        expected = _testData1.Columns["Id"].Add(1).ElementwiseLessThanOrEqual(_testData1.Columns["Age"].Multiply(10d));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("30 / [Id] > [Age] + 2", _testData1);
        expected = _testData1.Columns["Id"].ReverseDivide(30).ElementwiseGreaterThan(_testData1.Columns["Age"].Add(2));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] - 3 >= 10.5 * [Id]", _testData1);
        expected = _testData1.Columns["Age"].Subtract(3).ElementwiseGreaterThanOrEqual(_testData1.Columns["Id"].Multiply(10.5));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] == [Age] / 3", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseEquals(_testData1.Columns["Age"].Divide(3));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] * 0.5 <> 7.5 * [Id]", _testData1);
        expected = _testData1.Columns["Age"].Multiply(0.5).ElementwiseNotEquals(_testData1.Columns["Id"].Multiply(7.5));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void LogicalBinaryOp_ColumnAndConstantComparisons_Success()
    {
        result = ParseWithDataFrame("[Id] < 2", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseLessThan(2);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] <= 2.0", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseLessThanOrEqual(2d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] > 30", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseGreaterThan(30);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] >= 30.000", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseGreaterThanOrEqual(30d);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] == 2", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseEquals(2);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Id] <> 2", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseNotEquals(2);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void UnaryOperators_Negation()
    {
        result = ParseWithDataFrame("NOT ([Id] <= 2.0)", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseLessThanOrEqual(2d).ElementwiseNotEquals(true);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("! ([Id] <= 2.0)", _testData1);
        expected = _testData1.Columns["Id"].ElementwiseLessThanOrEqual(2d).ElementwiseNotEquals(true);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Age] < 30", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseLessThan(30);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("!([Age] < 30)", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseLessThan(30).ElementwiseNotEquals(true);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void UnaryOperators_NullCheck()
    {
        result = ParseWithDataFrame("[Name] IS NULL", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseIsNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] IS NOT NULL", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] <> NULL", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] != NULL", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Nickname] IS NULL", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Nickname] IS NOT NULL", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Nickname] == NULL", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("NULL == [Nickname]", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Nickname] != NULL", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Nickname] <> NULL", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("NULL != [Nickname]", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("NULL <> [Nickname]", _testData1);
        expected = _testData1.Columns["Nickname"].ElementwiseIsNotNull();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void InClause_Tests()
    {
        result = ParseWithDataFrame("[Name] IN ('Charlie')", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseEquals("Charlie");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] IN ('Nobody')", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseEquals("Nobody");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] IN ('Charlie','Gina','Adam','Alice')", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseEquals("Charlie")
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Gina"))
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Adam"))
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Alice"));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] IN ('Steve','Gina','Adam','Diane')", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseEquals("Steve")
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Gina"))
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Adam"))
            .OrSafe(_testData1.Columns["Name"].ElementwiseEquals("Diane"));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Theory]
    // parser syntax errors
    [InlineData("Id]", "unexpected right bracket (']")]
    [InlineData("bad_string", "unexpected end of stream")]
    [InlineData("NOT NOT TRUE", "unexpected NOT ('NOT")]
    [InlineData("ISNULL([Nickname]", "unexpected end of stream.")]
    public void ParseErrors_InvalidOperations(string expression, params string[] errors)
    {
        var parseResult = _parser.Parse(expression);
        Assert.True(parseResult.IsError);
        Assert.Equal(errors.Length, parseResult.Errors.Count);
        foreach (var (expected, result) in errors.Zip(parseResult.Errors))
            Assert.StartsWith(expected, result.ErrorMessage);
    }

    [Theory]
    // evaluation errors
    [InlineData("NOT [Id]", "'NOT' operator is invalid for column type 'Int32'")]
    [InlineData("[Id] < [Name]", "'<' operator is invalid between columns of type 'Int32' and 'String'")]

    // function call errors
    [InlineData("BAD_FUNCTION([Id])", "Unknown function 'BAD_FUNCTION'")]
    [InlineData("ISNULL([Nickname])", "Error evaluating function 'ISNULL': Expected 2 args, got 1")]
    [InlineData("IIF([Nickname], [Id], [Age])", "Error evaluating function 'IIF': Conditional must be boolean")]
    public void EvaluationErrors_InvalidOperations(string expression, string message)
    {
        var ex = Assert.Throws<DataFrameExpressionException>(() => ParseWithDataFrame(expression, _testData1));
        Assert.Equal(message, ex.Message);
    }

    [Theory]
    // arithmetic, precedence is P, E, MD, mod, AS
    [InlineData("1 * 2 + 3 * 4", "(1 * 2) + (3 * 4)")]
    [InlineData("1 - 2 / 3 - 4", "(1 - (2 / 3)) - 4")]
    [InlineData("1 / 2 + 3 - 4", "((1 / 2) + 3) - 4")]
    [InlineData("1 + 2 * 3 * 4", "1 + ((2 * 3) * 4)")]
    [InlineData("1 * 2 * 3 - 4", "((1 * 2) * 3) - 4")]
    [InlineData("1 + (2 + 3) + (4)", "((1 + 2) + 3) + 4")]
    [InlineData("1 * (2 + 3) * 4", "1 * (2 + 3) * 4")]
    [InlineData("1 * (6 % 4) * 4", "1 * 2 * 4")]
    [InlineData("1 * 6 % 4 * 4", "(1 * 6) % (4 * 4)")]
    [InlineData("1 * 2 ^ 2 - 4", "(1 * (2^2)) - 4")]
    [InlineData("3 ^ 4 - 2 * 3", "(3^4) - (2 * 3)")]

    // left associative and same precedence for addition / subtraction and multiplication / division
    [InlineData("1 - 2 - 3", "(1 - 2) - 3")]
    [InlineData("1 + 2 + 3 + 4", "((1 + 2) + 3) + 4")]
    [InlineData("1 + 2 - 3 + 4", "((1 + 2) - 3) + 4")]
    [InlineData("1 * 2 / 3", "(1 * 2) / 3")]
    [InlineData("1 * 2 / 3 * 4", "((1 * 2) / 3) * 4")]

    // boolean, AND before OR and XOR
    [InlineData("TRUE OR FALSE", "TRUE")]
    [InlineData("TRUE AND FALSE", "FALSE")]
    [InlineData("TRUE OR TRUE AND FALSE", "TRUE")]
    [InlineData("TRUE AND FALSE OR TRUE", "TRUE")]
    [InlineData("TRUE XOR FALSE", "TRUE")]
    [InlineData("TRUE XOR TRUE", "FALSE")]
    [InlineData("NOT TRUE", "FALSE")]
    [InlineData("NOT FALSE", "TRUE")]
    [InlineData("NOT FALSE AND NOT TRUE", "FALSE")]
    [InlineData("NOT FALSE OR NOT TRUE", "TRUE")]
    [InlineData("NOT (NOT FALSE)", "FALSE")]
    [InlineData("NOT FALSE OR NOT (NOT TRUE)", "TRUE")]
    [InlineData("NOT (FALSE OR NOT (NOT TRUE))", "FALSE")]
    [InlineData("TRUE XOR FALSE AND TRUE", "TRUE")]
    [InlineData("TRUE OR FALSE XOR TRUE", "FALSE")]
    [InlineData("TRUE OR (FALSE XOR TRUE)", "TRUE")]
    [InlineData("FALSE OR (FALSE XOR FALSE)", "FALSE")]
    [InlineData("TRUE OR FALSE XOR FALSE OR TRUE", "TRUE")]
    [InlineData("TRUE OR (FALSE XOR FALSE) OR TRUE", "TRUE")]

    // null checks with is and equals
    [InlineData("NULL IS NULL", "TRUE")]
    [InlineData("NULL IS NOT NULL", "FALSE")]
    [InlineData("NOT NULL IS NOT NULL", "TRUE")]
    [InlineData("NOT (NULL IS NULL)", "FALSE")]

    // string concatenation
    [InlineData("1 + '2'", "'12'")]
    [InlineData("'1' + 2", "'12'")]
    [InlineData("1 + '2' + 'Three'", "'12Three'")]
    [InlineData("'One' + '2' + 'Three'", "'One2Three'")]
    [InlineData("'One' + 2 + 'Three'", "'One2Three'")]

    // date comparisons
    [InlineData("'2022-01-01' > '2023-01-01'", "FALSE")]
    [InlineData("'2022-01-01' == '2023-01-01'", "FALSE")]
    [InlineData("'2022-01-01' <= '2023-01-01'", "TRUE")]
    public void OrderOfOperationsAndAssociativity_Literals(string expression, string expected)
    {
        var expressionResult = ParseWithDataFrame(expression, _testData1);
        var expectedResult = ParseWithDataFrame(expected, _testData1);
        DataFrameTestHelpers.AssertDataColumnsEqual(expectedResult, expressionResult);
    }

    [Fact]
    public void LikeClause_Tests()
    {
        result = ParseWithDataFrame("[Name] LIKE 'Bob'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("Bob");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE 'Sally'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("Sally");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE 'Al%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("Al%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE 'al%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("al%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKEI 'al%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("al%", true);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%li%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%li%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%li_'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%li_");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%Li_'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%Li_");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKEI '%Li_'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%Li_", true);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE 'li'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("li");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%$%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%$%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%$_*%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%$_*%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("[Name] LIKE '%.org'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike("%.org");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame(@"[Name] LIKE '\%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike(@"\%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame(@"[Name] LIKE '%\%%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike(@"%\%%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame(@"[Name] LIKE 'Liz\_O%'", _testData1);
        expected = _testData1.Columns["Name"].ElementwiseLike(@"Liz\_O%");
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void StringTests_Concat()
    {
        result = ParseWithDataFrame("[Name] + ' aka ' + [Nickname]", _testData1);
        expected = _testData1.Columns["Name"]
            .ElementwiseConcat(DataFrameUtilities.CreateConstantDataFrameColumn(" aka ", _testData1.Rows.Count))
            .ElementwiseConcat(_testData1.Columns["Nickname"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void StringTests_EmptyLiteral()
    {
        result = ParseWithDataFrame("[Name] + '' + [Nickname]", _testData1);
        expected = _testData1.Columns["Name"]
            .ElementwiseConcat(_testData1.Columns["Nickname"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    #endregion

    #region Function tests

    [Fact]
    public void Functions_ISNULL_Test()
    {
        result = ParseWithDataFrame("ISNULL([Name], 'Default')", _testData1);
        expected = _testData1.Columns["Name"];
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("ISNULL([Nickname], [Name])", _testData1);
        expected = _testData1.Columns["Nickname"].Clone();
        expected[1] = "Bob";
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_IIF_Test()
    {
        result = ParseWithDataFrame("IIF([Age] >= 30, 'Over30', 'Under30')", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseGreaterThanOrEqual(30).ElementwiseIfThenElse(
            DataFrameUtilities.CreateConstantDataFrameColumn("Over30", _testData1.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn("Under30", _testData1.Rows.Count));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_IIF_MixedTypeToString()
    {
        result = ParseWithDataFrame("IIF([Age] >= 30, 'Over30', [Age])", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseGreaterThanOrEqual(30).ElementwiseIfThenElse(
            DataFrameUtilities.CreateConstantDataFrameColumn("Over30", _testData1.Rows.Count),
            _testData1.Columns["Age"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_IIF_WithNullLiteral()
    {
        result = ParseWithDataFrame("IIF([Age] > 30, 'Over30', NULL)", _testData1);
        expected = _testData1.Columns["Age"].ElementwiseGreaterThan(30).ElementwiseIfThenElse(
            DataFrameUtilities.CreateConstantDataFrameColumn("Over30", _testData1.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn<string>(null!, _testData1.Rows.Count));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_LEN_Test()
    {
        result = ParseWithDataFrame("IIF(LEN([Nickname]) == 0, 'No nickname', [Nickname])", _testData1);
        expected = _testData2.Columns["Nickname"].ElementwiseLength().ElementwiseEquals(0).ElementwiseIfThenElse(
            DataFrameUtilities.CreateConstantDataFrameColumn("No nickname", _testData2.Rows.Count),
            _testData2.Columns["Nickname"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_ROUND_Test()
    {
        result = ParseWithDataFrame("ROUND([Age] / 2)", _testData1);
        expected = _testData2.Columns["Age"].Divide(2d).ElementwiseRound();
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("ROUND([Age] / 2, 1)", _testData1);
        expected = _testData2.Columns["Age"].Divide(2d).ElementwiseRound(
            DataFrameUtilities.CreateConstantDataFrameColumn(1, _testData1.Rows.Count));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    [Fact]
    public void Functions_SUBSTRING_Test()
    {
        result = ParseWithDataFrame("SUBSTRING([Name], 5)", _testData1);
        expected = _testData2.Columns["Name"].ElementwiseSubstring(
            DataFrameUtilities.CreateConstantDataFrameColumn(5, _testData1.Rows.Count));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);

        result = ParseWithDataFrame("SUBSTRING([Name], 2, 3)", _testData1);
        expected = _testData2.Columns["Name"].ElementwiseSubstring(
            DataFrameUtilities.CreateConstantDataFrameColumn(2, _testData1.Rows.Count),
            DataFrameUtilities.CreateConstantDataFrameColumn(3, _testData1.Rows.Count));
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }

    #endregion

    #region Misc tests

    [Fact]
    public void Equality_BlankString()
    {
        result = ParseWithDataFrame("IIF([Nickname] == '', 'No nickname', [Nickname])", _testData2);
        expected = _testData2.Columns["Nickname"].ElementwiseEquals(string.Empty).ElementwiseIfThenElse(
            DataFrameUtilities.CreateConstantDataFrameColumn("No nickname", _testData2.Rows.Count),
            _testData2.Columns["Nickname"]);
        DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
    }


    [Fact]
    public void ParserThreadSafety()
    {
        using var threadsReady = new ManualResetEvent(false);
        var nThreads = 40;
        var nWaiting = nThreads;
        var nFailures = 0;

        Exception? exception = null;
        var threads = new Thread[nThreads];

        var parser = GetParser(_outputHelper);
        var expression = "[Age] - 3 >= 10.5 * [Id]";
        var expected = _testData1.Columns["Age"].Subtract(3).ElementwiseGreaterThanOrEqual(_testData1.Columns["Id"].Multiply(10.5));
        for (var j = 0; j < nThreads; j++)
        {
            threads[j] = new Thread(() =>
            {
                try
                {
                    if (Interlocked.Decrement(ref nWaiting) == 0)
                        threadsReady.Set();
                    else
                        threadsReady.WaitOne();

                    // once all threads are ready, then parse with parser
                    var parsed = parser.Parse(expression);
                    if (parsed.IsError)
                    {
                        foreach (var error in parsed.Errors)
                            _outputHelper.WriteLine(error.ErrorMessage);
                    }
                    Assert.True(parsed.IsOk);

                    var context = new DataFrameContext(_testData1);
                    var result = parsed.Result.Evaluate(context);
                    DataFrameTestHelpers.AssertDataColumnsEqual(expected, result);
                }
                catch (Exception ex)
                {
                    Interlocked.CompareExchange(ref exception, ex, null);
                    Interlocked.Increment(ref nFailures);
                }
            });
        }

        // first start all threads
        for (var j = 0; j < nThreads; j++)
            threads[j].Start();

        // once last thread joined they all should hit parser
        for (var j = 0; j < nThreads; j++)
            threads[j].Join();

        Assert.Null(exception);
        Assert.Equal(0, Interlocked.CompareExchange(ref nFailures, 0, 0));
    }

    #endregion
}
