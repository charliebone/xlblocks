namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class LikeClauseExpression : IColumnExpression
{
    private readonly IColumnExpression _expression;
    private readonly IColumnExpression _patternExpression;
    private readonly bool _caseInsensitive;

    public LikeClauseExpression(IColumnExpression expression, IColumnExpression patternExpression, bool caseInsensitive)
    {
        _expression = expression;
        _patternExpression = patternExpression;
        _caseInsensitive = caseInsensitive;
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        var expressionColumn = _expression.Evaluate(context);
        var patternColumn = _patternExpression.Evaluate(context);

        if (expressionColumn is not StringDataFrameColumn || patternColumn is not StringDataFrameColumn)
            throw new DataFrameExpressionException("LIKE operator only valid for comparing string columns or literals");

        return expressionColumn.ElementwiseLike(patternColumn, _caseInsensitive);
    }
}

