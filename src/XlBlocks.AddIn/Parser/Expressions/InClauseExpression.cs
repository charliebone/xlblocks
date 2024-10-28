namespace XlBlocks.AddIn.Parser.Expressions;

using System.Collections.Generic;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class InClauseExpression : IColumnExpression
{
    private readonly IColumnExpression _expression;
    private readonly List<IColumnExpression> _inElements;

    public InClauseExpression(IColumnExpression expression, IColumnExpression inList)
    {
        _expression = expression;
        if (inList is ExpressionListExpression expressionListExpression)
            _inElements = expressionListExpression.Expressions;
        else
            _inElements = new List<IColumnExpression>() { inList };
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        var column = _expression.Evaluate(context);
        var args = _inElements.Select(x => x.Evaluate(context)).ToList();

        if (args.Count < 1)
            throw new DataFrameExpressionException("IN clause requires at least one test expression");

        var result = args.Select(x => column.ElementwiseEquals(x))
            .Cast<DataFrameColumn>()
            .Aggregate((x, y) => x.OrSafe(y));

        return result;
    }
}

