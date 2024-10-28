namespace XlBlocks.AddIn.Parser.Expressions;

using System.Collections.Generic;
using Microsoft.Data.Analysis;
using sly.parser.parser;
using XlBlocks.AddIn.Parser;

internal sealed class ExpressionListExpression : IColumnExpression
{
    public List<IColumnExpression> Expressions { get; }

    public ExpressionListExpression(IColumnExpression firstExpression, List<Group<DataFrameExpressionToken, IColumnExpression>> restExpressions)
    {
        Expressions = new List<IColumnExpression>();
        if (firstExpression != null)
        {
            Expressions.Add(firstExpression);
            if (restExpressions.Count > 0)
            {
                foreach (var expressionGroup in restExpressions)
                    Expressions.Add(expressionGroup.Items[0].Value);
            }
        }
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        throw new NotImplementedException();
    }
}
