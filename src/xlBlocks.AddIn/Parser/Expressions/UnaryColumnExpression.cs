namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using sly.lexer;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class UnaryColumnExpression : IColumnExpression
{
    private readonly IColumnExpression _expression;
    private readonly Token<DataFrameExpressionToken> _opToken;

    public UnaryColumnExpression(Token<DataFrameExpressionToken> opToken, IColumnExpression expression)
    {
        _expression = expression;
        _opToken = opToken;
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        var operand = _expression.Evaluate(context);

        if (operand is NullDataFrameColumn nullColumn)
        {
            if (_opToken.TokenID == DataFrameExpressionToken.NOT || _opToken.TokenID == DataFrameExpressionToken.EXCLAMATION_POINT)
                return nullColumn.Negate();

            throw new DataFrameExpressionException($"'{_opToken.Value}' operator is invalid for column '{operand.DataType.Name}'", _opToken.TokenID);
        }

        try
        {
            return _opToken.TokenID switch
            {
                DataFrameExpressionToken.NOT => operand.ElementwiseNotEquals(true),
                DataFrameExpressionToken.EXCLAMATION_POINT => operand.ElementwiseNotEquals(true),

                DataFrameExpressionToken.ARITH_MINUS => operand.Multiply(-1),

                _ => throw new DataFrameExpressionException($"no unary operation defined for token {_opToken.Label}"),
            };
        }
        catch (NotSupportedException ex)
        {
            throw new DataFrameExpressionException($"'{_opToken.Value}' operator is invalid for column type '{operand.DataType.Name}'", ex, _opToken.TokenID);
        }
        catch (Exception ex)
        {
            throw new DataFrameExpressionException($"Error parsing operation '{_opToken.Value}' on column of type '{operand.DataType.Name}'", ex, _opToken.TokenID);
        }
    }
}
