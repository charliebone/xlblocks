namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using sly.lexer;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class BinaryColumnExpression : IColumnExpression
{
    private readonly IColumnExpression _left;
    private readonly Token<DataFrameExpressionToken> _opToken;
    private readonly IColumnExpression _right;

    public BinaryColumnExpression(IColumnExpression left, Token<DataFrameExpressionToken> opToken, IColumnExpression right)
    {
        _left = left;
        _opToken = opToken;
        _right = right;
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        var left = _left.Evaluate(context);
        var right = _right.Evaluate(context);

        try
        {
            var leftNullColumn = left as NullDataFrameColumn;
            var rightNullColumn = right as NullDataFrameColumn;
            if (leftNullColumn is not null && rightNullColumn is not null)
            {
                if (_opToken.TokenID == DataFrameExpressionToken.IS || _opToken.TokenID == DataFrameExpressionToken.COMP_EQUALS)
                {
                    return new ConstantColumnExpression<bool>(leftNullColumn.IsNull == rightNullColumn.IsNull).Evaluate(context);
                }
                else if (_opToken.TokenID == DataFrameExpressionToken.COMP_NOTEQUALS)
                {
                    return new ConstantColumnExpression<bool>(leftNullColumn.IsNull != rightNullColumn.IsNull).Evaluate(context);
                }
                throw new DataFrameExpressionException($"'{_opToken.Value}' cannot be used to compare two NULLs");
            }

            if (leftNullColumn is not null || rightNullColumn is not null)
            {
                if (_opToken.TokenID == DataFrameExpressionToken.IS || _opToken.TokenID == DataFrameExpressionToken.COMP_EQUALS)
                {
                    if (leftNullColumn is not null)
                        return leftNullColumn.IsNull ? right.ElementwiseIsNull() : right.ElementwiseIsNotNull();
                    else if (rightNullColumn is not null)
                        return rightNullColumn.IsNull ? left.ElementwiseIsNull() : left.ElementwiseIsNotNull();
                }
                else if (_opToken.TokenID == DataFrameExpressionToken.COMP_NOTEQUALS)
                {
                    if (leftNullColumn is not null)
                        return leftNullColumn.IsNull ? right.ElementwiseIsNotNull() : right.ElementwiseIsNull();
                    else if (rightNullColumn is not null)
                        return rightNullColumn.IsNull ? left.ElementwiseIsNotNull() : left.ElementwiseIsNull();
                }
                throw new DataFrameExpressionException($"'{_opToken.Value}' not valid for null comparison, only equals, not equals and is");
            }

            if (_opToken.TokenID == DataFrameExpressionToken.ARITH_PLUS && (left is StringDataFrameColumn || right is StringDataFrameColumn))
            {
                return left.ElementwiseConcat(right);
            }

            return _opToken.TokenID switch
            {
                DataFrameExpressionToken.ARITH_PLUS => left.Add(right),
                DataFrameExpressionToken.ARITH_MINUS => left.Subtract(right),
                DataFrameExpressionToken.ARITH_TIMES => left.Multiply(right),
                DataFrameExpressionToken.ARITH_DIVIDE => left.Divide(right),
                DataFrameExpressionToken.ARITH_MODULO => left.Modulo(right),
                DataFrameExpressionToken.CARET => left.ElementwiseExponent(right),

                // use safe extensions for logical comparisons
                DataFrameExpressionToken.AND => left.And(right),
                DataFrameExpressionToken.OR => left.Or(right),
                DataFrameExpressionToken.XOR => left.Xor(right),

                DataFrameExpressionToken.COMP_EQUALS => left.ElementwiseEqualsDateAware(right),
                DataFrameExpressionToken.COMP_NOTEQUALS => left.ElementwiseNotEqualsDateAware(right),
                DataFrameExpressionToken.COMP_LT => left.ElementwiseLessThanDateAware(right),
                DataFrameExpressionToken.COMP_GT => left.ElementwiseGreaterThanDateAware(right),
                DataFrameExpressionToken.COMP_LTE => left.ElementwiseLessThanOrEqualDateAware(right),
                DataFrameExpressionToken.COMP_GTE => left.ElementwiseGreaterThanOrEqualDateAware(right),
                _ => throw new DataFrameExpressionException($"no binary operation defined for token {_opToken.Label}"),
            };
        }
        catch (NotSupportedException ex)
        {
            throw new DataFrameExpressionException($"'{_opToken.Value}' operator is invalid between columns of type '{left.DataType.Name}' and '{right.DataType.Name}'", ex, _opToken.TokenID);
        }
        catch (Exception ex) when (ex is not DataFrameExpressionException)
        {
            throw new DataFrameExpressionException($"Error parsing operation '{_opToken.Value}' on columns of type '{left.DataType.Name}' and '{right.DataType.Name}'", ex, _opToken.TokenID);
        }
    }
}
