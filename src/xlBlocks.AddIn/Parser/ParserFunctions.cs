namespace XlBlocks.AddIn.Parser;

using System.Collections.Generic;
using Microsoft.Data.Analysis;
using xlBlocks.AddIn.Utilities;
using XlBlocks.AddIn.Parser.Expressions;

internal static class ParserFunctions
{
    private static void CheckArgCount(IList<IColumnExpression> args, int count)
    {
        if (args.Count != count)
            throw new DataFrameExpressionException($"Expected {count} arg{(count > 1 ? "s" : "")}, got {args.Count}");
    }

    private static void CheckArgCount(IList<IColumnExpression> args, int lowerBound, int upperBound)
    {
        if (args.Count < lowerBound)
            throw new DataFrameExpressionException($"Expected at least {lowerBound} arg{(lowerBound > 1 ? "s" : "")}, got {args.Count}");

        if (args.Count > upperBound)
            throw new DataFrameExpressionException($"Expected at most {upperBound} arg{(upperBound > 1 ? "s" : "")}, got {args.Count}");
    }

    #region Expression parser functions

    public static DataFrameColumn IsNull(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var column = argExpressions[0].Evaluate(context) ?? throw new DataFrameExpressionException("Error processing argument");
        return column.ElementwiseIsNull();
    }

    public static DataFrameColumn Iif(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 3);

        var conditionalColumn = argExpressions[0].Evaluate(context);
        if (conditionalColumn is not PrimitiveDataFrameColumn<bool> boolColumn)
            throw new DataFrameExpressionException("The conditional argument must evaluate to a boolean");

        var trueColumn = argExpressions[1].Evaluate(context);
        var falseColumn = argExpressions[2].Evaluate(context);

        if (trueColumn.DataType != falseColumn.DataType)
            throw new DataFrameExpressionException($"True and false columns must be the same data type, got '{trueColumn.DataType.Name}' and '{falseColumn.DataType.Name}'");

        return boolColumn.ElmentwiseIfThenElse(trueColumn, falseColumn);
    }

    public static DataFrameColumn Len(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);
        if (expressionColumn is not StringDataFrameColumn strColumn)
            throw new DataFrameExpressionException("Function is only valid for computing length on string columns");

        return strColumn.ElementwiseLength();
    }

    public static DataFrameColumn Substring(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2, 3);

        var expressionColumn = argExpressions[0].Evaluate(context);
        if (expressionColumn is not StringDataFrameColumn strColumn)
            throw new DataFrameExpressionException("Function is only valid on string columns");

        var startIndexColumn = argExpressions[1].Evaluate(context);
        if (!startIndexColumn.IsNumericColumn())
            throw new DataFrameExpressionException("Start index must be numeric");

        if (argExpressions.Count == 3)
        {
            var lengthColumn = argExpressions[2].Evaluate(context);
            if (!lengthColumn.IsNumericColumn())
                throw new DataFrameExpressionException("Length must be numeric");

            return strColumn.ElementwiseSubstring(startIndexColumn, lengthColumn);
        }

        return strColumn.ElementwiseSubstring(startIndexColumn);
    }

    public static DataFrameColumn Trim(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);
        if (expressionColumn is not StringDataFrameColumn strColumn)
            throw new DataFrameExpressionException("Function is only valid on string columns");

        return strColumn.ElementwiseTrim();
    }

    #endregion
}
