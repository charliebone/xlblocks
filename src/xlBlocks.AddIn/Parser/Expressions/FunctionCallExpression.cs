namespace XlBlocks.AddIn.Parser.Expressions;

using System.Collections.Generic;
using Microsoft.Data.Analysis;
using sly.lexer;
using XlBlocks.AddIn.Parser;

internal sealed class FunctionCallExpression : IColumnExpression
{
    private readonly string _functionName;
    private readonly List<IColumnExpression> _arguments;

    public FunctionCallExpression(Token<DataFrameExpressionToken> identifier, IColumnExpression expressionList)
    {
        _functionName = identifier.StringWithoutQuotes;

        if (expressionList is ExpressionListExpression expressionListExpression)
            _arguments = expressionListExpression.Expressions;
        else
            _arguments = new List<IColumnExpression>();
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        try
        {
            switch (_functionName.ToUpperInvariant())
            {
                case "ISNULL":
                    return ParserFunctions.IsNull(context, _arguments);
                case "IIF":
                    return ParserFunctions.Iif(context, _arguments);
                case "LEN":
                    return ParserFunctions.Len(context, _arguments);
                case "EXP":
                    return ParserFunctions.Exp(context, _arguments);
                case "LOG":
                    return ParserFunctions.Log(context, _arguments);
                case "CUMSUM":
                    return ParserFunctions.CumulativeSum(context, _arguments);
                case "CUMPROD":
                    return ParserFunctions.CumulativeProduct(context, _arguments);
                case "CUMMIN":
                    return ParserFunctions.CumulativeMin(context, _arguments);
                case "CUMMAX":
                    return ParserFunctions.CumulativeMax(context, _arguments);
                case "ROUND":
                    return ParserFunctions.Round(context, _arguments);
                case "SUBSTRING":
                    return ParserFunctions.Substring(context, _arguments);
                case "LEFT":
                    return ParserFunctions.Left(context, _arguments);
                case "RIGHT":
                    return ParserFunctions.Right(context, _arguments);
                case "TRIM":
                    return ParserFunctions.Trim(context, _arguments);
                case "REPLACE":
                    return ParserFunctions.Replace(context, _arguments);
                case "REGEX_TEST":
                    return ParserFunctions.Regex_Test(context, _arguments);
                case "REGEX_FIND":
                    return ParserFunctions.Regex_Find(context, _arguments);
                case "REGEX_REPLACE":
                    return ParserFunctions.Regex_Replace(context, _arguments);
                case "FORMAT":
                    return ParserFunctions.Format(context, _arguments);
                case "TODATE":
                    return ParserFunctions.ToDateTime(context, _arguments);
            }
        }
        catch (Exception ex)
        {
            throw new DataFrameExpressionException($"Error evaluating function '{_functionName}': {ex.Message}", ex);
        }
        throw new DataFrameExpressionException($"Unknown function '{_functionName}'");
    }
}
