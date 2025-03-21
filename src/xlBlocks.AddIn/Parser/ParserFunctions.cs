﻿namespace XlBlocks.AddIn.Parser;

using System.Collections.Generic;
using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser.Expressions;
using XlBlocks.AddIn.Utilities;

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
        CheckArgCount(argExpressions, 2);

        var column = argExpressions[0].Evaluate(context);
        var valueIfNullColumn = argExpressions[1].Evaluate(context);

        return column.ElementwiseIsNotNull().ElementwiseIfThenElse(column, valueIfNullColumn);
    }

    public static DataFrameColumn Iif(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 3);

        var conditionalColumn = argExpressions[0].Evaluate(context);
        var trueColumn = argExpressions[1].Evaluate(context);
        var falseColumn = argExpressions[2].Evaluate(context);

        return conditionalColumn.ElementwiseIfThenElse(trueColumn, falseColumn);
    }

    public static DataFrameColumn Len(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);
        return expressionColumn.ElementwiseLength();
    }

    public static DataFrameColumn Exp(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        if (argExpressions.Count == 2)
        {
            var baseColumn = argExpressions[0].Evaluate(context);
            var exponentColumn = argExpressions[1].Evaluate(context);
            return baseColumn.ElementwiseExponent(exponentColumn);
        }
        else
        {
            var exponentColumn = argExpressions[0].Evaluate(context);
            var baseColumn = DataFrameUtilities.CreateConstantDataFrameColumn(Math.E, exponentColumn.Length);
            return baseColumn.ElementwiseExponent(exponentColumn);
        }
    }

    public static DataFrameColumn Log(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var baseColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.ElementwiseLog(baseColumn);
        }

        return expressionColumn.ElementwiseLog();
    }

    public static DataFrameColumn CumulativeSum(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var conditionalColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.CumulativeSumIf(conditionalColumn);
        }

        return expressionColumn.CumulativeSum();
    }

    public static DataFrameColumn CumulativeProduct(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var conditionalColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.CumulativeProductIf(conditionalColumn);
        }

        return expressionColumn.CumulativeProduct();
    }

    public static DataFrameColumn CumulativeMin(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var conditionalColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.CumulativeMinIf(conditionalColumn);
        }

        return expressionColumn.CumulativeMin();
    }

    public static DataFrameColumn CumulativeMax(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var conditionalColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.CumulativeMaxIf(conditionalColumn);
        }

        return expressionColumn.CumulativeMax();
    }

    public static DataFrameColumn Round(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);

        if (argExpressions.Count == 2)
        {
            var digitsColumn = argExpressions[1].Evaluate(context);
            return expressionColumn.ElementwiseRound(digitsColumn);
        }

        return expressionColumn.ElementwiseRound();
    }

    public static DataFrameColumn Abs(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);

        return expressionColumn.ElementwiseAbs();
    }

    public static DataFrameColumn Min(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var otherColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseMin(otherColumn);
    }

    public static DataFrameColumn Max(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var otherColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseMax(otherColumn);
    }

    public static DataFrameColumn Floor(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);

        return expressionColumn.ElementwiseFloor();
    }

    public static DataFrameColumn Ceiling(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);

        return expressionColumn.ElementwiseCeiling();
    }

    public static DataFrameColumn Substring(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2, 3);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var startIndexColumn = argExpressions[1].Evaluate(context);

        if (argExpressions.Count == 3)
        {
            var lengthColumn = argExpressions[2].Evaluate(context);
            return expressionColumn.ElementwiseSubstring(startIndexColumn, lengthColumn);
        }

        return expressionColumn.ElementwiseSubstring(startIndexColumn);
    }

    public static DataFrameColumn Left(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var lengthColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseLeft(lengthColumn);
    }

    public static DataFrameColumn Right(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var lengthColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseRight(lengthColumn);
    }

    public static DataFrameColumn Trim(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 1);

        var expressionColumn = argExpressions[0].Evaluate(context);
        return expressionColumn.ElementwiseTrim();
    }

    public static DataFrameColumn Replace(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 3, 4);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var oldValueColumn = argExpressions[1].Evaluate(context);
        var newValueColumn = argExpressions[2].Evaluate(context);

        if (argExpressions.Count == 4)
        {
            var caseSensitiveColumn = argExpressions[3].Evaluate(context);
            return expressionColumn.ElementwiseReplace(oldValueColumn, newValueColumn, caseSensitiveColumn);
        }

        return expressionColumn.ElementwiseReplace(oldValueColumn, newValueColumn, null);
    }

    public static DataFrameColumn Regex_Test(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2, 3);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var patternColumn = argExpressions[1].Evaluate(context);

        if (argExpressions.Count == 3)
        {
            var caseSensitiveColumn = argExpressions[2].Evaluate(context);
            return expressionColumn.ElementwiseRegexTest(patternColumn, caseSensitiveColumn);
        }

        return expressionColumn.ElementwiseRegexTest(patternColumn, null);
    }

    public static DataFrameColumn Regex_Find(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2, 3);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var patternColumn = argExpressions[1].Evaluate(context);

        if (argExpressions.Count == 3)
        {
            var caseSensitiveColumn = argExpressions[2].Evaluate(context);
            return expressionColumn.ElementwiseRegexFind(patternColumn, caseSensitiveColumn);
        }

        return expressionColumn.ElementwiseRegexFind(patternColumn, null);
    }

    public static DataFrameColumn Regex_Replace(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 3);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var patternColumn = argExpressions[1].Evaluate(context);
        var replaceColumn = argExpressions[2].Evaluate(context);

        return expressionColumn.ElementwiseRegexReplace(patternColumn, replaceColumn);
    }

    public static DataFrameColumn Format(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var formatColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseFormat(formatColumn);
    }

    public static DataFrameColumn ToDateTime(DataFrameContext context, IList<IColumnExpression> argExpressions)
    {
        CheckArgCount(argExpressions, 2);

        var expressionColumn = argExpressions[0].Evaluate(context);
        var formatColumn = argExpressions[1].Evaluate(context);

        return expressionColumn.ElementwiseToDateTime(formatColumn);
    }

    #endregion
}
