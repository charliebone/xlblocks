namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class ColumnExpression : IColumnExpression
{
    private readonly string _columnName;

    public ColumnExpression(string columnName)
    {
        _columnName = columnName;
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        return context.DataFrame[_columnName].Clone();
    }
}

internal sealed class ConstantColumnExpression<T> : IColumnExpression
{
    private readonly T _constant;

    public ConstantColumnExpression(T constant)
    {
        if (!DataFrameUtilities.IsValidColumnType(typeof(T)))
            throw new DataFrameExpressionException($"type of {typeof(T).Name} is not valid for use as a constant expression");

        _constant = constant;
    }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        return DataFrameUtilities.CreateDataFrameColumn(RepeatLong(_constant, context.DataFrame.Rows.Count));
    }

    private static IEnumerable<TResult> RepeatLong<TResult>(TResult element, long count)
    {
        for (var i = 0L; i < count; i++) yield return element;
    }
}
