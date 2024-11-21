namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;
using XlBlocks.AddIn.Utilities;

internal sealed class NullExpression : IColumnExpression
{
    public NullExpression() { }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        return new NullDataFrameColumn(true);
    }
}
