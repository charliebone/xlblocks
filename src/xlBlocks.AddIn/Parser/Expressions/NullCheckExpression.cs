namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;

internal sealed class NullExpression : IColumnExpression
{
    public NullExpression() { }

    public DataFrameColumn Evaluate(DataFrameContext context)
    {
        return new NullDataFrameColumn(true);
    }
}

internal sealed class NullDataFrameColumn : PrimitiveDataFrameColumn<bool>
{
    public bool IsNull { get; private set; }

    public NullDataFrameColumn(bool isNull) : base("null", 0)
    {
        IsNull = isNull;
    }

    public NullDataFrameColumn Negate()
    {
        IsNull = !IsNull;
        return this;
    }
}
