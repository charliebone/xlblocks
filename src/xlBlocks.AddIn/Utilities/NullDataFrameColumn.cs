namespace XlBlocks.AddIn.Utilities;

using Microsoft.Data.Analysis;

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