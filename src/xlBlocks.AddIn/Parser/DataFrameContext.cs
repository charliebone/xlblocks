namespace XlBlocks.AddIn.Parser;

using Microsoft.Data.Analysis;

internal class DataFrameContext
{
    private readonly DataFrame _dataFrame;

    public DataFrameContext(DataFrame dataFrame)
    {
        _dataFrame = dataFrame;
    }

    public DataFrame DataFrame => _dataFrame;
}
