namespace XlBlocks.AddIn.Parser;

using System;

internal class DataFrameExpressionException : Exception
{
    public DataFrameExpressionToken Token { get; }

    public DataFrameExpressionException(string message, DataFrameExpressionToken? token = null) : base(message)
    {
        if (token.HasValue)
            Token = token.Value;
    }

    public DataFrameExpressionException(string message, Exception parserException, DataFrameExpressionToken? token = null) : base(message, parserException)
    {
        if (token.HasValue)
            Token = token.Value;
    }
}
