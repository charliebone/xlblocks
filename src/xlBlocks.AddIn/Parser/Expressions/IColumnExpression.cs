namespace XlBlocks.AddIn.Parser.Expressions;

using Microsoft.Data.Analysis;
using XlBlocks.AddIn.Parser;

internal interface IColumnExpression
{
    DataFrameColumn Evaluate(DataFrameContext context);
}
