namespace XlBlocks.AddIn.Excel.Utilities;

using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;

internal static class Utilities
{
    [ExcelFunction(Description = "Toggle error output behavior, useful when debugging")]
    public static object XBUtils_SetErrorOutput(
        [ExcelArgument(Description = "Boolean, will print error output to screen when TRUE")] bool printErrors)
    {
        return XlBlocksAddIn.PrintExceptions = printErrors;
    }

    [ExcelFunction(Description = "Get the XlBlocks version string")]
    public static object? XBUtils_GetVersion()
    {
        return XlBlocksAddIn.Version;
    }
}
