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
    public static string? XBUtils_GetVersion()
    {
        return XlBlocksAddIn.Version;
    }

    [ExcelFunction(Description = "Get the username of the current user")]
    public static string XBUtils_GetUsername(
        [ExcelArgument(Description = "Include the domain name in the username (FALSE)")] bool includeDomain = false)
    {
        return $"{(includeDomain ? $"{Environment.UserDomainName}\\" : "")}{Environment.UserName}";
    }

    [ExcelFunction(Description = "Get the name of the current machine")]
    public static string XBUtils_GetMachineName()
    {
        return Environment.MachineName;
    }
}
