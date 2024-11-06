namespace XlBlocks.AddIn.Excel.Utilities;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlBlocks.AddIn.Cache;
using XlBlocks.AddIn.Dna;

internal static class CacheUtilities
{
    [return: CacheContents]
    [ExcelFunction(Description = "Put a value into the XlBlocks cache", IsThreadSafe = true)]
    public static object XBCache_Set(
        [ExcelArgument(Description = "The object to be cached")] object obj)
    {
        return obj;
    }

    [ExcelFunction(Description = "Get a value from the XlBlocks cache", IsThreadSafe = true)]
    public static object XBCache_Get(
        [ExcelArgument(Description = "The object to get from the cache"), CacheContents] object obj)
    {
        return obj;
    }

    [ExcelFunction(Description = "Check if a non-null value exists in the XlBlocks cache", IsThreadSafe = true)]
    public static object XBCache_Exists(
        [ExcelArgument(Description = "The object to check for"), CacheContents(DefaultIfMissing = true)] object obj)
    {
        return obj is not null;
    }

    [ExcelFunction(Description = "Get a cache handle for a cell", IsThreadSafe = true)]
    public static object XBCache_Handle(
        [ExcelArgument(AllowReference = true, Description = "The cell for which to get the cache handle. Optional, defaults to current cell"), Optional] object cell)
    {
        string reference;
        if (cell is not null && cell is ExcelReference excelReference)
        {
            reference = excelReference.ExcelReferenceToString();
        }
        else
        {
            if (!CacheHelper.TryGetCallingReference(out var callingReference, () => "Unknown"))
                throw new Exception($"{nameof(XBCache_Handle)} can only be called with a cell reference or from a spreadsheet");

            reference = callingReference;
        }

        var _ = ObjectCache.GetCacheKey(reference, out var hexKey);
        return hexKey;
    }
}
