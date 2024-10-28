namespace XlBlocks.AddIn.Excel.DataTypes;

using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataTypes_String
{
    [return: CacheContents]
    [ExcelFunction(Description = "Split a string into a range", IsThreadSafe = true)]
    public static XlBlockRange XBString_Split(
        [ExcelArgument(Description = "A string")] string stringValue,
        [ExcelArgument(Description = "A delimiter")] string delimiter,
        [ExcelArgument(Description = "Trim strings (TRUE)")] bool trimStrings = true,
        [ExcelArgument(Description = "Ignore empty strings (TRUE)")] bool ignoreEmpty = true)
    {
        return XlBlockRange.BuildFromString(stringValue, delimiter, trimStrings, ignoreEmpty);
    }

}
