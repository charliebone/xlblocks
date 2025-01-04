namespace XlBlocks.AddIn.Excel.DataTypes;

using System.Text.RegularExpressions;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

internal static class DataTypes_String
{
    [ExcelFunction(Description = "Split a string into a range", IsThreadSafe = true)]
    public static XlBlockRange XBString_Split(
        [ExcelArgument(Description = "A string")] string stringValue,
        [ExcelArgument(Description = "A delimiter")] string delimiter,
        [ExcelArgument(Description = "Trim strings")] bool trimStrings = true,
        [ExcelArgument(Description = "Ignore empty strings")] bool ignoreEmpty = true)
    {
        return XlBlockRange.BuildFromString(stringValue, delimiter, trimStrings, ignoreEmpty);
    }

    [ExcelFunction(Description = "Check a string for a regex match", IsThreadSafe = true)]
    public static bool XBString_RegexMatch(
        [ExcelArgument(Description = "A string")] string stringValue,
        [ExcelArgument(Description = "A regex pattern")] string pattern,
        [ExcelArgument(Description = "Case insensitive")] bool caseSensitive = true)
    {
        return Regex.IsMatch(stringValue, pattern, caseSensitive == false ? RegexOptions.IgnoreCase : RegexOptions.None);
    }

    [ExcelFunction(Description = "Split a string into a range", IsThreadSafe = true)]
    public static string XBString_RegexReplace(
        [ExcelArgument(Description = "A string")] string stringValue,
        [ExcelArgument(Description = "A regex pattern")] string pattern,
        [ExcelArgument(Description = "A replacement value")] string replacement)
    {
        return Regex.Replace(stringValue, pattern, replacement);
    }
}
