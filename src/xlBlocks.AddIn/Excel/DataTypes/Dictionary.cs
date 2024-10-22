namespace XlBlocks.AddIn.Excel.DataTypes;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataTypes_Dictionary
{
    [return: CacheContents]
    [ExcelFunction(Description = "Build a dictionary from a range", IsThreadSafe = true)]
    public static XlBlockDictionary? xbDict_Build(
        [ExcelArgument(Description = "A range of data to use for dictionary keys")] XlBlockRange keys,
        [ExcelArgument(Description = "A range of data to use for dictionary values")] XlBlockRange values,
        [ExcelArgument(Description = "Error handling ('drop' or 'error'). Default is 'drop'")] string onErrors = "drop")
    {
        return XlBlockDictionary.Build(keys, values, onErrors);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a dictionary of a given key type from a range", IsThreadSafe = true)]
    public static XlBlockDictionary? xbDict_BuildTyped(
        [ExcelArgument(Description = "A range of data to use for dictionary keys")] XlBlockRange keys,
        [ExcelArgument(Description = "The data type of the dictionary keys")] string keyType,
        [ExcelArgument(Description = "A range of data to use for dictionary values")] XlBlockRange values,
        [ExcelArgument(Description = "Error handling ('drop' or 'error'). Default is 'drop'")] string onErrors = "drop")
    {
        return XlBlockDictionary.BuildTyped(keys, keyType, values, onErrors);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a dictionary from two lists", IsThreadSafe = true)]
    public static XlBlockDictionary? xbDict_BuildFromLists(
        [ExcelArgument(Description = "A list of keys to use in constructing the dictionary"), CacheContents(AsReference = true)] XlBlockList keyList,
        [ExcelArgument(Description = "A list of values to use in constructing the dictionary"), CacheContents(AsReference = true)] XlBlockList valueList)
    {
        return XlBlockDictionary.BuildFromLists(keyList, valueList);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a list from the contents of a file", IsThreadSafe = true)]
    public static XlBlockDictionary? xbDict_BuildFromFile(
        [ExcelArgument(Description = "A filepath")] string filepath,
        [ExcelArgument(Description = "A delimiter")] string delimiter,
        [ExcelArgument(Description = "Trim strings, optional (FALSE)")] bool trimStrings = false,
        [ExcelArgument(Description = "Ignore empty strings (TRUE)")] bool ignoreEmpty = true)
    {
        throw new NotImplementedException();
    }

    [ExcelFunction(Description = "Get the contents of a dictionary as a range", IsThreadSafe = true)]
    public static object[,] xbDict_Get(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict,
        [ExcelArgument(Description = "Include header in output (FALSE)")] bool includeHeader = false,
        [ExcelArgument(Description = "The name to use for the key column ('Key')")] string keyColumn = "Key",
        [ExcelArgument(Description = "The name to use for the value column ('Value')")] string valueColumn = "Value")
    {
        return dict.AsArray(includeHeader, keyColumn, valueColumn);
    }

    [ExcelFunction(Description = "Get the keys of a dictionary as a range", IsThreadSafe = true)]
    public static object? xbDict_GetValue(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict,
        [ExcelArgument(Description = "A key for the dictionary")] string key)
    {
        return dict[key];
    }

    [ExcelFunction(Description = "Get the keys of a dictionary as a range", IsThreadSafe = true)]
    public static object[] xbDict_GetKeys(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict)
    {
        return dict.Keys;
    }

    [ExcelFunction(Description = "Get the values of a dictionary as a range", IsThreadSafe = true)]
    public static object[] xbDict_GetValues(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict)
    {
        return dict.Values;
    }

    [ExcelFunction(Description = "Get the keys of a dictionary as a range", IsThreadSafe = true)]
    public static bool xbDict_ContainsKey(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict,
        [ExcelArgument(Description = "A key for the dictionary")] string key)
    {
        return dict.ContainsKey(key);
    }

    [ExcelFunction(Description = "Return the number of entries in a dictionary", IsThreadSafe = true)]
    public static int xbDict_Count(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict)
    {
        return dict.Count;
    }

    [ExcelFunction(Description = "Return the number of entries in a dictionary as a string", IsThreadSafe = true)]
    public static string xbDict_Dim(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict)
    {
        return dict.Count.ToString("N0");
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Convert a dictionary into a table", IsThreadSafe = true)]
    public static object? xbDict_ToTable(
        [ExcelArgument(Description = "A dictionary"), CacheContents(AsReference = true)] XlBlockDictionary dict,
        [ExcelArgument(Description = "The name to use for the key column ('Key')")] string keyColumn = "Key",
        [ExcelArgument(Description = "The name to use for the value column ('Value')")] string valueColumn = "Value")
    {
        throw new NotImplementedException();
    }
}
