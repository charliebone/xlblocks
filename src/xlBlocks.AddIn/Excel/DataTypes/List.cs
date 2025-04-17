namespace XlBlocks.AddIn.Excel.DataTypes;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.Types;

internal static class DataTypes_List
{
    [return: CacheContents]
    [ExcelFunction(Description = "Build a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Build(
        [ExcelArgument(Description = "A range of data to include in the list")] XlBlockRange range,
        [ExcelArgument(Description = "Error handling ('drop', 'keep', or 'error')")] string onErrors = "drop")
    {
        return XlBlockList.Build(range, onErrors);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a list of a given type", IsThreadSafe = true)]
    public static XlBlockList XBList_BuildTyped(
        [ExcelArgument(Description = "A range of data to include in the list")] XlBlockRange items,
        [ExcelArgument(Description = "The data type of the list")] string type,
        [ExcelArgument(Description = "Error handling ('drop', 'keep', or 'error')")] string onErrors = "drop")
    {
        return XlBlockList.BuildTyped(items, type, onErrors);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a list from a string", IsThreadSafe = true)]
    public static XlBlockList XBList_BuildFromString(
        [ExcelArgument(Description = "A string")] string str,
        [ExcelArgument(Description = "A delimiter")] string delimiter,
        [ExcelArgument(Description = "Trim strings")] bool trimStrings = false,
        [ExcelArgument(Description = "Ignore empty strings")] bool ignoreEmpty = true)
    {
        return XlBlockList.BuildFromString(str, delimiter, trimStrings, ignoreEmpty);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Build a list from the contents of a file", IsThreadSafe = true)]
    public static XlBlockList XBList_BuildFromFile(
        [ExcelArgument(Description = "A filepath")] string filepath,
        [ExcelArgument(Description = "A delimiter")] string delimiter,
        [ExcelArgument(Description = "Trim strings")] bool trimStrings = false,
        [ExcelArgument(Description = "Ignore empty strings")] bool ignoreEmpty = true)
    {
        var fileContents = File.ReadAllText(filepath);
        return XlBlockList.BuildFromString(fileContents, delimiter, trimStrings, ignoreEmpty);
    }

    [ExcelFunction(Description = "Concatenate the items of a list into a string", IsThreadSafe = true)]
    public static string XBList_ToString(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list,
        [ExcelArgument(Description = "A string to use as the separator between the joined items")] string separator = ", ")
    {
        return list.ContentsToString(separator);
    }

    [ExcelFunction(Description = "Get the contents of a list as a range", IsThreadSafe = true)]
    public static XlBlockList XBList_Get(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list)
    {
        return list;
    }

    [ExcelFunction(Description = "Get an element of a list", IsThreadSafe = true)]
    public static object? XBList_GetAt(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list,
        [ExcelArgument(Description = "A 1-based index of the list element to return")] int index)
    {
        if (index <= 0 || index > list.Count)
            throw new IndexOutOfRangeException($"Index '{index}' is out of range");

        return list.GetAt(index - 1);
    }

    [ExcelFunction(Description = "Get the size of a list", IsThreadSafe = true)]
    public static int XBList_Count(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list)
    {
        return list.Count;
    }

    [ExcelFunction(Description = "Get the size of a list as a string", IsThreadSafe = true)]
    public static string XBList_Dim(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list)
    {
        return list.Count.ToString("N0");
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Add an item to a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Add(
        [ExcelArgument(Description = "A list"), CacheContents] XlBlockList list,
        [ExcelArgument(Description = "One or more items to be removed")] XlBlockRange items,
        [ExcelArgument(Description = "Error handling ('drop', 'keep', or 'error')")] string onErrors = "drop")
    {
        list.Add(items, onErrors);
        return list;
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Remove an item from a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Remove(
        [ExcelArgument(Description = "A list"), CacheContents] XlBlockList list,
        [ExcelArgument(Description = "One or more items to be removed")] XlBlockRange items)
    {
        list.Remove(items);
        return list;
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Sort a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Sort(
        [ExcelArgument(Description = "A list"), CacheContents] XlBlockList list,
        [ExcelArgument(Description = "Descending sort")] bool descending = false)
    {
        list.Sort(descending);
        return list;
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Take the first N items of a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Take(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list,
        [ExcelArgument(Description = "The number of items to take")] int n)
    {
        return list.Take(n);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Skip the first N items of a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Skip(
        [ExcelArgument(Description = "A list"), CacheContents(AsReference = true)] XlBlockList list,
        [ExcelArgument(Description = "The number of items to skip")] int n)
    {
        return list.Skip(n);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Get a list of the unique items in a list", IsThreadSafe = true)]
    public static XlBlockList? XBList_UniqueItems(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents(AsReference = true)] XlBlockList list)
    {
        return list.GetUniqueItems();
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Get a list of duplicated items in a list", IsThreadSafe = true)]
    public static XlBlockList XBList_DuplicateItems(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents(AsReference = true)] XlBlockList list)
    {
        return list.GetDuplicateItems();
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Reverse the order of the contents of a list", IsThreadSafe = true)]
    public static XlBlockList XBList_Reverse(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents] XlBlockList list)
    {
        return list.Reverse();
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Combine multiple lists into one, retaining duplicate values", IsThreadSafe = true)]
    public static XlBlockList XBList_UnifyLists(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents] params XlBlockList[] lists)
    {
        return XlBlockList.UnifyLists(lists);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Merge multiple lists into one, dropping duplicate values", IsThreadSafe = true)]
    public static XlBlockList XBList_MergeLists(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents] params XlBlockList[] lists)
    {
        return XlBlockList.MergeLists(lists);
    }

    [return: CacheContents]
    [ExcelFunction(Description = "Make a list containing the intersection (shared items) of input lists", IsThreadSafe = true)]
    public static XlBlockList XBList_IntersectLists(
        [ExcelArgument(Name = "list", Description = "A list"), CacheContents] params XlBlockList[] lists)
    {
        return XlBlockList.IntersectLists(lists);
    }
}
