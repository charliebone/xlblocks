namespace XlBlocks.AddIn.Tests.Types;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

public class XlBlockListTests
{
    #region Build tests

    [Fact]
    public void Build_EmptyArray_ReturnsEmptyList()
    {
        var result = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Build_IgnoreMissingOrError_RemovesBadValues()
    {
        var items = new object[,] { { 1, 2, ExcelEmpty.Value, ExcelError.ExcelErrorNA, ExcelMissing.Value, 3 } };
        var result = XlBlockList.Build(XlBlockRange.Build(items), "drop");
        var expected = new object[] { 1, 2, 3 };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void Build_IncludeMissingOrError_ThrowsArgumentException()
    {
        var items = new object[,] { { 1, 2, ExcelEmpty.Value, ExcelError.ExcelErrorNA, ExcelMissing.Value, 3 } };
        Assert.Throws<ArgumentException>(() => XlBlockList.Build(XlBlockRange.Build(items), "error"));
    }

    [Fact]
    public void BuildTyped_InvalidType_ThrowsArgumentException()
    {
        var items = new object[,] { { "abc" } };
        Assert.Throws<ArgumentException>(() => XlBlockList.BuildTyped(XlBlockRange.Build(items), "double", "drop"));
    }

    [Fact]
    public void BuildTyped_IncludeMissingOrError_ThrowsArgumentException()
    {
        var items = new object[,] { { 1, 2, ExcelEmpty.Value, ExcelError.ExcelErrorNA, ExcelMissing.Value, 3 } };
        Assert.Throws<ArgumentException>(() => XlBlockList.BuildTyped(XlBlockRange.Build(items), "double", "error"));
    }

    [Fact]
    public void BuildTyped_IgnoreMissingOrError_RemovesBadValues()
    {
        var items = new object[,] { { 1, 2, ExcelEmpty.Value, ExcelError.ExcelErrorNA, ExcelMissing.Value, 3 } };
        var result = XlBlockList.BuildTyped(XlBlockRange.Build(items), "double", "drop");
        var expected = new object[] { 1.0, 2.0, 3.0 };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void BuildFromString_DelimitedString_ReturnsList()
    {
        var result = XlBlockList.BuildFromString("a,b,c", ",");
        var expected = new object[] { "a", "b", "c" };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void BuildFromString_TrimStrings_ReturnsTrimmedList()
    {
        var result = XlBlockList.BuildFromString(" a , b , c ", ",", true);
        var expected = new object[] { "a", "b", "c" };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void BuildFromString_IgnoreEmpty_ReturnsNonEmptyList()
    {
        var result = XlBlockList.BuildFromString("a,,b,,c", ",", false, true);
        var expected = new object[] { "a", "b", "c" };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void BuildFromString_WithEmptyStrings_IncludesEmpty()
    {
        var result = XlBlockList.BuildFromString("a,,b,,c", ",", false, false);
        var expected = new object[] { "a", "", "b", "", "c" };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void BuildFromString_CustomDelimiter_ReturnsList()
    {
        var result = XlBlockList.BuildFromString("a|b|c", "|");
        var expected = new object[] { "a", "b", "c" };
        Assert.Equal(expected, result.Get());
    }

    #endregion

    #region Add / Remove tests

    [Fact]
    public void Add_InvalidType_ThrowsArgumentException()
    {
        var list = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1.0 } }), "double", "error");
        Assert.Throws<ArgumentException>(() => list.Add("string"));
    }

    [Fact]
    public void Add_ArrayOfItems_AddsItemsToList()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        var itemsToAdd = new object[,] { { 1, 2, 3 } };
        list.Add(XlBlockRange.Build(itemsToAdd), "drop");

        var expected = new object[] { 1, 2, 3 };
        Assert.Equal(expected, list.Get());
    }

    [Fact]
    public void Add_ArrayDropErrors_RemovesThem()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        var itemsToAdd = new object[,] { { 1, ExcelEmpty.Value, 2, ExcelError.ExcelErrorNA, 3 } };
        list.Add(XlBlockRange.Build(itemsToAdd), "drop");

        var expected = new object[] { 1, 2, 3 };
        Assert.Equal(expected, list.Get());
    }

    [Fact]
    public void Add_ArrayThrowOnError_ThrowsArgumentException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        var itemsToAdd = new object[,] { { 1, ExcelEmpty.Value, 2, ExcelError.ExcelErrorNA, 3 } };
        Assert.Throws<ArgumentException>(() => list.Add(XlBlockRange.Build(itemsToAdd), "error"));
    }

    [Fact]
    public void Add_Item_AddsToList()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        list.Add(5);
        Assert.Equal(1, list.Count);
    }

    [Fact]
    public void Remove_ArrayOfItems_RemovesItemsFromList()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var itemsToRemove = new object[,] { { 2, 4 } };
        list.Remove(XlBlockRange.Build(itemsToRemove));

        var expected = new object[] { 1, 3, 5 };
        Assert.Equal(expected, list.Get());
    }

    [Fact]
    public void Remove_ArrayWithMissingOrErrorValues_IgnoresMissing()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, ExcelEmpty.Value, 3, ExcelError.ExcelErrorNA, 4, ExcelMissing.Value, 5 } }), "drop");
        var itemsToRemove = new object[,] { { 2, ExcelEmpty.Value, 4, ExcelError.ExcelErrorNA, ExcelMissing.Value } };
        list.Remove(XlBlockRange.Build(itemsToRemove));

        var expected = new object[] { 1, 3, 5 };
        Assert.Equal(expected, list.Get());
    }

    [Fact]
    public void RemoveAt_ValidIndex_RemovesItemAtIndex()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        list.RemoveAt(2);

        var expected = new object[] { 1, 2, 4, 5 };
        Assert.Equal(expected, list.Get());
    }

    [Fact]
    public void RemoveAt_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        Assert.Throws<ArgumentOutOfRangeException>(() => list.RemoveAt(5));
    }

    [Fact]
    public void Remove_Item_RemovesFromList()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3 } }), "drop");
        list.Remove(2);
        Assert.Equal(2, list.Count);
    }

    #endregion

    #region List tests

    [Fact]
    public void Take_NegativeNumber_ThrowsArgumentOutOfRangeException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        Assert.Throws<ArgumentOutOfRangeException>(() => list.Take(-1));
    }

    [Fact]
    public void Skip_NegativeNumber_ThrowsArgumentOutOfRangeException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { }), "drop");
        Assert.Throws<ArgumentOutOfRangeException>(() => list.Skip(-1));
    }

    [Fact]
    public void Contains_Item_ReturnsTrue()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3 } }), "drop");
        Assert.True(list.Contains(2));
    }

    [Fact]
    public void IEnumerable_Implementation_WorksCorrectly()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");

        var enumerator = list.GetEnumerator();
        var items = new List<object>();

        while (enumerator.MoveNext())
        {
            items.Add(enumerator.Current);
        }

        var expected = new object[] { 1, 2, 3, 4, 5 };
        Assert.Equal(expected, items.ToArray());
    }

    [Fact]
    public void GetAt_ValidIndex_ReturnsItem()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var item = list.GetAt(2);

        Assert.Equal(3, item);
    }

    [Fact]
    public void GetAt_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        Assert.Throws<ArgumentOutOfRangeException>(() => list.GetAt(5));
    }

    [Fact]
    public void Indexer_ValidIndex_ReturnsItem()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var item = list[2];

        Assert.Equal(3, item);
    }

    [Fact]
    public void Indexer_InvalidIndex_ThrowsArgumentOutOfRangeException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        Assert.Throws<ArgumentOutOfRangeException>(() => list[5]);
    }

    #endregion

    #region Sort tests

    [Fact]
    public void Sort_Descending_SortsCorrectly()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 3, 2 } }), "drop");
        list.Sort(true);
        var sorted = list.Get();
        Assert.Equal(new object[] { 3, 2, 1 }, sorted);
    }

    [Fact]
    public void Sort_Default_SortsAscending()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 3, 1, 2 } }), "drop");
        list.Sort(false);
        var sorted = list.Get();
        Assert.Equal(new object[] { 1, 2, 3 }, sorted);
    }

    [Fact]
    public void Sort_Strings_SortsAlphabetically()
    {
        var list = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { "banana", "apple", "cherry" } }), "string", "drop");
        list.Sort(false);
        var sorted = list.Get();
        Assert.Equal(new object[] { "apple", "banana", "cherry" }, sorted);
    }

    [Fact]
    public void Sort_Dates_SortsChronologically()
    {
        var date1 = new DateTime(2023, 1, 1);
        var date2 = new DateTime(2023, 5, 1);
        var date3 = new DateTime(2023, 3, 1);
        var list = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { date2, date1, date3 } }), "dateTime", "drop");
        list.Sort(false);
        var sorted = list.Get();
        Assert.Equal(new object[] { date1, date3, date2 }, sorted);
    }

    [Fact]
    public void Sort_MixedObjects_SortsByToString()
    {
        var date = new DateTime(2023, 1, 1);
        var mixedList = new object[,] { { "banana", date, 3.14, "apple" } };
        var list = XlBlockList.BuildTyped(XlBlockRange.Build(mixedList), "object", "drop");
        list.Sort(false);
        var sorted = list.Get();
        // Objects of different types will be sorted based on their ToString() representation
        Assert.Equal(new object[] { date, 3.14, "apple", "banana" }, sorted);
    }

    #endregion

    #region Misc tests

    [Fact]
    public void Copy_CreatesNewInstanceWithSameValues()
    {
        var original = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3 } }), "drop");
        var copy = original.Copy();

        Assert.NotSame(original, copy);
        Assert.Equal(original.Get(), copy.Get());
    }

    [Fact]
    public void Copy_DifferentReferenceForElements()
    {
        var original = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "a", 3.14 } }), "drop");
        var copy = original.Copy();

        Assert.NotSame(original, copy);

        // Ensure elements are the same
        var originalItems = original.Get();
        var copyItems = copy.Get();
        Assert.Equal(originalItems.Length, copyItems.Length);
        for (int i = 0; i < originalItems.Length; i++)
        {
            Assert.Equal(originalItems[i], copyItems[i]);
        }
    }

    [Fact]
    public void GetUniqueItems_ReturnsUniqueItems()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 2, 3, 4, 4, 5 } }), "drop");
        var uniqueList = list.GetUniqueItems();

        var expected = new object[] { 1, 2, 3, 4, 5 };
        Assert.Equal(expected, uniqueList.Get());
    }

    [Fact]
    public void GetDuplicateItems_ReturnsDuplicateItems()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 2, 3, 4, 4, 5 } }), "drop");
        var duplicateList = list.GetDuplicateItems();

        var expected = new object[] { 2, 4 };
        Assert.Equal(expected, duplicateList.Get());
    }

    [Fact]
    public void ContentsToString_CommaSeparator_ReturnsCommaSeparatedString()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.ContentsToString(",");

        var expected = "1,2,3,4,5";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ContentsToString_CustomSeparator_ReturnsCustomSeparatedString()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.ContentsToString("|");

        var expected = "1|2|3|4|5";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ContentsToString_EmptySeparator_ReturnsConcatenatedString()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.ContentsToString("");

        var expected = "12345";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ContentsToString_CommaSeparator_ReturnsCommaSeparatedString_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.ContentsToString(",");

        var expected = "1,apple,3.14,banana";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ContentsToString_CustomSeparator_ReturnsCustomSeparatedString_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.ContentsToString("|");

        var expected = "1|apple|3.14|banana";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ContentsToString_EmptySeparator_ReturnsConcatenatedString_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.ContentsToString("");

        var expected = "1apple3.14banana";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ByColumn_ReturnsColumnArray()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.AsArray(RangeOrientation.ByColumn);

        var expected = new object[,] { { 1 }, { 2 }, { 3 }, { 4 }, { 5 } };
        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ByRow_ReturnsRowArray()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.AsArray(RangeOrientation.ByRow);

        var expected = new object[,] { { 1, 2, 3, 4, 5 } };
        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ByColumn_ReturnsColumnArray_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.AsArray(RangeOrientation.ByColumn);

        var expected = new object[,] { { 1 }, { "apple" }, { 3.14 }, { "banana" } };
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToArray_ByRow_ReturnsRowArray_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.AsArray(RangeOrientation.ByRow);

        var expected = new object[,] { { 1, "apple", 3.14, "banana" } };
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Reverse_NumberTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, 2, 3, 4, 5 } }), "drop");
        var result = list.Reverse();

        var expected = new object[] { 5, 4, 3, 2, 1 };
        Assert.Equal(expected, result.Get());
    }

    [Fact]
    public void Reverse_MixedTypes()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 1, "apple", 3.14, "banana" } }), "drop");
        var result = list.Reverse();

        var expected = new object[] { "banana", 3.14, "apple", 1 };
        Assert.Equal(expected, result.Get());
    }

    #endregion

    #region Static method tests

    [Fact]
    public void MergeLists_DifferentTypes_ThrowsArgumentException()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { "string" } }), "string", "drop");
        Assert.Throws<ArgumentException>(() => XlBlockList.MergeLists(list1, list2));
    }

    [Fact]
    public void MergeLists_SameTypes_MergesLists()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1, 2, 3 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 2, 3, 4 } }), "int", "drop");
        var mergedList = XlBlockList.MergeLists(list1, list2);
        Assert.NotNull(mergedList);

        var expected = new object[] { 1, 2, 3, 4 };
        Assert.Equal(expected, mergedList.Get());
    }

    [Fact]
    public void UnifyLists_DifferentTypes_ThrowsArgumentException()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { "string" } }), "string", "drop");
        Assert.Throws<ArgumentException>(() => XlBlockList.UnifyLists(list1, list2));
    }

    [Fact]
    public void UnifyLists_SameTypes_UnifiesLists()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1, 2 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 2, 3 } }), "int", "drop");
        var unifiedList = XlBlockList.UnifyLists(list1, list2);
        Assert.NotNull(unifiedList);

        var expected = new object[] { 1, 2, 2, 3 };
        Assert.Equal(expected, unifiedList.Get());
    }

    [Fact]
    public void IntersectLists_DifferentTypes_ThrowsArgumentException()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { "string" } }), "string", "drop");
        Assert.Throws<ArgumentException>(() => XlBlockList.IntersectLists(list1, list2));
    }

    [Fact]
    public void IntersectLists_SameTypes_IntersectsLists()
    {
        var list1 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 1, 2, 3 } }), "int", "drop");
        var list2 = XlBlockList.BuildTyped(XlBlockRange.Build(new object[,] { { 2, 3, 4 } }), "int", "drop");
        var intersectedList = XlBlockList.IntersectLists(list1, list2);
        Assert.NotNull(intersectedList);

        var expected = new object[] { 2, 3 };
        Assert.Equal(expected, intersectedList.Get());
    }

    #endregion
}
