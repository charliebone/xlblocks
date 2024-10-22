namespace XlBlocks.AddIn.Tests.Types;

using System;
using System.Collections;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

public class XlBlockDictionaryTests
{
    #region Build tests

    [Fact]
    public void Build_UntypedKeysAndValues_StringValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { "value1" }, { "value2" }, { "value3" } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal("value1", dict["key1"]);
        Assert.Equal("value2", dict["key2"]);
        Assert.Equal("value3", dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeysAndValues_DoubleValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1.1 }, { 2.2 }, { 3.3 } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1.1, dict["key1"]);
        Assert.Equal(2.2, dict["key2"]);
        Assert.Equal(3.3, dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeys_StringValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { "value1" }, { "value2" }, { "value3" } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal("value1", dict["key1"]);
        Assert.Equal("value2", dict["key2"]);
        Assert.Equal("value3", dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeys_DoubleValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1.1 }, { 2.2 }, { 3.3 } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1.1, dict["key1"]);
        Assert.Equal(2.2, dict["key2"]);
        Assert.Equal(3.3, dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeys_DateValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { new DateTime(2020, 1, 1) }, { new DateTime(2021, 2, 2) }, { new DateTime(2022, 3, 3) } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(new DateTime(2020, 1, 1), dict["key1"]);
        Assert.Equal(new DateTime(2021, 2, 2), dict["key2"]);
        Assert.Equal(new DateTime(2022, 3, 3), dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeys_MixedValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { "value1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal("value1", dict["key1"]);
        Assert.Equal(2, dict["key2"]);
        Assert.Equal(new DateTime(2023, 1, 1), dict["key3"]);
    }

    [Fact]
    public void Build_UntypedKeys_MixedKeysAndValues()
    {
        object[,] keys = { { "key1" }, { "2" }, { "3.3" }, { new DateTime(2022, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { 3.3 }, { new DateTime(2023, 1, 1) } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error");

        Assert.Equal(4, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("2"));
        Assert.True(dict.ContainsKey("3.3"));
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.Equal("value1", dict["key1"]);
        Assert.Equal("two", dict["2"]);
        Assert.Equal(3.3, dict["3.3"]);
        Assert.Equal(new DateTime(2023, 1, 1), dict[new DateTime(2022, 1, 1)]);
    }

    [Fact]
    public void Build_OnErrors_Error_ThrowsException_MultipleInvalidKeys()
    {
        object[,] keys = { { ExcelError.ExcelErrorNA }, { "key2" }, { ExcelMissing.Value }, { ExcelEmpty.Value } };
        object[,] values = { { 1 }, { 2 }, { 3 }, { 4 } };

        Assert.Throws<ArgumentException>(() => XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void Build_OnErrors_Drop_Success_MultipleInvalidKeys()
    {
        object[,] keys = { { ExcelError.ExcelErrorNA }, { "key2" }, { ExcelMissing.Value }, { ExcelEmpty.Value } };
        object[,] values = { { 1 }, { 2 }, { 3 }, { 4 } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(1, dict.Count);  // Only "key2" should be added
        Assert.True(dict.ContainsKey("key2"));
        Assert.Equal(2, dict["key2"]);
    }

    [Fact]
    public void Build_OnErrors_Error_ThrowsException()
    {
        object[,] keys = { { "key1" }, { ExcelError.ExcelErrorNA }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };

        Assert.Throws<ArgumentException>(() => XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void Build_OnErrors_Drop_Success()
    {
        object[,] keys = { { "key1" }, { ExcelError.ExcelErrorNA }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.Build(XlBlockRange.Build(keys), XlBlockRange.Build(values), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(2, dict.Count);
        Assert.Equal(typeof(string), dict.KeyType);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal(3, dict["key3"]);
    }

    #endregion

    #region BuildTyped tests

    [Fact]
    public void BuildTyped_IntKeys_StringValues()
    {
        object[,] keys = { { "1" }, { "2" }, { "3" } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Int32", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(int), dict.KeyType);
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(2));
        Assert.True(dict.ContainsKey(3));
        Assert.Equal("one", dict[1]);
        Assert.Equal("two", dict[2]);
        Assert.Equal("three", dict[3]);
    }

    [Fact]
    public void BuildTyped_DoubleKeys_DateValues()
    {
        object[,] keys = { { "1.1" }, { "2.2" }, { "3.3" } };
        object[,] values = { { new DateTime(2020, 1, 1) }, { new DateTime(2021, 2, 2) }, { new DateTime(2022, 3, 3) } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Double", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.Equal(typeof(double), dict.KeyType);
        Assert.True(dict.ContainsKey(1.1));
        Assert.True(dict.ContainsKey(2.2));
        Assert.True(dict.ContainsKey(3.3));
        Assert.Equal(new DateTime(2020, 1, 1), dict[1.1]);
        Assert.Equal(new DateTime(2021, 2, 2), dict[2.2]);
        Assert.Equal(new DateTime(2022, 3, 3), dict[3.3]);
    }

    [Fact]
    public void BuildTyped_OnErrors_Error_ThrowsException()
    {
        object[,] keys = { { "1" }, { ExcelError.ExcelErrorNA }, { "3" } };
        object[,] values = { { "one" }, { "two" }, { "three" } };

        Assert.Throws<ArgumentException>(() => XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_OnErrors_Drop_Success()
    {
        object[,] keys = { { "1" }, { ExcelError.ExcelErrorNA }, { "3" } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int", XlBlockRange.Build(values), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(2, dict.Count);
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(3));
        Assert.Equal("one", dict[1]);
        Assert.Equal("three", dict[3]);
    }

    [Fact]
    public void BuildTyped_IntKeys_IntValues()
    {
        object[,] keys = { { "1" }, { "2" }, { "3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Int32", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(2));
        Assert.True(dict.ContainsKey(3));
        Assert.Equal(1, dict[1]);
        Assert.Equal(2, dict[2]);
        Assert.Equal(3, dict[3]);
    }

    [Fact]
    public void BuildTyped_DoubleKeys_StringValues()
    {
        object[,] keys = { { "1.1" }, { "2.2" }, { "3.3" } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Double", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(1.1));
        Assert.True(dict.ContainsKey(2.2));
        Assert.True(dict.ContainsKey(3.3));
        Assert.Equal("one", dict[1.1]);
        Assert.Equal("two", dict[2.2]);
        Assert.Equal("three", dict[3.3]);
    }

    [Fact]
    public void BuildTyped_DateTimeKeys_ObjectValues()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" }, { 42 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.DateTime", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2023, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2024, 1, 1)));
        Assert.Equal("New Year 2022", dict[new DateTime(2022, 1, 1)]);
        Assert.Equal("New Year 2023", dict[new DateTime(2023, 1, 1)]);
        Assert.Equal(42, dict[new DateTime(2024, 1, 1)]);
    }

    [Fact]
    public void BuildTyped_StringKeys_ObjectValues()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { "two" }, { new DateTime(2022, 1, 1) } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.String", XlBlockRange.Build(values), "error");

        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal("two", dict["key2"]);
        Assert.Equal(new DateTime(2022, 1, 1), dict["key3"]);
    }

    [Fact]
    public void BuildTyped_KeyCoercionFailure_IntKeys()
    {
        object[,] keys = { { "invalidKey" } };
        object[,] values = { { 1 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Int32", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_KeyCoercionFailure_DoubleKeys()
    {
        object[,] keys = { { "invalidKey" } };
        object[,] values = { { 1.1 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Double", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_KeyCoercionFailure_DateTimeKeys()
    {
        object[,] keys = { { "invalidKey" } };
        object[,] values = { { new DateTime(2022, 1, 1) } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.DateTime", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongIntKeys()
    {
        object[,] keys = { { "stringKey" } };
        object[,] values = { { 1 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongDoubleKeys()
    {
        object[,] keys = { { "stringKey" } };
        object[,] values = { { 1.1 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "System.Double", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongDateTimeKeys()
    {
        object[,] keys = { { "stringKey" } };
        object[,] values = { { new DateTime(2022, 1, 1) } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "DateTime", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongIntKeys_MixedTypes()
    {
        object[,] keys = { { "1" }, { "2.2" }, { "wrongKey" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "integer", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongDoubleKeys_MixedTypes()
    {
        object[,] keys = { { "1.1" }, { "2.2" }, { "wrongKey" } };
        object[,] values = { { 1.1 }, { 2.2 }, { 3.3 } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "Double", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_Exception_WrongDateTimeKeys_MixedTypes()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { "stringKey" }, { "wrongKey" } };
        object[,] values = { { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) }, { new DateTime(2025, 1, 1) } };
        Assert.Throws<ArgumentException>(() =>
            XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "date", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_OnErrors_Error_ThrowsException_WithMultipleInvalidKeys()
    {
        object[,] keys = { { "1" }, { ExcelError.ExcelErrorNA }, { ExcelMissing.Value }, { "4" } };
        object[,] values = { { "one" }, { "two" }, { "three" }, { "four" } };

        Assert.Throws<ArgumentException>(() => XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error"));
    }

    [Fact]
    public void BuildTyped_OnErrors_Drop_Success_WithMultipleInvalidKeys()
    {
        object[,] keys = { { "1" }, { ExcelError.ExcelErrorNA }, { ExcelMissing.Value }, { "4" } };
        object[,] values = { { "one" }, { "two" }, { "three" }, { "four" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(2, dict.Count);  // Only "1" and "4" should be added
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(4));
        Assert.Equal("one", dict[1]);
        Assert.Equal("four", dict[4]);
    }

    #endregion

    #region BuildFromLists tests

    [Fact]
    public void BuildFromLists_IntKeys_StringValues()
    {
        var keyList = new XlBlockList(new List<object> { 1, 2, 3 });
        var valueList = new XlBlockList(new List<object> { "one", "two", "three" });
        var dict = XlBlockDictionary.BuildFromLists(keyList, valueList);

        Assert.Equal(keyList.DataType, dict.KeyType);
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(2));
        Assert.True(dict.ContainsKey(3));
        Assert.Equal("one", dict[1]);
        Assert.Equal("two", dict[2]);
        Assert.Equal("three", dict[3]);
    }

    [Fact]
    public void BuildFromLists_DoubleKeys_DateValues()
    {
        var keyList = new XlBlockList(new List<object> { 1.1, 2.2, 3.3 });
        var valueList = new XlBlockList(new List<object> { new DateTime(2020, 1, 1), new DateTime(2021, 2, 2), new DateTime(2022, 3, 3) });
        var dict = XlBlockDictionary.BuildFromLists(keyList, valueList);

        Assert.Equal(keyList.DataType, dict.KeyType);
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(1.1));
        Assert.True(dict.ContainsKey(2.2));
        Assert.True(dict.ContainsKey(3.3));
        Assert.Equal(new DateTime(2020, 1, 1), dict[1.1]);
        Assert.Equal(new DateTime(2021, 2, 2), dict[2.2]);
        Assert.Equal(new DateTime(2022, 3, 3), dict[3.3]);
    }

    [Fact]
    public void BuildFromLists_DateTimeKeys_ObjectValues()
    {
        var keyList = new XlBlockList(new List<object> { new DateTime(2022, 1, 1), new DateTime(2023, 1, 1), new DateTime(2024, 1, 1) });
        var valueList = new XlBlockList(new List<object> { "New Year 2022", "New Year 2023", 42 });
        var dict = XlBlockDictionary.BuildFromLists(keyList, valueList);

        Assert.Equal(keyList.DataType, dict.KeyType);
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2023, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2024, 1, 1)));
        Assert.Equal("New Year 2022", dict[new DateTime(2022, 1, 1)]);
        Assert.Equal("New Year 2023", dict[new DateTime(2023, 1, 1)]);
        Assert.Equal(42, dict[new DateTime(2024, 1, 1)]);
    }

    [Fact]
    public void BuildFromLists_StringKeys_MixedValues()
    {
        var keyList = new XlBlockList(new List<string> { "key1", "key2", "key3" });
        var valueList = new XlBlockList(new List<object> { 1, "two", new DateTime(2022, 1, 1) });
        var dict = XlBlockDictionary.BuildFromLists(keyList, valueList);

        Assert.Equal(keyList.DataType, dict.KeyType);
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal("two", dict["key2"]);
        Assert.Equal(new DateTime(2022, 1, 1), dict["key3"]);
    }

    [Fact]
    public void BuildFromLists_ObjectKeys_MixedValues()
    {
        var keyList = new XlBlockList(new List<object> { "key1", 23, "key3", new DateTime(2022, 1, 1), "key5" });
        var valueList = new XlBlockList(new List<object> { 1, "two", new DateTime(2022, 1, 1), 23.45, null! });
        var dict = XlBlockDictionary.BuildFromLists(keyList, valueList);

        Assert.Equal(keyList.DataType, dict.KeyType);
        Assert.Equal(5, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey(23));
        Assert.True(dict.ContainsKey("key3"));
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.True(dict.ContainsKey("key5"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal("two", dict[23]);
        Assert.Equal(new DateTime(2022, 1, 1), dict["key3"]);
        Assert.Equal(23.45, dict[new DateTime(2022, 1, 1)]);
        Assert.Equal(null!, dict["key5"]);
    }

    #endregion

    #region AddOrReplace tests

    [Fact]
    public void AddOrReplace_IntKeys_Success()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(4, "four");
        Assert.Equal(4, dict.Count);
        Assert.True(dict.ContainsKey(4));
        Assert.Equal("four", dict[4]);
    }

    [Fact]
    public void AddOrReplace_IntKeys_ReplaceSuccess()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(2, "TWO");
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(2));
        Assert.Equal("TWO", dict[2]);
    }

    [Fact]
    public void AddOrReplace_StringKeys_Success()
    {
        object[,] keys = { { "key1" }, { "key2" } };
        object[,] values = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.AddOrReplace("key3", 3);
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(3, dict["key3"]);
    }

    [Fact]
    public void AddOrReplace_StringKeys_ReplaceSuccess()
    {
        object[,] keys = { { "key1" }, { "key2" } };
        object[,] values = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "string", XlBlockRange.Build(values), "error");

        dict.AddOrReplace("key2", 22);
        Assert.Equal(2, dict.Count);
        Assert.True(dict.ContainsKey("key2"));
        Assert.Equal(22, dict["key2"]);
    }

    [Fact]
    public void AddOrReplace_DateTimeKeys_Success()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(new DateTime(2024, 1, 1), "New Year 2024");
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey(new DateTime(2024, 1, 1)));
        Assert.Equal("New Year 2024", dict[new DateTime(2024, 1, 1)]);
    }

    [Fact]
    public void AddOrReplace_DateTimeKeys_ReplaceSuccess()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "date", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(new DateTime(2022, 1, 1), "New Year 2022 Updated");
        Assert.Equal(2, dict.Count);
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.Equal("New Year 2022 Updated", dict[new DateTime(2022, 1, 1)]);
    }

    [Fact]
    public void AddOrReplace_IntKeys_Exception_WrongKeyType()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "integer", XlBlockRange.Build(values), "error");

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace("wrongKey", "value"));
    }

    [Fact]
    public void AddOrReplace_DateTimeKeys_Exception_WrongKeyType()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace("wrongKey", "value"));
    }

    [Fact]
    public void AddOrReplace_StringKeys_CoerceIntKey_Success()
    {
        object[,] keys = { { "key1" }, { "key2" } };
        object[,] values = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(3, "value");
        Assert.Equal("value", dict["3"]);
    }

    [Fact]
    public void AddOrReplace_StringKeys_CoerceDoubleKey_Success()
    {
        object[,] keys = { { "key1" }, { "key2" } };
        object[,] values = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(3.3, "value");
        Assert.Equal("value", dict["3.3"]);
    }

    [Fact]
    public void AddOrReplace_StringKeys_CoerceDateTimeKey_Success()
    {
        object[,] keys = { { "key1" }, { "key2" } };
        object[,] values = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.AddOrReplace(new DateTime(2023, 1, 1), "date_value");
        Assert.Equal("date_value", dict[new DateTime(2023, 1, 1)]);
    }

    [Fact]
    public void AddOrReplace_IntKeys_CoerceStringKey_Success()
    {
        object[,] keys = { { 1 }, { 2 } };
        object[,] values = { { "one" }, { "two" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        dict.AddOrReplace("3", "three");
        Assert.Equal("three", dict[3]);
    }

    [Fact]
    public void AddOrReplace_IntKeys_CoerceInvalidKey_Exception()
    {
        object[,] keys = { { 1 }, { 2 } };
        object[,] values = { { "one" }, { "two" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace("invalid", "value"));
    }

    [Fact]
    public void AddOrReplace_DoubleKeys_CoerceStringKey_Success()
    {
        object[,] keys = { { 1.1 }, { 2.2 } };
        object[,] values = { { "one.one" }, { "two.two" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "double", XlBlockRange.Build(values), "error");

        dict.AddOrReplace("3.3", "three.three");
        Assert.Equal("three.three", dict[3.3]);
    }

    [Fact]
    public void AddOrReplace_DoubleKeys_CoerceInvalidKey_Exception()
    {
        object[,] keys = { { 1.1 }, { 2.2 } };
        object[,] values = { { "one.one" }, { "two.two" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "double", XlBlockRange.Build(values), "error");

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace("invalid", "value"));
    }

    [Fact]
    public void AddOrReplace_OnErrors_Error_ThrowsException_WithInitializedDict()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with errors
        object[,] errorKeys = { { ExcelError.ExcelErrorNA }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 } };

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "error"));
    }

    [Fact]
    public void AddOrReplace_OnErrors_Drop_Success_WithInitializedDict()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with errors
        object[,] errorKeys = { { ExcelError.ExcelErrorNA }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 } };

        dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal(2, dict["key2"]);
        Assert.Equal(4, dict["key3"]);
    }

    [Fact]
    public void AddOrReplace_OnErrors_Error_ThrowsException_WithMultipleInvalidKeys()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with errors and missing values
        object[,] errorKeys = { { ExcelError.ExcelErrorNA }, { ExcelMissing.Value }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 }, { 5 } };

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "error"));
    }

    [Fact]
    public void AddOrReplace_OnErrors_Drop_Success_WithMultipleInvalidKeys()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with errors and missing values
        object[,] errorKeys = { { ExcelError.ExcelErrorNA }, { ExcelMissing.Value }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 }, { 5 } };

        dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal(2, dict["key2"]);
        Assert.Equal(5, dict["key3"]);
    }

    [Fact]
    public void AddOrReplace_OnErrors_Error_ThrowsException_WithEmptyValues()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with empty values
        object[,] errorKeys = { { ExcelEmpty.Value }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 } };

        Assert.Throws<ArgumentException>(() => dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "error"));
    }

    [Fact]
    public void AddOrReplace_OnErrors_Drop_Success_WithEmptyValues()
    {
        object[,] initialKeys = { { "key1" }, { "key2" } };
        object[,] initialValues = { { 1 }, { 2 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(initialKeys), "str", XlBlockRange.Build(initialValues), "error");

        // Keys with empty values
        object[,] errorKeys = { { ExcelEmpty.Value }, { "key3" } };
        object[,] errorValues = { { 3 }, { 4 } };

        dict.AddOrReplace(XlBlockRange.Build(errorKeys), XlBlockRange.Build(errorValues), "drop");

        // Ensure only valid key-value pairs are added
        Assert.Equal(3, dict.Count);
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key2"));
        Assert.True(dict.ContainsKey("key3"));
        Assert.Equal(1, dict["key1"]);
        Assert.Equal(2, dict["key2"]);
        Assert.Equal(4, dict["key3"]);
    }

    #endregion

    #region Remove tests

    [Fact]
    public void Remove_IntKeys_Success()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        dict.Remove(2);
        Assert.False(dict.ContainsKey(2));
        Assert.Equal(2, dict.Count);
    }

    [Fact]
    public void Remove_IntKeys_NonExistentKey_SilentFail()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        dict.Remove(4);
        Assert.Equal(3, dict.Count);  // No change in count
    }

    [Fact]
    public void Remove_StringKeys_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.Remove("key2");
        Assert.False(dict.ContainsKey("key2"));
        Assert.Equal(2, dict.Count);
    }

    [Fact]
    public void Remove_StringKeys_NonExistentKey_SilentFail()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.Remove("key4");
        Assert.Equal(3, dict.Count);  // No change in count
    }

    [Fact]
    public void Remove_DateTimeKeys_Success()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" }, { "New Year 2024" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        dict.Remove(new DateTime(2023, 1, 1));
        Assert.False(dict.ContainsKey(new DateTime(2023, 1, 1)));
        Assert.Equal(2, dict.Count);
    }

    [Fact]
    public void Remove_DateTimeKeys_NonExistentKey_SilentFail()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" }, { "New Year 2024" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        dict.Remove(new DateTime(2025, 1, 1));
        Assert.Equal(3, dict.Count);  // No change in count
    }

    [Fact]
    public void Remove_StringKeys_CoerceIntKey_SilentFail()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        dict.Remove(3);  // 3 coerced to "3" which doesn't match any key
        Assert.Equal(3, dict.Count);  // No change in count
    }

    [Fact]
    public void Remove_IntKeys_CoerceStringKey_SilentFail()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        dict.Remove("4");  // "4" can't be coerced to int
        Assert.Equal(3, dict.Count);  // No change in count
    }

    [Fact]
    public void Remove_MultipleIntKeys_Success()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 }, { 4 } };
        object[,] values = { { "one" }, { "two" }, { "three" }, { "four" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        object[,] keysToRemove = { { 2 }, { 4 } };
        dict.Remove(XlBlockRange.Build(keysToRemove));

        Assert.Equal(2, dict.Count);
        Assert.False(dict.ContainsKey(2));
        Assert.False(dict.ContainsKey(4));
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(3));
    }

    [Fact]
    public void Remove_MultipleStringKeys_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" }, { "key4" } };
        object[,] values = { { 1 }, { 2 }, { 3 }, { 4 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        object[,] keysToRemove = { { "key2" }, { "key4" } };
        dict.Remove(XlBlockRange.Build(keysToRemove));

        Assert.Equal(2, dict.Count);
        Assert.False(dict.ContainsKey("key2"));
        Assert.False(dict.ContainsKey("key4"));
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey("key3"));
    }

    [Fact]
    public void Remove_MultipleDateTimeKeys_Success()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) }, { new DateTime(2025, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" }, { "New Year 2024" }, { "New Year 2025" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        object[,] keysToRemove = { { new DateTime(2023, 1, 1) }, { new DateTime(2025, 1, 1) } };
        dict.Remove(XlBlockRange.Build(keysToRemove));

        Assert.Equal(2, dict.Count);
        Assert.False(dict.ContainsKey(new DateTime(2023, 1, 1)));
        Assert.False(dict.ContainsKey(new DateTime(2025, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2022, 1, 1)));
        Assert.True(dict.ContainsKey(new DateTime(2024, 1, 1)));
    }

    [Fact]
    public void Remove_MultipleMixedKeys_Success()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) }, { "key4" } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" }, { "value4" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        object[,] keysToRemove = { { 2 }, { "key4" } };
        dict.Remove(XlBlockRange.Build(keysToRemove));

        Assert.Equal(2, dict.Count);
        Assert.False(dict.ContainsKey(2));
        Assert.False(dict.ContainsKey("key4"));
        Assert.True(dict.ContainsKey("key1"));
        Assert.True(dict.ContainsKey(new DateTime(2023, 1, 1)));
    }

    #endregion

    #region Copy tests

    [Fact]
    public void Copy_IntKeys_Success()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var originalDict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        var copiedDict = originalDict.Copy();

        // Check that the copied dictionary is equal but not the same reference
        Assert.NotSame(originalDict, copiedDict);
        Assert.Equal(originalDict.Count, copiedDict.Count);
        Assert.Equal(originalDict.KeyType, copiedDict.KeyType);

        foreach (var key in originalDict.Keys)
        {
            Assert.True(copiedDict.ContainsKey(key));
            Assert.Equal(originalDict[key], copiedDict[key]);
        }
    }

    [Fact]
    public void Copy_StringKeys_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var originalDict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var copiedDict = originalDict.Copy();

        // Check that the copied dictionary is equal but not the same reference
        Assert.NotSame(originalDict, copiedDict);
        Assert.Equal(originalDict.Count, copiedDict.Count);
        Assert.Equal(originalDict.KeyType, copiedDict.KeyType);

        foreach (var key in originalDict.Keys)
        {
            Assert.True(copiedDict.ContainsKey(key));
            Assert.Equal(originalDict[key], copiedDict[key]);
        }
    }

    [Fact]
    public void Copy_DateTimeKeys_Success()
    {
        object[,] keys = { { new DateTime(2022, 1, 1) }, { new DateTime(2023, 1, 1) }, { new DateTime(2024, 1, 1) } };
        object[,] values = { { "New Year 2022" }, { "New Year 2023" }, { "New Year 2024" } };
        var originalDict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "datetime", XlBlockRange.Build(values), "error");

        var copiedDict = originalDict.Copy();

        // Check that the copied dictionary is equal but not the same reference
        Assert.NotSame(originalDict, copiedDict);
        Assert.Equal(originalDict.Count, copiedDict.Count);
        Assert.Equal(originalDict.KeyType, copiedDict.KeyType);

        foreach (var key in originalDict.Keys)
        {
            Assert.True(copiedDict.ContainsKey(key));
            Assert.Equal(originalDict[key], copiedDict[key]);
        }
    }

    [Fact]
    public void Copy_MixedKeys_Success()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var originalDict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var copiedDict = originalDict.Copy();

        // Check that the copied dictionary is equal but not the same reference
        Assert.NotSame(originalDict, copiedDict);
        Assert.Equal(originalDict.Count, copiedDict.Count);
        Assert.Equal(originalDict.KeyType, copiedDict.KeyType);

        foreach (var key in originalDict.Keys)
        {
            Assert.True(copiedDict.ContainsKey(key));
            Assert.Equal(originalDict[key], copiedDict[key]);
        }
    }

    #endregion

    #region AsArray tests

    [Fact]
    public void AsArray_ByColumn_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(RangeOrientation.ByColumn);

        Assert.Equal(3, array.GetLength(0));  // Number of rows
        Assert.Equal(2, array.GetLength(1));  // Number of columns
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal(1, array[0, 1]);
        Assert.Equal("key2", array[1, 0]);
        Assert.Equal(2, array[1, 1]);
        Assert.Equal("key3", array[2, 0]);
        Assert.Equal(3, array[2, 1]);
    }

    [Fact]
    public void AsArray_ByRow_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(RangeOrientation.ByRow);

        Assert.Equal(2, array.GetLength(0));  // Number of rows
        Assert.Equal(3, array.GetLength(1));  // Number of columns
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal("key2", array[0, 1]);
        Assert.Equal("key3", array[0, 2]);
        Assert.Equal(1, array[1, 0]);
        Assert.Equal(2, array[1, 1]);
        Assert.Equal(3, array[1, 2]);
    }

    [Fact]
    public void AsArray_ByColumn_Success_MixedTypes()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(RangeOrientation.ByColumn);

        Assert.Equal(3, array.GetLength(0));  // Number of rows
        Assert.Equal(2, array.GetLength(1));  // Number of columns
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal("value1", array[0, 1]);
        Assert.Equal(2, array[1, 0]);
        Assert.Equal("two", array[1, 1]);
        Assert.Equal(new DateTime(2023, 1, 1), array[2, 0]);
        Assert.Equal("New Year 2023", array[2, 1]);
    }

    [Fact]
    public void AsArray_ByRow_Success_MixedTypes()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(RangeOrientation.ByRow);

        Assert.Equal(2, array.GetLength(0));  // Number of rows
        Assert.Equal(3, array.GetLength(1));  // Number of columns
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal(2, array[0, 1]);
        Assert.Equal(new DateTime(2023, 1, 1), array[0, 2]);
        Assert.Equal("value1", array[1, 0]);
        Assert.Equal("two", array[1, 1]);
        Assert.Equal("New Year 2023", array[1, 2]);
    }

    [Fact]
    public void AsArray_WithHeader_ByColumn_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(true, "KeyName", "ValueName", RangeOrientation.ByColumn);

        // Ensure array dimensions (including headers)
        Assert.Equal(4, array.GetLength(0));  // 3 entries + 1 header row
        Assert.Equal(2, array.GetLength(1));  // 2 columns

        // Check headers
        Assert.Equal("KeyName", array[0, 0]);
        Assert.Equal("ValueName", array[0, 1]);

        // Check key-value pairs
        Assert.Equal("key1", array[1, 0]);
        Assert.Equal(1, array[1, 1]);
        Assert.Equal("key2", array[2, 0]);
        Assert.Equal(2, array[2, 1]);
        Assert.Equal("key3", array[3, 0]);
        Assert.Equal(3, array[3, 1]);
    }

    [Fact]
    public void AsArray_WithHeader_ByRow_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(true, "KeyName", "ValueName", RangeOrientation.ByRow);

        // Ensure array dimensions (including headers)
        Assert.Equal(2, array.GetLength(0));  // 2 rows
        Assert.Equal(4, array.GetLength(1));  // 3 entries + 1 header column

        // Check headers
        Assert.Equal("KeyName", array[0, 0]);
        Assert.Equal("ValueName", array[1, 0]);

        // Check key-value pairs
        Assert.Equal("key1", array[0, 1]);
        Assert.Equal("key2", array[0, 2]);
        Assert.Equal("key3", array[0, 3]);
        Assert.Equal(1, array[1, 1]);
        Assert.Equal(2, array[1, 2]);
        Assert.Equal(3, array[1, 3]);
    }

    [Fact]
    public void AsArray_WithoutHeader_ByColumn_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(false, "KeyName", "ValueName", RangeOrientation.ByColumn);

        // Ensure array dimensions
        Assert.Equal(3, array.GetLength(0));  // 3 entries
        Assert.Equal(2, array.GetLength(1));  // 2 columns

        // Check key-value pairs
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal(1, array[0, 1]);
        Assert.Equal("key2", array[1, 0]);
        Assert.Equal(2, array[1, 1]);
        Assert.Equal("key3", array[2, 0]);
        Assert.Equal(3, array[2, 1]);
    }

    [Fact]
    public void AsArray_WithoutHeader_ByRow_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var array = dict.AsArray(false, "KeyName", "ValueName", RangeOrientation.ByRow);

        // Ensure array dimensions
        Assert.Equal(2, array.GetLength(0));  // 2 rows
        Assert.Equal(3, array.GetLength(1));  // 3 entries

        // Check key-value pairs
        Assert.Equal("key1", array[0, 0]);
        Assert.Equal("key2", array[0, 1]);
        Assert.Equal("key3", array[0, 2]);
        Assert.Equal(1, array[1, 0]);
        Assert.Equal(2, array[1, 1]);
        Assert.Equal(3, array[1, 2]);
    }

    #endregion

    #region IEnumerable and keys and values properties

    [Fact]
    public void GetEnumerator_IntKeys_Success()
    {
        object[,] keys = { { 1 }, { 2 }, { 3 } };
        object[,] values = { { "one" }, { "two" }, { "three" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "int32", XlBlockRange.Build(values), "error");

        foreach (DictionaryEntry entry in dict)
        {
            Assert.True(dict.ContainsKey(entry.Key));
            Assert.Equal(dict[entry.Key], entry.Value);
        }
    }

    [Fact]
    public void GetEnumerator_StringKeys_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var enumerator = dict.GetEnumerator();
        while (enumerator.MoveNext())
        {
            var entry = (DictionaryEntry)enumerator.Current;
            Assert.True(dict.ContainsKey(entry.Key));
            Assert.Equal(dict[entry.Key], entry.Value);
        }
    }

    [Fact]
    public void GetEnumerator_MixedTypes_Success()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var enumerator = dict.GetEnumerator();
        while (enumerator.MoveNext())
        {
            var entry = (DictionaryEntry)enumerator.Current;
            Assert.True(dict.ContainsKey(entry.Key));
            Assert.Equal(dict[entry.Key], entry.Value);
        }
    }

    [Fact]
    public void Keys_Property_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var dictKeys = dict.Keys;
        Assert.Equal(3, dictKeys.Length);
        Assert.Contains("key1", dictKeys);
        Assert.Contains("key2", dictKeys);
        Assert.Contains("key3", dictKeys);
    }

    [Fact]
    public void Values_Property_Success()
    {
        object[,] keys = { { "key1" }, { "key2" }, { "key3" } };
        object[,] values = { { 1 }, { 2 }, { 3 } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "str", XlBlockRange.Build(values), "error");

        var dictValues = dict.Values;
        Assert.Equal(3, dictValues.Length);
        Assert.Contains(1, dictValues);
        Assert.Contains(2, dictValues);
        Assert.Contains(3, dictValues);
    }

    [Fact]
    public void Keys_Property_MixedTypes_Success()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var dictKeys = dict.Keys;
        Assert.Equal(3, dictKeys.Length);
        Assert.Contains("key1", dictKeys);
        Assert.Contains(2, dictKeys);
        Assert.Contains(new DateTime(2023, 1, 1), dictKeys);
    }

    [Fact]
    public void Values_Property_MixedTypes_Success()
    {
        object[,] keys = { { "key1" }, { 2 }, { new DateTime(2023, 1, 1) } };
        object[,] values = { { "value1" }, { "two" }, { "New Year 2023" } };
        var dict = XlBlockDictionary.BuildTyped(XlBlockRange.Build(keys), "object", XlBlockRange.Build(values), "error");

        var dictValues = dict.Values;
        Assert.Equal(3, dictValues.Length);
        Assert.Contains("value1", dictValues);
        Assert.Contains("two", dictValues);
        Assert.Contains("New Year 2023", dictValues);
    }

    #endregion
}
