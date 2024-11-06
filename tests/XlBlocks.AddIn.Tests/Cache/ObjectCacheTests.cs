namespace XlBlocks.AddIn.Tests.Cache;

using System.Collections.Generic;
using XlBlocks.AddIn.Cache;
using XlBlocks.AddIn.Types;

public class ObjectCacheTests
{
    [Fact]
    public void GetCacheKey_ValidReferences_GeneratesCorrectHexKey()
    {
        var cacheKeyVals = new Dictionary<string, string>
        {
            ["Sheet1!A1"] = "@2dc1c006719dabc0b2a311e75cc26f11@",
            ["Sheet2!$B$5"] = "@298ddb4bd41c14a5f0672a989afe279a@",
            ["AnotherSheet!F7"] = "@e9a3142065b04349727118a6baf54319@"
        };

        foreach (var cacheKey in cacheKeyVals)
        {
            var _ = ObjectCache.GetCacheKey(cacheKey.Key, out var hex);
            Assert.Equal(cacheKey.Value, hex);
        }
    }

    [Fact]
    public void GetCacheCollectionKey_ValidReferencesWithIndex_GeneratesCorrectHexKey()
    {
        var cacheKeyVals = new Dictionary<string, string>
        {
            ["Sheet1!A1_123"] = "@2dc1c006719dabc0b2a311e75cc26f11@_@b5c26b422c0ba9342c0ba9342c0ba934@",
            ["Sheet2!$B$5_456"] = "@298ddb4bd41c14a5f0672a989afe279a@_@9c5137e882eb4fff82eb4fff82eb4fff@",
            ["AnotherSheet!F7_42"] = "@e9a3142065b04349727118a6baf54319@_@e0710fe233bb23ba33bb23ba33bb23ba@"
        };

        foreach (var cacheKey in cacheKeyVals)
        {
            var keyParts = cacheKey.Key.Split("_");
            var referenceHash = ObjectCache.GetCacheKey(keyParts[0], out var referenceHex);
            var _ = ObjectCache.GetCacheCollectionKey(referenceHash, referenceHex, int.Parse(keyParts[1]), out var hex);
            Assert.Equal(cacheKey.Value, hex);
        }
    }

    [Fact]
    public void GetCacheCollectionKey_ValidReferencesWithStringKey_GeneratesCorrectHexKey()
    {
        var cacheKeyVals = new Dictionary<string, string>
        {
            ["Sheet1!A1_dictkey"] = "@2dc1c006719dabc0b2a311e75cc26f11@_@4344600f2ac46b6cfa76478cfa76478c@",
            ["Sheet2!$B$5_anotherkey"] = "@298ddb4bd41c14a5f0672a989afe279a@_@49019db3ac784397aa8cf5a5e82ed1a0@",
            ["AnotherSheet!F7_asdf"] = "@e9a3142065b04349727118a6baf54319@_@e6c8515a1d70c9981d70c9981d70c998@"
        };

        foreach (var cacheKey in cacheKeyVals)
        {
            var keyParts = cacheKey.Key.Split("_");
            var referenceHash = ObjectCache.GetCacheKey(keyParts[0], out var referenceHex);
            var _ = ObjectCache.GetCacheCollectionKey(referenceHash, referenceHex, keyParts[1], out var hex);
            Assert.Equal(cacheKey.Value, hex);
        }
    }

    [Fact]
    public void AddOrReplace_ValueTypeReference_AddsAndRetrievesSuccessfully()
    {
        var reference = "Sheet1!A1";
        var value = 123;
        var expectedHex = "@2dc1c006719dabc0b2a311e75cc26f11@";

        var cache = new ObjectCache();
        cache.AddOrReplace(reference, value, out var hexKey);
        Assert.Equal(expectedHex, hexKey);

        Assert.True(cache.ContainsKey(hexKey));

        var success = cache.TryGetValue(hexKey, out var retrievedValue);
        Assert.True(success);
        Assert.Equal(value, retrievedValue);

        success = cache.TryGetReference(hexKey, out var retrievedReference);
        Assert.True(success);
        Assert.Equal(reference, retrievedReference);
    }

    [Fact]
    public void AddOrReplace_ReferenceTypeReference_AddsAndRetrievesSuccessfully()
    {
        var reference = "Sheet1!A1";
        var value = "i am a string";
        var expectedHex = "@2dc1c006719dabc0b2a311e75cc26f11@";

        var cache = new ObjectCache();
        cache.AddOrReplace(reference, value, out var hexKey);
        Assert.Equal(expectedHex, hexKey);

        Assert.True(cache.ContainsKey(hexKey));

        var success = cache.TryGetValue(hexKey, out var retrievedValue);
        Assert.True(success);
        Assert.Equal(value, retrievedValue);

        success = cache.TryGetReference(hexKey, out var retrievedReference);
        Assert.True(success);
        Assert.Equal(reference, retrievedReference);
    }

    [Fact]
    public void AddOrReplace_CacheableCollection_IntKey()
    {
        var reference = "SheetName!C3";
        var listToCache = new List<object> { "item1", new List<double> { 2, 3, 4, 5 }, EventArgs.Empty };

        var cache = new ObjectCache();
        var xlList = new XlBlockList(listToCache);
        var elementCacher = cache.GetElementCacher(reference);
        var collection = xlList.CacheCollectionValues(elementCacher) as XlBlockList;
        Assert.NotNull(collection);

        for (var i = 0; i < listToCache.Count; i++) // originalItem in listToCache)
        {
            var originalItem = listToCache[i];
            if (xlList.ShouldCacheElement(originalItem))
            {
                var elementCacheKey = collection[i]?.ToString();
                Assert.NotNull(elementCacheKey);
                Assert.True(cache.TryGetValue(elementCacheKey, out var cachedItem));
                Assert.Equal(originalItem, cachedItem);
            }
            else
            {
                Assert.Equal(originalItem, collection[i]);
            }
        }
    }

    [Fact]
    public void AddOrReplace_CacheableCollection_StringKey()
    {
        var reference = "Sheet1!A1";
        var dictToCache = new Dictionary<string, object>
        {
            ["key1"] = "some value",
            ["another key"] = new List<double> { 2, 3, 4, 5 },
            ["third_key"] = EventArgs.Empty
        };

        var cache = new ObjectCache();
        var xlDictionary = new XlBlockDictionary(dictToCache, typeof(string));
        var elementCacher = cache.GetElementCacher(reference);
        var collection = xlDictionary.CacheCollectionValues(elementCacher) as XlBlockDictionary;
        Assert.NotNull(collection);

        foreach (var originalItem in dictToCache)
        {
            if (xlDictionary.ShouldCacheElement(originalItem.Value))
            {
                var elementCacheKey = collection[originalItem.Key]?.ToString();
                Assert.NotNull(elementCacheKey);
                Assert.True(cache.TryGetValue(elementCacheKey, out var cachedItem));
                Assert.Equal(originalItem.Value, cachedItem);
            }
            else
            {
                Assert.Equal(originalItem.Value, collection[originalItem.Key]);
            }
        }
    }

    [Fact]
    public void Remove_ValidHexKey_RemovesSuccessfully()
    {
        var key = "Sheet1!A1";
        var value = "i am a string";
        var expectedHex = "@2dc1c006719dabc0b2a311e75cc26f11@";

        var cache = new ObjectCache();
        cache.AddOrReplace(key, value, out var hexKey);
        Assert.Equal(expectedHex, hexKey);

        Assert.True(cache.ContainsKey(hexKey));

        var success = cache.Remove(hexKey, out var removedValue);
        Assert.True(success);
        Assert.Equal(value, removedValue);
    }

    [Fact]
    public void Clear_ClearsAllItemsSuccessfully()
    {
        var cache = new ObjectCache();
        cache.AddOrReplace("key1", 1, out var _);
        cache.AddOrReplace("key2", "value", out var _);
        cache.AddOrReplace("key3", new { a = "some value" }, out var _);

        Assert.Equal(3, cache.Count);

        cache.Clear();
        Assert.Equal(0, cache.Count);
    }

    [Fact]
    public void AddOrReplace_DuplicateReference_ReplacesValue()
    {
        var reference = "Sheet1!A1";
        var initialValue = 123;
        var newValue = 456;
        var expectedHex = "@2dc1c006719dabc0b2a311e75cc26f11@";

        var cache = new ObjectCache();
        cache.AddOrReplace(reference, initialValue, out var hexKey);
        cache.AddOrReplace(reference, newValue, out var newHexKey);

        Assert.Equal(expectedHex, hexKey);
        Assert.Equal(hexKey, newHexKey);

        Assert.True(cache.ContainsKey(hexKey));

        var success = cache.TryGetValue(hexKey, out var retrievedValue);
        Assert.True(success);
        Assert.Equal(newValue, retrievedValue);
    }

    [Fact]
    public void Remove_NonExistentKey_ReturnsFalse()
    {
        var cache = new ObjectCache();
        var nonExistentHexKey = "@abcdefabcdefabcdefabcdefabcdefab@";

        var success = cache.Remove(nonExistentHexKey, out var removedValue);
        Assert.False(success);
        Assert.Null(removedValue);
    }

    [Fact]
    public void CacheInvalidatedEvent_AddOrReplace_RaisesEvent()
    {
        var cache = new ObjectCache();
        bool eventRaised = false;

        cache.CacheInvalidated += (sender, e) => { eventRaised = true; };

        cache.AddOrReplace("Sheet1!A1", 123, out var _);
        Assert.True(eventRaised);
    }

    [Fact]
    public void CacheInvalidatedEvent_Remove_RaisesEvent()
    {
        var cache = new ObjectCache();
        bool eventRaised = false;

        cache.AddOrReplace("Sheet1!A1", 123, out var hexKey);

        cache.CacheInvalidated += (sender, e) => { eventRaised = true; };

        cache.Remove(hexKey);
        Assert.True(eventRaised);
    }

    [Fact]
    public void GetItemInfo_SingleItem_ReturnsCorrectInfo()
    {
        var reference = "Sheet1!A1";
        var value = 123;
        var expectedHex = "@2dc1c006719dabc0b2a311e75cc26f11@";

        var cache = new ObjectCache();
        cache.AddOrReplace(reference, value, out _);

        var itemInfo = cache.GetItemInfo();

        Assert.Single(itemInfo);
        Assert.Equal(reference, itemInfo[0].Reference);
        Assert.Equal(expectedHex, itemInfo[0].HexKey);
        Assert.Equal(typeof(int), itemInfo[0].Type);
    }

    [Fact]
    public void GetItemInfo_MultipleItems_ReturnsCorrectInfo()
    {
        var items = new Dictionary<string, object>
        {
            { "Sheet1!A1", 123 },
            { "Sheet2!$B$5", "test" },
            { "AnotherSheet!F7", new DateTime(2023, 1, 1) }
        };

        var expectedHexes = new Dictionary<string, string>
        {
            ["Sheet1!A1"] = "@2dc1c006719dabc0b2a311e75cc26f11@",
            ["Sheet2!$B$5"] = "@298ddb4bd41c14a5f0672a989afe279a@",
            ["AnotherSheet!F7"] = "@e9a3142065b04349727118a6baf54319@"
        };

        var cache = new ObjectCache();

        int index = 0;
        foreach (var item in items)
        {
            cache.AddOrReplace(item.Key, item.Value, out var hexKey);
            Assert.Equal(expectedHexes[item.Key], hexKey);
            index++;
        }

        var itemInfo = cache.GetItemInfo();
        Assert.Equal(3, itemInfo.Count);

        foreach (var info in itemInfo)
        {
            Assert.Equal(info.HexKey, expectedHexes[info.Reference]);
            Assert.Equal(info.Type, items[info.Reference].GetType());
        }
    }
}
