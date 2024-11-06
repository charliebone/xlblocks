namespace XlBlocks.AddIn.Types;

public interface IXlBlockCopyableObject<T>
{
    T Copy();
}

public enum RangeOrientation
{
    ByRow,
    ByColumn
}

public interface IXlBlockArrayableObject
{
    object[,] AsArray(RangeOrientation orientation);
}

public interface IElementCacher
{
    public string CollectionHexKey { get; }

    string CacheElement(object element, int index);
    string CacheElement(object element, string key);
}

public interface IXlBlockCacheableCollection
{
    IXlBlockCacheableCollection CacheCollectionValues(IElementCacher elementCacher);
    bool ShouldCacheElement(object element);
}
