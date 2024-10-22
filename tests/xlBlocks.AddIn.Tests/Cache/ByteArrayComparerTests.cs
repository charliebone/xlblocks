namespace XlBlocks.AddIn.Tests.Cache;

using XlBlocks.AddIn.Cache;

public class ByteArrayComparerTests
{
    [Fact]
    public void ByteArrayComparer_EqualByteArrays_ReturnsTrue()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        var byteArray2 = new byte[] { 1, 2, 3, 4 };

        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.True(result);
    }

    [Fact]
    public void ByteArrayComparer_DifferentByteArrays_ReturnsFalse()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        var byteArray2 = new byte[] { 4, 3, 2, 1 };

        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.False(result);
    }

    [Fact]
    public void ByteArrayComparer_NullByteArrays_ReturnsTrue()
    {
        var comparer = new ByteArrayComparer();
        byte[] byteArray1 = null!;
        byte[] byteArray2 = null!;

        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.True(result);
    }

    [Fact]
    public void ByteArrayComparer_OneNullByteArray_ReturnsFalse()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        byte[] byteArray2 = null!;

        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.False(result);
    }

    [Fact]
    public void ByteArrayComparer_DifferentLengthByteArrays_ReturnsFalse()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        var byteArray2 = new byte[] { 1, 2, 3 };

        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.False(result);
    }

    [Fact]
    public void ByteArrayComparer_EqualByteArrays_SameHashCode()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        var byteArray2 = new byte[] { 1, 2, 3, 4 };

        var hashCode1 = comparer.GetHashCode(byteArray1);
        var hashCode2 = comparer.GetHashCode(byteArray2);

        Assert.Equal(hashCode1, hashCode2);
    }

    [Fact]
    public void ByteArrayComparer_DifferentByteArrays_DifferentHashCode()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4 };
        var byteArray2 = new byte[] { 4, 3, 2, 1 };

        var hashCode1 = comparer.GetHashCode(byteArray1);
        var hashCode2 = comparer.GetHashCode(byteArray2);

        Assert.NotEqual(hashCode1, hashCode2);
    }

    [Fact]
    public void ByteArrayComparer_NullByteArray_ThrowsArgumentNullException()
    {
        var comparer = new ByteArrayComparer();
        byte[] byteArray = null!;

        Assert.Throws<ArgumentNullException>(() => comparer.GetHashCode(byteArray));
    }

    [Fact]
    public void ByteArrayComparer_SameHashCodeDifferentByteArrays_ReturnsFalse()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4, 5, 6 };
        var byteArray2 = new byte[] { 1, 2, 3, 4, 7, 8 };

        var hashCode1 = comparer.GetHashCode(byteArray1);
        var hashCode2 = comparer.GetHashCode(byteArray2);

        Assert.Equal(hashCode1, hashCode2);
        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.False(result);
    }

    [Fact]
    public void ByteArrayComparer_SameHashCodeDifferentLengthByteArrays_ReturnsFalse()
    {
        var comparer = new ByteArrayComparer();
        var byteArray1 = new byte[] { 1, 2, 3, 4, 5 };
        var byteArray2 = new byte[] { 1, 2, 3, 4, 5, 6, 7 };

        var hashCode1 = comparer.GetHashCode(byteArray1);
        var hashCode2 = comparer.GetHashCode(byteArray2);

        Assert.Equal(hashCode1, hashCode2);
        var result = comparer.Equals(byteArray1, byteArray2);
        Assert.False(result);
    }

}
