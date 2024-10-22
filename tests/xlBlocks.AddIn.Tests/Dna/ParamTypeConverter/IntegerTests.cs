namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using System.Globalization;
using XlBlocks.AddIn.Dna;

public class IntegerTests
{
    [Fact]
    public void TryConvertToInt32_ValidInteger_ReturnsTrue()
    {
        object input = 42;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToInt32_ValidStringNumber_ReturnsTrue()
    {
        object input = "42";
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToInt32_ValidDouble_ReturnsTrue()
    {
        object input = 42.5;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToInt32_NonNumericString_ReturnsFalse()
    {
        object input = "abc";
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToInt32_NullObject_ReturnsFalse()
    {
        object input = null!;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToInt32_VeryLargeDouble_ReturnsFalse()
    {
        object input = double.MaxValue;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToInt32_StringWithCommaDifferentCulture_ReturnsTrue()
    {
        var number = 1234567;
        var frenchCulture = new CultureInfo("fr-FR");
        object input = number.ToString("N0", frenchCulture);
        Thread.CurrentThread.CurrentCulture = frenchCulture;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToInt32_StringWithThousandSeparator_ReturnsTrue()
    {
        int number = 1234567;
        object input = number.ToString("N0", CultureInfo.CurrentCulture); // Format with thousands separator
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToInt32_StringWithDecimalPoint_ReturnsTrue()
    {
        object input = "1234.567";
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(1235, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingUp_ReturnsTrue()
    {
        object input = 42.6;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(43, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingDown_ReturnsTrue()
    {
        object input = 42.4;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingToEven_RoundDown_ReturnsTrue()
    {
        object input = 42.5;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingToEven_RoundUp_ReturnsTrue()
    {
        object input = 43.5;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(44, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingUpBoundary_ReturnsTrue()
    {
        object input = 42.99;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(43, converted);
    }

    [Fact]
    public void TryConvertToInt32_RoundingDownBoundary_ReturnsTrue()
    {
        object input = 43.01;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(43, converted);
    }

    [Fact]
    public void TryConvertToInt32_TrueBoolean_ReturnsTrue()
    {
        object input = true;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(1, converted);
    }

    [Fact]
    public void TryConvertToInt32_FalseBoolean_ReturnsTrue()
    {
        object input = false;
        var result = ParamTypeConverter.TryConvertToInt32(input, out var converted);
        Assert.True(result);
        Assert.Equal(0, converted);
    }

    [Fact]
    public void ConvertToInt32_ValidInteger_ReturnsInt32()
    {
        object input = 42;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToInt32_ValidStringNumber_ReturnsInt32()
    {
        object input = "42";
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToInt32_ValidDouble_ReturnsInt32()
    {
        object input = 42.5;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToInt32_NonNumericString_ThrowsInvalidCastException()
    {
        object input = "abc";
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToInt32(input));
    }

    [Fact]
    public void ConvertToInt32_NullObject_ThrowsInvalidCastException()
    {
        object input = null!;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToInt32(input));
    }

    [Fact]
    public void ConvertToInt32_VeryLargeDouble_ThrowsInvalidCastException()
    {
        object input = double.MaxValue;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToInt32(input));
    }

    [Fact]
    public void ConvertToInt32_RoundingDown_ReturnsInt32()
    {
        object input = 42.4;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingUp_ReturnsInt32()
    {
        object input = 42.6;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(43, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingUpBoundary_ReturnsInt32()
    {
        object input = 42.99;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(43, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingDownBoundary_ReturnsInt32()
    {
        object input = 43.01;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(43, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingToEven_RoundUp_ReturnsInt32()
    {
        object input = 43.5;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(44, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingToEven_RoundDown_ReturnsInt32()
    {
        object input = 42.5;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToInt32_RoundingAwayFromZero_ReturnsInt32()
    {
        object input = 43.5;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(44, result);
    }

    [Fact]
    public void ConvertToInt32_TrueBoolean_ReturnsInt32()
    {
        object input = true;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(1, result);
    }

    [Fact]
    public void ConvertToInt32_FalseBoolean_ReturnsInt32()
    {
        object input = false;
        var result = ParamTypeConverter.ConvertToInt32(input);
        Assert.Equal(0, result);
    }
}
