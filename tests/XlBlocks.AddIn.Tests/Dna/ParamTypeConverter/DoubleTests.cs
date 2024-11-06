namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using System.Globalization;
using XlBlocks.AddIn.Dna;

public class DoubleTests
{
    [Fact]
    public void TryConvertToDouble_ValidInteger_ReturnsTrue()
    {
        object input = 42;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(42.0, converted);
    }

    [Fact]
    public void TryConvertToDouble_ValidStringNumber_ReturnsTrue()
    {
        object input = "42.5";
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(42.5, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithScientificNotation_ReturnsTrue()
    {
        object input = "4.25e1";
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(42.5, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithOverflow_ReturnsFalse()
    {
        object input = "1.7976931348623159E+308"; // Just above double.MaxValue
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithThousandSeparator_ReturnsTrue()
    {
        var number = 1234567.89;
        object input = number.ToString("N", CultureInfo.CurrentCulture); // Format with thousands separator
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithDecimalPoint_ReturnsTrue()
    {
        var number = 1234.567;
        object input = number.ToString(CultureInfo.CurrentCulture); // Format with decimal point
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToDouble_ValidDouble_ReturnsTrue()
    {
        object input = 42.5;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(42.5, converted);
    }

    [Fact]
    public void TryConvertToDouble_NonNumericString_ReturnsFalse()
    {
        object input = "abc";
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToDouble_NullObject_ReturnsFalse()
    {
        object input = null!;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithComma_ReturnsTrue()
    {
        var number = 1234.567;
        object input = number.ToString("N3", CultureInfo.CurrentCulture);
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithCommaDifferentCulture_ReturnsTrue()
    {
        var number = 1234.567;
        var frenchCulture = new CultureInfo("fr-FR");
        object input = number.ToString("N3", frenchCulture);
        Thread.CurrentThread.CurrentCulture = frenchCulture;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToDouble_StringWithMultipleCommas_ReturnsTrue()
    {
        var number = 1234567.89;
        object input = number.ToString("N", CultureInfo.CurrentCulture); // Format with multiple commas
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(number, converted);
    }

    [Fact]
    public void TryConvertToDouble_TrueBoolean_ReturnsTrue()
    {
        object input = true;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(1.0, converted);
    }

    [Fact]
    public void TryConvertToDouble_FalseBoolean_ReturnsTrue()
    {
        object input = false;
        var result = ParamTypeConverter.TryConvertToDouble(input, out var converted);
        Assert.True(result);
        Assert.Equal(0.0, converted);
    }

    [Fact]
    public void ConvertToDouble_ValidInteger_ReturnsDouble()
    {
        object input = 42;
        var result = ParamTypeConverter.ConvertToDouble(input);
        Assert.Equal(42.0, result);
    }

    [Fact]
    public void ConvertToDouble_ValidStringNumber_ReturnsDouble()
    {
        object input = "42.5";
        var result = ParamTypeConverter.ConvertToDouble(input);
        Assert.Equal(42.5, result);
    }

    [Fact]
    public void ConvertToDouble_ValidDouble_ReturnsDouble()
    {
        object input = 42.5;
        var result = ParamTypeConverter.ConvertToDouble(input);
        Assert.Equal(42.5, result);
    }

    [Fact]
    public void ConvertToDouble_NonNumericString_ThrowsInvalidCastException()
    {
        object input = "abc";
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToDouble(input));
    }

    [Fact]
    public void ConvertToDouble_NullObject_ThrowsInvalidCastException()
    {
        object input = null!;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToDouble(input));
    }

    [Fact]
    public void ConvertToDouble_TrueBoolean_ReturnsDouble()
    {
        object input = true;
        var result = ParamTypeConverter.ConvertToDouble(input);
        Assert.Equal(1.0, result);
    }

    [Fact]
    public void ConvertToDouble_FalseBoolean_ReturnsDouble()
    {
        object input = false;
        var result = ParamTypeConverter.ConvertToDouble(input);
        Assert.Equal(0.0, result);
    }
}
