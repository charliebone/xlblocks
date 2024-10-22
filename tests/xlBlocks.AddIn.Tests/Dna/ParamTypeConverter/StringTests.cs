namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using System.Globalization;
using XlBlocks.AddIn.Dna;

public class StringTests
{
    [Fact]
    public void TryConvertToString_ValidInteger_ReturnsTrue()
    {
        object input = 42;
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("42", converted);
    }

    [Fact]
    public void TryConvertToString_ValidDouble_ReturnsTrue()
    {
        object input = 42.5;
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("42.5", converted);
    }

    [Fact]
    public void TryConvertToString_ValidBooleanTrue_ReturnsTrue()
    {
        object input = true;
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("true", converted);
    }

    [Fact]
    public void TryConvertToString_ValidBooleanFalse_ReturnsTrue()
    {
        object input = false;
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("false", converted);
    }

    [Fact]
    public void TryConvertToString_ValidDateTime_DateOnly_ReturnsString()
    {
        object input = new DateTime(2023, 1, 23);
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("2023-01-23", converted);
    }

    [Fact]
    public void TryConvertToString_ValidDateTime_DateAndTime_ReturnsString()
    {
        object input = new DateTime(2023, 1, 23, 10, 42, 23);
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("2023-01-23 10:42:23", converted);
    }

    [Fact]
    public void TryConvertToString_ValidDateTime_DateAndTimeWithMs_ReturnsString()
    {
        object input = new DateTime(2023, 1, 23, 10, 42, 23, 456);
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal("2023-01-23 10:42:23.456", converted);
    }

    [Fact]
    public void TryConvertToString_StringWithCultureInvariant_ReturnsTrue()
    {
        double number = 1234.567;
        var cultureInfo = new CultureInfo("fr-FR"); // French culture uses commas as decimal separators
        Thread.CurrentThread.CurrentCulture = cultureInfo;
        object input = number.ToString(cultureInfo);
        var result = ParamTypeConverter.TryConvertToString(input, out var converted);
        Assert.True(result);
        Assert.Equal(input.ToString(), converted);
    }

    [Fact]
    public void ConvertToString_ValidInteger_ReturnsString()
    {
        object input = 42;
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("42", result);
    }

    [Fact]
    public void ConvertToString_ValidDouble_ReturnsString()
    {
        object input = 42.5;
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("42.5", result);
    }

    [Fact]
    public void ConvertToString_AnotherValidDouble_ReturnsString()
    {
        object input = 12346.678;
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("12346.678", result);
    }

    [Fact]
    public void ConvertToString_ValidBooleanTrue_ReturnsString()
    {
        object input = true;
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("true", result);
    }

    [Fact]
    public void ConvertToString_ValidBooleanFalse_ReturnsString()
    {
        object input = false;
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("false", result);
    }

    [Fact]
    public void ConvertToString_ValidDateTime_DateOnly_ReturnsString()
    {
        object input = new DateTime(2023, 1, 23);
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("2023-01-23", result);
    }

    [Fact]
    public void ConvertToString_ValidDateTime_DateAndTime_ReturnsString()
    {
        object input = new DateTime(2023, 1, 23, 10, 42, 23);
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal("2023-01-23 10:42:23", result);
    }

    [Fact]
    public void ConvertToString_StringWithCultureInvariant_ReturnsString()
    {
        double number = 1234.567;
        var cultureInfo = new CultureInfo("fr-FR"); // French culture uses commas as decimal separators
        Thread.CurrentThread.CurrentCulture = cultureInfo;
        object input = number.ToString(cultureInfo);
        var result = ParamTypeConverter.ConvertToString(input);
        Assert.Equal(input.ToString(), result);
    }
}
