namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using System.Globalization;
using XlBlocks.AddIn.Dna;

public class DateTimeTests
{
    [Fact]
    public void TryConvertToDateTime_ValidStringDate_ReturnsTrue()
    {
        object input = "2023-01-23";
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23), converted);
    }

    [Fact]
    public void TryConvertToDateTime_ValidStringDateTime_ReturnsTrue()
    {
        object input = "2023-01-23T12:34:56";
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23, 12, 34, 56), converted);
    }

    [Fact]
    public void TryConvertToDateTime_ValidStringDateTimeWithMs_ReturnsTrue()
    {
        object input = "2023-01-23T12:34:56.789";
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23, 12, 34, 56, 789), converted);
    }

    [Fact]
    public void TryConvertToDateTime_ValidDateTime_ReturnsTrue()
    {
        object input = new DateTime(2023, 1, 23);
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23), converted);
    }

    [Fact]
    public void TryConvertToDateTime_NonDateString_ReturnsFalse()
    {
        object input = "not a date";
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToDateTime_NullObject_ReturnsFalse()
    {
        object input = null!;
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToDateTime_ValidDouble_ReturnsTrue()
    {
        object input = 44197.0; // OLE Automation date for "2021-01-01"
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2021, 1, 1), converted);
    }

    [Fact]
    public void TryConvertToDateTime_ValidStringWithCultureInvariant_ReturnsTrue()
    {
        object input = "01/02/2023";
        var cultureInfo = new CultureInfo("fr-FR"); // French culture: day/month/year
        Thread.CurrentThread.CurrentCulture = cultureInfo;
        var result = ParamTypeConverter.TryConvertToDateTime(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 2, 1), converted);
    }

    [Fact]
    public void ConvertToDateTime_ValidStringDate_ReturnsDateTime()
    {
        object input = "2023-01-23";
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2023, 1, 23), result);
    }

    [Fact]
    public void ConvertToDateTime_ValidStringDateTime_ReturnsDateTime()
    {
        object input = "2023-01-23T12:34:56";
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2023, 1, 23, 12, 34, 56), result);
    }

    [Fact]
    public void ConvertToDateTime_ValidStringDateTimeWithMs_ReturnsDateTime()
    {
        object input = "2023-01-23T12:34:56.789";
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2023, 1, 23, 12, 34, 56, 789), result);
    }

    [Fact]
    public void ConvertToDateTime_ValidDateTime_ReturnsDateTime()
    {
        object input = new DateTime(2023, 1, 23);
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2023, 1, 23), result);
    }

    [Fact]
    public void ConvertToDateTime_NonDateString_ThrowsInvalidCastException()
    {
        object input = "not a date";
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToDateTime(input));
    }

    [Fact]
    public void ConvertToDateTime_NullObject_ThrowsInvalidCastException()
    {
        object input = null!;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToDateTime(input));
    }

    [Fact]
    public void ConvertToDateTime_ValidDouble_ReturnsDateTime()
    {
        object input = 44197.0; // OLE Automation date for "2021-01-01"
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2021, 1, 1), result);
    }

    [Fact]
    public void ConvertToDateTime_ValidStringWithCultureInvariant_ReturnsDateTime()
    {
        object input = "01/02/2023";
        var cultureInfo = new CultureInfo("fr-FR"); // French culture: day/month/year
        Thread.CurrentThread.CurrentCulture = cultureInfo;
        var result = ParamTypeConverter.ConvertToDateTime(input);
        Assert.Equal(new DateTime(2023, 2, 1), result);
    }

}
