namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using XlBlocks.AddIn.Dna;

public class BooleanTests
{
    [Fact]
    public void TryConvertToBoolean_ValidStringTrue_ReturnsTrue()
    {
        object input = "true";
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidStringFalse_ReturnsTrue()
    {
        object input = "false";
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.False(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidStringTrue_Capitalization_ReturnsTrue()
    {
        object input = "tRUe";
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidStringFalse_Capitalization_ReturnsTrue()
    {
        object input = "False";
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.False(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidTrue_ReturnsTrue()
    {
        object input = true;
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidFalse_ReturnsTrue()
    {
        object input = false;
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.True(result);
        Assert.False(converted);
    }

    [Fact]
    public void TryConvertToBoolean_InvalidInteger_ReturnsFalse()
    {
        object input = 1;
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidIntegerOne_ReturnsTrue()
    {
        object input = 1;
        var result = ParamTypeConverter.TryConvertToBoolean(input, true, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidIntegerZero_ReturnsTrue()
    {
        object input = 0;
        var result = ParamTypeConverter.TryConvertToBoolean(input, true, out var converted);
        Assert.True(result);
        Assert.False(converted);
    }

    [Fact]
    public void TryConvertToBoolean_NonBooleanString_ReturnsFalse()
    {
        object input = "yes";
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.False(result);
        Assert.Equal(false, converted);
    }

    [Fact]
    public void TryConvertToBoolean_NullObject_ReturnsFalse()
    {
        object input = null!;
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToBoolean_InvalidDouble_ReturnsFalse()
    {
        object input = 1.0;
        var result = ParamTypeConverter.TryConvertToBoolean(input, false, out var converted);
        Assert.False(result);
        Assert.Equal(default, converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidDoubleOne_ReturnsTrue()
    {
        object input = 1.0;
        var result = ParamTypeConverter.TryConvertToBoolean(input, true, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToBoolean_ValidDoubleZero_ReturnsTrue()
    {
        object input = 0.0;
        var result = ParamTypeConverter.TryConvertToBoolean(input, true, out var converted);
        Assert.True(result);
        Assert.False(converted);
    }

    [Fact]
    public void ConvertToBoolean_ValidStringTrue_ReturnsBoolean()
    {
        object input = "true";
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.True(result);
    }

    [Fact]
    public void ConvertToBoolean_ValidStringFalse_ReturnsBoolean()
    {
        object input = "false";
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.False(result);
    }

    [Fact]
    public void ConvertToBoolean_ValidStringTrue_Capitalization_ReturnsBoolean()
    {
        object input = "tRUe";
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.True(result);
    }

    [Fact]
    public void ConvertToBoolean_ValidStringFalse_Capitalization_ReturnsBoolean()
    {
        object input = "False";
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.False(result);
    }

    [Fact]
    public void ConvertToBoolean_ValidTrue_ReturnsBoolean()
    {
        object input = true;
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.True(result);
    }

    [Fact]
    public void ConvertToBoolean_ValidFalse_ReturnsBoolean()
    {
        object input = false;
        var result = ParamTypeConverter.ConvertToBoolean(input, false);
        Assert.False(result);
    }

    [Fact]
    public void ConvertToBoolean_NonBooleanString_ThrowsInvalidCastException()
    {
        object input = "yes";
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToBoolean(input, false));
    }

    [Fact]
    public void ConvertToBoolean_NullObject_ThrowsInvalidCastException()
    {
        object input = null!;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToBoolean(input, false));
    }

    [Fact]
    public void ConvertToBoolean_Numeric_ThrowsInvalidCastException()
    {
        object input = 1.0;
        Assert.Throws<InvalidCastException>(() => ParamTypeConverter.ConvertToBoolean(input, false));
    }
}
