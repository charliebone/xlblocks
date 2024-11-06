namespace XlBlocks.AddIn.Tests.Dna.ParamTypeConverter;

using XlBlocks.AddIn.Dna;

public class ParamTypeConverterTests
{
    [Fact]
    public void ConvertToGeneric_Double_ReturnsDouble()
    {
        var input = "42.5";
        var result = ParamTypeConverter.ConvertTo<double>(input);
        Assert.Equal(42.5, result);
    }

    [Fact]
    public void ConvertToGeneric_String_ReturnsString()
    {
        var input = 42.5;
        var result = ParamTypeConverter.ConvertTo<string>(input);
        Assert.Equal("42.5", result);
    }

    [Fact]
    public void ConvertToGeneric_Int_ReturnsInt()
    {
        var input = "42";
        var result = ParamTypeConverter.ConvertTo<int>(input);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertToGeneric_Bool_ReturnsBool()
    {
        var input = "true";
        var result = ParamTypeConverter.ConvertTo<bool>(input);
        Assert.True(result);
    }

    [Fact]
    public void ConvertToGeneric_DateTime_ReturnsDateTime()
    {
        var input = "2023-01-23";
        var result = ParamTypeConverter.ConvertTo<DateTime>(input);
        Assert.Equal(new DateTime(2023, 1, 23), result);
    }

    [Fact]
    public void TryConvertToGeneric_Double_ReturnsTrue()
    {
        var input = "42.5";
        var result = ParamTypeConverter.TryConvertTo<double>(input, out var converted);
        Assert.True(result);
        Assert.Equal(42.5, converted);
    }

    [Fact]
    public void TryConvertToGeneric_String_ReturnsTrue()
    {
        var input = 42.5;
        var result = ParamTypeConverter.TryConvertTo<string>(input, out var converted);
        Assert.True(result);
        Assert.Equal("42.5", converted);
    }

    [Fact]
    public void TryConvertToGeneric_Int_ReturnsTrue()
    {
        var input = "42";
        var result = ParamTypeConverter.TryConvertTo<int>(input, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
    }

    [Fact]
    public void TryConvertToGeneric_Bool_ReturnsTrue()
    {
        var input = "true";
        var result = ParamTypeConverter.TryConvertTo<bool>(input, out var converted);
        Assert.True(result);
        Assert.True(converted);
    }

    [Fact]
    public void TryConvertToGeneric_DateTime_ReturnsTrue()
    {
        var input = "2023-01-23";
        var result = ParamTypeConverter.TryConvertTo<DateTime>(input, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23), converted);
    }

    [Fact]
    public void TryConvertToProvidedType_DoubleType_ReturnsTrue()
    {
        var input = "42.5";
        var type = typeof(double);
        var result = ParamTypeConverter.TryConvertToProvidedType(input, type, out var converted);
        Assert.True(result);
        Assert.Equal(42.5, converted);
        Assert.Equal(type, converted?.GetType());
    }

    [Fact]
    public void TryConvertToProvidedType_StringType_ReturnsTrue()
    {
        var input = 42.5;
        var type = typeof(string);
        var result = ParamTypeConverter.TryConvertToProvidedType(input, type, out var converted);
        Assert.True(result);
        Assert.Equal("42.5", converted);
        Assert.Equal(type, converted?.GetType());
    }

    [Fact]
    public void TryConvertToProvidedType_IntType_ReturnsTrue()
    {
        var input = "42";
        var type = typeof(int);
        var result = ParamTypeConverter.TryConvertToProvidedType(input, type, out var converted);
        Assert.True(result);
        Assert.Equal(42, converted);
        Assert.Equal(type, converted?.GetType());
    }

    [Fact]
    public void TryConvertToProvidedType_BoolType_ReturnsTrue()
    {
        var input = "true";
        var type = typeof(bool);
        var result = ParamTypeConverter.TryConvertToProvidedType(input, type, out var converted);
        Assert.True(result);
        Assert.Equal(true, converted);
        Assert.Equal(type, converted?.GetType());
    }

    [Fact]
    public void TryConvertToProvidedType_DateTimeType_ReturnsTrue()
    {
        var input = "2023-01-23";
        var type = typeof(DateTime);
        var result = ParamTypeConverter.TryConvertToProvidedType(input, type, out var converted);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23), converted);
        Assert.Equal(typeof(DateTime), converted?.GetType());
    }

    [Fact]
    public void TryConvertToProvidedType_DoubleString_ReturnsTrue()
    {
        var input = "42.5";
        var result = ParamTypeConverter.TryConvertToProvidedType(input, "double", out var converted, out var convertedType);
        Assert.True(result);
        Assert.Equal(42.5, converted);
        Assert.Equal(typeof(double), convertedType);
    }

    [Fact]
    public void TryConvertToProvidedType_StringString_ReturnsTrue()
    {
        var input = 42.5;
        var result = ParamTypeConverter.TryConvertToProvidedType(input, "string", out var converted, out var convertedType);
        Assert.True(result);
        Assert.Equal("42.5", converted);
        Assert.Equal(typeof(string), convertedType);
    }

    [Fact]
    public void TryConvertToProvidedType_IntString_ReturnsTrue()
    {
        var input = "42";
        var result = ParamTypeConverter.TryConvertToProvidedType(input, "int", out var converted, out var convertedType);
        Assert.True(result);
        Assert.Equal(42, converted);
        Assert.Equal(typeof(int), convertedType);
    }

    [Fact]
    public void TryConvertToProvidedType_BoolString_ReturnsTrue()
    {
        var input = "true";
        var result = ParamTypeConverter.TryConvertToProvidedType(input, "bool", out var converted, out var convertedType);
        Assert.True(result);
        Assert.Equal(true, converted);
        Assert.Equal(typeof(bool), convertedType);
    }

    [Fact]
    public void TryConvertToProvidedType_DateTimeString_ReturnsTrue()
    {
        var input = "2023-01-23";
        var result = ParamTypeConverter.TryConvertToProvidedType(input, "datetime", out var converted, out var convertedType);
        Assert.True(result);
        Assert.Equal(new DateTime(2023, 1, 23), converted);
        Assert.Equal(typeof(DateTime), convertedType);
    }

    [Fact]
    public void StringToType_Double_ReturnsDoubleType()
    {
        var result = ParamTypeConverter.StringToType("double");
        Assert.Equal(typeof(double), result);
    }

    [Fact]
    public void StringToType_String_ReturnsStringType()
    {
        var result = ParamTypeConverter.StringToType("string");
        Assert.Equal(typeof(string), result);
    }

    [Fact]
    public void StringToType_Int_ReturnsIntType()
    {
        var result = ParamTypeConverter.StringToType("int");
        Assert.Equal(typeof(int), result);
    }

    [Fact]
    public void StringToType_Bool_ReturnsBoolType()
    {
        var result = ParamTypeConverter.StringToType("bool");
        Assert.Equal(typeof(bool), result);
    }

    [Fact]
    public void StringToType_DateTime_ReturnsDateTimeType()
    {
        var result = ParamTypeConverter.StringToType("datetime");
        Assert.Equal(typeof(DateTime), result);
    }

    [Fact]
    public void StringToType_InvalidString_ReturnsNull()
    {
        var result = ParamTypeConverter.StringToType("invalid");
        Assert.Null(result);
    }

}
