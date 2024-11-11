namespace XlBlocks.AddIn.Dna;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using ExcelDna.Integration;

internal static class ParamTypeConverter
{
    /* 
     * input parameters will have their type based on what excel knows it as; strings -> strings, numeric->doubles, logical->bools
     * can also get ExcelError, ExcelEmpty, ExcelMising, ExcelReference, object[,], 
     * excel dna will pass these converted parameters boxed as objects
     * with that said this class doesn't use the excel c api to coerce and instead uses native .net so its usable outside of excel functions
     * some of the overall logic here is still based on the excel-dna type converter
     */

    #region Type conversions via generics and non-generic type parameters

    public static Expression GetConversionExpression(Expression input, Type type)
    {
        return Expression.Call(typeof(ParamTypeConverter), nameof(ConvertTo), new[] { type }, input);
    }

    public static object? GetDefault(Type type)
    {
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }

    public static T ConvertTo<T>(object value)
    {
        if (!TryConvertTo<T>(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(T).Name}");

        return converted;
    }

    public static bool TryConvertTo<T>(object input, [NotNullWhen(true)] out T? convertedInput)
    {
        if (!TryConvertToProvidedType(input, typeof(T), out var convertedObject))
        {
            convertedInput = default;
            return false;
        }

        convertedInput = (T)convertedObject;
        return true;
    }

    public static Type? StringToType(string typeStr)
    {
        return typeStr.ToLowerInvariant().Replace("system.", "") switch
        {
            "string" or "str" => typeof(string),
            "double" or "dbl" => typeof(double),
            "single" or "float" => typeof(float),
            "boolean" or "bool" => typeof(bool),
            "datetime" or "date" => typeof(DateTime),
            "int" or "integer" or "int32" => typeof(int),
            "uint" or "uint32" => typeof(uint),
            "long" or "int64" => typeof(long),
            "ulong" or "uint64" => typeof(ulong),
            "short" or "int16" => typeof(short),
            "ushort" or "uint16" => typeof(ushort),
            "sbyte" or "int8" => typeof(sbyte),
            "byte" or "uint8" => typeof(byte),
            "decimal" => typeof(decimal),
            "object" or "obj" => typeof(object),
            _ => null
        };
    }

    public static bool TryConvertToProvidedType(object input, string type, [NotNullWhen(true)] out object? convertedInput, out Type convertedType)
    {
        var desiredType = StringToType(type);
        if (desiredType is null)
        {
            convertedInput = input;
            convertedType = typeof(object);
            return false;
        }

        // we could resolve a type from the string...
        convertedType = desiredType;
        return TryConvertToProvidedType(input, desiredType, out convertedInput);
    }

    public static bool TryConvertToProvidedType(object input, Type type, [NotNullWhen(true)] out object? convertedInput)
    {
        if (type == typeof(object))
        {
            convertedInput = input;
            return true;
        }

        var nullableUnderlying = Nullable.GetUnderlyingType(type);
        if (nullableUnderlying is not null)
            type = nullableUnderlying;

        if (type == typeof(string) && TryConvertToString(input, out var convertedString))
        {
            convertedInput = convertedString;
        }
        else if (type == typeof(double) && TryConvertToDouble(input, out var convertedDouble))
        {
            convertedInput = convertedDouble;
        }
        else if (type == typeof(float) && TryConvertToDouble(input, out var convertedSingle))
        {
            convertedInput = convertedSingle;
        }
        else if (type == typeof(bool) && TryConvertToBoolean(input, true, out var convertedBool))
        {
            convertedInput = convertedBool;
        }
        else if (type == typeof(DateTime) && TryConvertToDateTime(input, out var convertedDate))
        {
            convertedInput = convertedDate;
        }
        else if (type == typeof(int) && TryConvertToInt32(input, out var convertedInt32))
        {
            convertedInput = convertedInt32;
        }
        else if (type == typeof(uint) && TryConvertToUInt32(input, out var convertedUInt32))
        {
            convertedInput = convertedUInt32;
        }
        else if (type == typeof(long) && TryConvertToInt64(input, out var convertedInt64))
        {
            convertedInput = convertedInt64;
        }
        else if (type == typeof(ulong) && TryConvertToInt64(input, out var convertedUInt64))
        {
            convertedInput = convertedUInt64;
        }
        else if (type == typeof(short) && TryConvertToInt16(input, out var convertedInt16))
        {
            convertedInput = convertedInt16;
        }
        else if (type == typeof(ushort) && TryConvertToUInt16(input, out var convertedUInt16))
        {
            convertedInput = convertedUInt16;
        }
        else if (type == typeof(sbyte) && TryConvertToInt8(input, out var convertedInt8))
        {
            convertedInput = convertedInt8;
        }
        else if (type == typeof(byte) && TryConvertToUInt8(input, out var convertedUInt8))
        {
            convertedInput = convertedUInt8;
        }
        else if (type == typeof(decimal) && TryConvertToDecimal(input, out var convertedDecimal))
        {
            convertedInput = convertedDecimal;
        }
        else if (type == typeof(object[,]) && TryConvertToObjectArray(input, out var convertedObjectArray))
        {
            convertedInput = convertedObjectArray;
        }
        else
        {
            // conversion has failed
            convertedInput = default;
            return false;
        }

        if (nullableUnderlying is not null)
        {
            var nullable = Activator.CreateInstance(typeof(Nullable<>).MakeGenericType(type), convertedInput);
            if (nullable is null)
                return false;
            convertedInput = nullable;
        }

        return true;
    }

    public static bool TryConvertToBestType(object input, [NotNullWhen(true)] out object? convertedInput, out Type convertedType)
    {
        convertedInput = default;
        convertedType = typeof(object);

        if (input is null or ExcelError or ExcelMissing or ExcelEmpty)
            return false;

        if (TryConvertToBoolean(input, false, out var boolValue))
        {
            convertedInput = boolValue;
            convertedType = typeof(bool);
            return true;
        }

        if (TryConvertToDouble(input, out var doubleValue))
        {
            convertedInput = doubleValue;
            convertedType = typeof(double);
            return true;
        }

        if (TryConvertToString(input, out var stringValue))
        {
            if (TryParseDateString(stringValue, out var parsedDate))
            {
                convertedInput = parsedDate;
                convertedType = typeof(DateTime);
                return true;
            }

            convertedInput = stringValue;
            convertedType = typeof(string);
            return true;
        }

        return false;
    }

    #endregion

    #region Type converter extension methods

    internal readonly record struct ParameterConversionResult(object Input, bool IsMissingOrError, bool Success, object? ConvertedInput, Type ConvertedType);

    public static bool IsMissing(object? input) => input is ExcelMissing || input is ExcelEmpty;

    public static bool IsError(object? input) => input is ExcelError;

    public static bool IsMissingOrError(object? input) => IsMissing(input) || IsError(input);

    public static IEnumerable<T> AsEnumerable<T>(this T[,] inputs)
    {
        for (var i = 0; i < inputs.GetLength(0); i++)
            for (var j = 0; j < inputs.GetLength(1); j++)
                yield return inputs[i, j];
    }

    public static T[,] AsColumnArray<T>(this IEnumerable<T> inputs)
    {
        var count = inputs.Count();
        var result = new T[count, 1];
        var i = 0;
        foreach (var input in inputs)
            result[i++, 0] = input;
        return result;
    }

    public static IEnumerable<T> TryConvertTo<T>(this IEnumerable<object> range, bool dropErrors)
    {
        foreach (var item in range.ConvertToProvidedType(typeof(T)))
        {
            if (!item.Success || item.ConvertedInput is not T typedInput)
            {
                if (!dropErrors)
                    throw new ArgumentException($"cannot convert value '{item.Input}' into type '{typeof(T).Name}'");
                else
                    continue;
            }

            yield return typedInput;
        }
    }

    public static IEnumerable<object> TryConvertToType(this IEnumerable<object> range, Type type, bool dropErrors)
    {
        foreach (var item in range.ConvertToProvidedType(type))
        {
            if (!item.Success || item.ConvertedInput is null || item.ConvertedInput.GetType() != type)
            {
                if (!dropErrors)
                    throw new ArgumentException($"cannot convert value '{item.Input}' into type '{type.Name}'");
                else
                    continue;
            }

            yield return item.ConvertedInput;
        }

    }

    public static IEnumerable<T> FilterMissing<T>(this IEnumerable<T> inputs, bool throwOnMissing = false)
    {
        return inputs.Where(x => !IsMissing(x) || (throwOnMissing ? throw new ArgumentException("input has missing values") : false));
    }

    public static IEnumerable<T> FilterErrors<T>(this IEnumerable<T> inputs, bool throwOnError = false)
    {
        return inputs.Where(x => !IsError(x) || (throwOnError ? throw new ArgumentException("input has error values") : false));
    }

    public static Type DetermineBestType(this IEnumerable<Type> types)
    {
        var hasDateTime = false;
        var hasDouble = false;
        var hasBool = false;
        foreach (var type in types)
        {
            // string is the fallback
            if (type == typeof(string))
                return typeof(string);

            if (type == typeof(DateTime))
                hasDateTime = true;

            if (type == typeof(double))
                hasDouble = true;

            if (type == typeof(bool))
                hasBool = true;
        }

        if (hasDateTime && !hasDouble && !hasBool)
            return typeof(DateTime);

        if (hasBool && !hasDateTime && !hasDouble)
            return typeof(bool);

        if (hasDouble && !hasDateTime && !hasBool)
            return typeof(double);

        return typeof(string);
    }

    public static IEnumerable<ParameterConversionResult> ConvertToBestTypes(this IEnumerable<object> inputs)
    {
        return inputs.Select(x =>
        {
            if (IsMissingOrError(x))
                return new ParameterConversionResult(x, true, false, null, typeof(object));

            var success = TryConvertToBestType(x, out var converted, out var convertedType);
            return new ParameterConversionResult(x, false, success, converted, convertedType);
        });
    }

    public static IEnumerable<ParameterConversionResult> ConvertToProvidedType(this object[,] inputs, Type type)
    {
        for (var i = 0; i < inputs.GetLength(0); i++)
        {
            for (var j = 0; j < inputs.GetLength(1); j++)
            {
                if (IsMissingOrError(inputs[i, j]))
                    yield return new ParameterConversionResult(inputs[i, j], true, false, null, type);

                var success = TryConvertToProvidedType(inputs[i, j], type, out var converted);
                yield return new ParameterConversionResult(inputs[i, j], false, success, converted, type);
            }
        }
    }

    public static IEnumerable<ParameterConversionResult> ConvertToProvidedType(this object[,] inputs, string type)
    {
        var convertedType = StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
        return ConvertToProvidedType(inputs, convertedType);
    }

    public static IEnumerable<ParameterConversionResult> ConvertToProvidedType(this IEnumerable<object> inputs, string type)
    {
        var convertedType = StringToType(type) ?? throw new ArgumentException($"unknown type '{type}'");
        return ConvertToProvidedType(inputs, convertedType);
    }

    public static IEnumerable<ParameterConversionResult> ConvertToProvidedType(this IEnumerable<object> inputs, Type type)
    {
        return inputs.Select(x =>
        {
            if (IsMissingOrError(x))
                return new ParameterConversionResult(x, true, false, null, type);

            var success = TryConvertToProvidedType(x, type, out var converted);
            return new ParameterConversionResult(x, false, success, converted, type);
        });
    }

    #endregion

    #region Type converters

    private static bool TryUnboxedDoubleConversion(object value, [NotNullWhen(true)] out double converted)
    {
        try
        {
            switch (value)
            {
                case double doubleValue:
                    converted = doubleValue;
                    break;
                case float floatValue:
                    converted = floatValue;
                    break;
                case long longValue:
                    converted = longValue;
                    break;
                case int intValue:
                    converted = intValue;
                    break;
                case uint uintValue:
                    converted = uintValue;
                    break;
                case short shortValue:
                    converted = shortValue;
                    break;
                case ushort ushortValue:
                    converted = ushortValue;
                    break;
                case sbyte sbyteValue:
                    converted = sbyteValue;
                    break;
                case byte byteValue:
                    converted = byteValue;
                    break;
                default:
                    converted = default;
                    return false;
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static double ConvertToDouble(object value)
    {
        if (!TryConvertToDouble(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(double).Name}");

        return (double)converted;
    }

    public static bool TryConvertToDouble(object value, [NotNullWhen(true)] out double converted)
    {
        if (value is null)
        {
            converted = default;
            return false;
        }

        if (value is double doubleValue)
        {
            converted = doubleValue;
            return true;
        }

        if (value is object[,] array)
            value = array[0, 0];

        if (value is string stringValue)
        {
            // perhaps think about something better here, worth testing more
            if ((double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parseResult) ||
                double.TryParse(stringValue, NumberStyles.Any, CultureInfo.CurrentCulture, out parseResult)) &&
                double.IsFinite(parseResult) && !double.IsNaN(parseResult))
            {
                converted = parseResult;
                return true;
            }

            converted = default;
            return false;
        }

        if (value is bool boolValue)
        {
            converted = boolValue ? 1d : 0d;
            return true;
        }

        if (value is not null && value.GetType().IsValueType && TryUnboxedDoubleConversion(value, out var castDouble))
        {
            converted = castDouble;
            return true;
        }

        converted = default;
        return false;
    }

    public static string ConvertToString(object value)
    {
        if (!TryConvertToString(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(string).Name}");
        return converted;
    }

    private static bool TryFormatDouble(double value, [NotNullWhen(true)] out string? formatted)
    {
        try
        {
            // think about this a bit more, but for now format according to invariant culture which should be like 123.45#### up to 6 decimal places
            formatted = Math.Round(value, 6).ToString("f7", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');
            return true;
        }
        catch
        {
            formatted = default;
            return false;
        }
    }

    public static bool TryConvertToString(object value, [NotNullWhen(true)] out string? converted)
    {
        if (value is string stringValue)
        {
            converted = stringValue;
            return true;
        }

        if (value is object[,] array)
            value = array[0, 0];


        if (value is double doubleValue)
        {
            return TryFormatDouble(doubleValue, out converted);
        }

        if (value is bool boolValue)
        {
            converted = boolValue ? "true" : "false";
            return true;
        }

        if (value is DateTime dateTime)
        {
            // consider what's best option in this case
            var formatStr = "yyyy-MM-dd";
            if (dateTime.Hour != 0 || dateTime.Minute != 0 || dateTime.Second != 0)
            {
                formatStr += " HH:mm:ss";
                if (dateTime.Millisecond != 0)
                    formatStr += ".fff";
            }

            converted = dateTime.ToString(formatStr);
            return true;
        }

        if (value is not null && value.GetType().IsValueType)
        {
            // all other types go through double first
            var success = TryConvertToDouble(value, out var interimDouble);
            if (!success)
            {
                converted = default;
                return false;
            }

            return TryFormatDouble(interimDouble, out converted);
        }

        converted = default;
        return false;
    }

    public static DateTime ConvertToDateTime(object value)
    {
        if (!TryConvertToDateTime(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(DateTime).Name}");

        return converted.Value;
    }

    private static bool TryFromOADate(double doubleValue, [NotNullWhen(true)] out DateTime? converted)
    {
        try
        {
            converted = DateTime.FromOADate(doubleValue);
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    private static bool TryParseDateString(string stringValue, [NotNullWhen(true)] out DateTime? converted)
    {
        stringValue = stringValue.Trim();

        // should put more thought into this
        var iso6801Formats = new[] { "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd" };
        if (DateTime.TryParseExact(stringValue, iso6801Formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate) ||
            DateTime.TryParse(stringValue, CultureInfo.CurrentCulture, DateTimeStyles.None, out parsedDate))
        {
            converted = parsedDate;
            return true;
        }
        else
        {
            converted = default;
            return false;
        }
    }

    public static bool TryConvertToDateTime(object value, [NotNullWhen(true)] out DateTime? converted)
    {
        if (value is DateTime dateValue)
        {
            converted = dateValue;
            return true;
        }

        if (value is object[,] array)
            value = array[0, 0];

        if (value is double doubleValue)
        {
            return TryFromOADate(doubleValue, out converted);
        }

        if (value is string stringValue)
        {
            return TryParseDateString(stringValue, out converted);
        }

        if (value is not null && value is not bool && value.GetType().IsValueType)
        {
            // all other types (except bool which fails) go through double first
            var success = TryConvertToDouble(value, out var interimDouble);
            if (!success)
            {
                converted = default;
                return false;
            }

            return TryFromOADate(interimDouble, out converted);
        }

        converted = default;
        return false;
    }

    public static bool ConvertToBoolean(object value, bool tryNumeric)
    {
        if (!TryConvertToBoolean(value, tryNumeric, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(bool).Name}");
        return converted.Value;
    }

    public static bool TryConvertToBoolean(object value, bool tryNumeric, [NotNullWhen(true)] out bool? converted)
    {
        if (value is bool boolValue)
        {
            converted = boolValue;
            return true;
        }

        if (value is object[,] array)
            value = array[0, 0];

        if (value is string stringValue)
        {
            var success = bool.TryParse(stringValue.ToLowerInvariant(), out var parseResult);
            converted = parseResult;
            return success;
        }

        if (!tryNumeric)
        {
            converted = default;
            return false;
        }

        if (value is double doubleValue)
        {
            converted = doubleValue != 0d;
            return true;
        }

        if (value is not null)
        {
            var valueType = value.GetType();
            if (valueType.IsValueType)
            {
                try
                {
                    converted = !value.Equals(Activator.CreateInstance(valueType));
                    return true;
                }
                catch
                {
                    converted = default;
                    return false;
                }
            }
        }

        converted = default;
        return false;
    }

    public static int ConvertToInt32(object value)
    {
        if (!TryConvertToInt32(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(int).Name}");
        return converted.Value;
    }

    public static bool TryConvertToInt32(object value, [NotNullWhen(true)] out int? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((int)longValue);
            else if (value is ulong ulongValue)
                converted = checked((int)ulongValue);
            else if (value is int intValue)
                converted = intValue;
            else if (value is uint uintValue)
                converted = checked((int)uintValue);
            else if (value is short shortValue)
                converted = shortValue;
            else if (value is ushort ushortValue)
                converted = ushortValue;
            else if (value is sbyte sbyteValue)
                converted = sbyteValue;
            else if (value is byte byteValue)
                converted = byteValue;
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((int)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static uint ConvertToUInt32(object value)
    {
        if (!TryConvertToUInt32(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to UInt");
        return converted.Value;
    }

    public static bool TryConvertToUInt32(object value, [NotNullWhen(true)] out uint? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((uint)longValue);
            else if (value is ulong ulongValue)
                converted = checked((uint)ulongValue);
            else if (value is int intValue)
                converted = checked((uint)intValue);
            else if (value is uint uintValue)
                converted = uintValue;
            else if (value is short shortValue)
                converted = checked((uint)shortValue);
            else if (value is ushort ushortValue)
                converted = ushortValue;
            else if (value is sbyte sbyteValue)
                converted = checked((uint)sbyteValue);
            else if (value is byte byteValue)
                converted = byteValue;
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((uint)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static short ConvertToInt16(object value)
    {
        if (!TryConvertToInt16(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(short).Name}");
        return converted.Value;
    }

    public static bool TryConvertToInt16(object value, [NotNullWhen(true)] out short? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((short)longValue);
            else if (value is ulong ulongValue)
                converted = checked((short)ulongValue);
            else if (value is int intValue)
                converted = checked((short)intValue);
            else if (value is uint uintValue)
                converted = checked((short)uintValue);
            else if (value is short shortValue)
                converted = shortValue;
            else if (value is ushort ushortValue)
                converted = checked((short)ushortValue);
            else if (value is sbyte sbyteValue)
                converted = checked(sbyteValue);
            else if (value is byte byteValue)
                converted = byteValue;
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((short)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static ushort ConvertToUInt16(object value)
    {
        if (!TryConvertToUInt16(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(ushort).Name}");
        return converted.Value;
    }

    public static bool TryConvertToUInt16(object value, [NotNullWhen(true)] out ushort? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((ushort)longValue);
            else if (value is ulong ulongValue)
                converted = checked((ushort)ulongValue);
            else if (value is int intValue)
                converted = checked((ushort)intValue);
            else if (value is uint uintValue)
                converted = checked((ushort)uintValue);
            else if (value is short shortValue)
                converted = checked((ushort)shortValue);
            else if (value is ushort ushortValue)
                converted = ushortValue;
            else if (value is sbyte sbyteValue)
                converted = checked((ushort)sbyteValue);
            else if (value is byte byteValue)
                converted = byteValue;
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((ushort)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static sbyte ConvertToInt8(object value)
    {
        if (!TryConvertToInt8(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(sbyte).Name}");
        return converted.Value;
    }

    public static bool TryConvertToInt8(object value, [NotNullWhen(true)] out sbyte? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((sbyte)longValue);
            else if (value is ulong ulongValue)
                converted = checked((sbyte)ulongValue);
            else if (value is int intValue)
                converted = checked((sbyte)intValue);
            else if (value is uint uintValue)
                converted = checked((sbyte)uintValue);
            else if (value is short shortValue)
                converted = checked((sbyte)shortValue);
            else if (value is ushort ushortValue)
                converted = checked((sbyte)ushortValue);
            else if (value is sbyte sbyteValue)
                converted = sbyteValue;
            else if (value is byte byteValue)
                converted = checked((sbyte)byteValue);
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((sbyte)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static byte ConvertToUInt8(object value)
    {
        if (!TryConvertToUInt8(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(byte).Name}");
        return converted.Value;
    }

    public static bool TryConvertToUInt8(object value, [NotNullWhen(true)] out byte? converted)
    {
        try
        {
            if (value is long longValue)
                converted = checked((byte)longValue);
            else if (value is ulong ulongValue)
                converted = checked((byte)ulongValue);
            else if (value is int intValue)
                converted = checked((byte)intValue);
            else if (value is uint uintValue)
                converted = checked((byte)uintValue);
            else if (value is short shortValue)
                converted = checked((byte)shortValue);
            else if (value is ushort ushortValue)
                converted = checked((byte)ushortValue);
            else if (value is sbyte sbyteValue)
                converted = checked((byte)sbyteValue);
            else if (value is byte byteValue)
                converted = byteValue;
            else
            {
                converted = default;
                if (!TryConvertToInt64(value, out var convertedLong))
                    return false;

                converted = checked((byte)convertedLong);
            }
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static decimal ConvertToDecimal(object value)
    {
        if (!TryConvertToDecimal(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(decimal).Name}");
        return converted.Value;
    }

    public static bool TryConvertToDecimal(object value, [NotNullWhen(true)] out decimal? converted)
    {
        if (value is decimal decimalValue)
        {
            converted = decimalValue;
            return true;
        }

        converted = default;
        if (!TryConvertToDouble(value, out var convertedDouble))
            return false;

        try
        {
            converted = checked((decimal)convertedDouble);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static long ConvertToInt64(object value)
    {
        if (!TryConvertToInt64(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(long).Name}");
        return converted.Value;
    }

    public static bool TryConvertToInt64(object value, [NotNullWhen(true)] out long? converted)
    {
        try
        {
            converted = default;
            if (value is double doubleValue) // all double -> integer types generally go through here first
                converted = checked((long)Math.Round(doubleValue, MidpointRounding.ToEven));
            else if (value is float floatValue)
                converted = checked((long)Math.Round(floatValue, MidpointRounding.ToEven));
            else if (value is long longValue)
                converted = checked(longValue);
            else if (value is int intValue)
                converted = checked(intValue);
            else if (value is uint uintValue)
                converted = checked(uintValue);
            else if (value is short shortValue)
                converted = checked(shortValue);
            else if (value is ushort ushortValue)
                converted = checked(ushortValue);
            else if (value is sbyte sbyteValue)
                converted = checked(sbyteValue);
            else if (value is byte byteValue)
                converted = checked(byteValue);
            else
            {
                converted = default;
                if (!TryConvertToDouble(value, out var convertedDouble))
                    return false;

                converted = checked((long)Math.Round(convertedDouble, MidpointRounding.ToEven));
            }

            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static ulong ConvertToUInt64(object value)
    {
        if (!TryConvertToUInt64(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(ulong).Name}");
        return converted.Value;
    }

    public static bool TryConvertToUInt64(object value, [NotNullWhen(true)] out ulong? converted)
    {
        try
        {
            converted = default;
            if (value is double doubleValue)
                converted = checked((ulong)Math.Round(doubleValue, MidpointRounding.ToEven));
            else if (value is float floatValue)
                converted = checked((ulong)Math.Round(floatValue, MidpointRounding.ToEven));
            else if (value is long longValue)
                converted = checked((ulong)longValue);
            else if (value is int intValue)
                converted = checked((ulong)intValue);
            else if (value is uint uintValue)
                converted = checked(uintValue);
            else if (value is short shortValue)
                converted = checked((ulong)shortValue);
            else if (value is ushort ushortValue)
                converted = checked(ushortValue);
            else if (value is sbyte sbyteValue)
                converted = checked((ulong)sbyteValue);
            else if (value is byte byteValue)
                converted = checked(byteValue);
            else
            {
                converted = default;
                if (!TryConvertToDouble(value, out var convertedDouble))
                    return false;

                converted = checked((ulong)Math.Round(convertedDouble, MidpointRounding.ToEven));
            }

            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static float ConvertToSingle(object value)
    {
        if (!TryConvertToSingle(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to {typeof(float).Name}");
        return converted.Value;
    }

    public static bool TryConvertToSingle(object value, [NotNullWhen(true)] out float? converted)
    {
        if (!TryConvertToDouble(value, out var convertedDouble))
        {
            converted = default;
            return false;
        }

        try
        {
            converted = checked((float)convertedDouble);
            return true;
        }
        catch
        {
            converted = default;
            return false;
        }
    }

    public static object[,] ConvertToObjectArray(object value)
    {
        if (!TryConvertToObjectArray(value, out var converted))
            throw new InvalidCastException($"Value {value} could not be converted to an Object array");
        return converted;
    }

    public static bool TryConvertToObjectArray(object value, [NotNullWhen(true)] out object[,] converted)
    {
        if (value is object[,] objectArray)
        {
            converted = objectArray;
        }
        else
        {
            converted = (new[,] { { value } });
        }

        return true;
    }

    #endregion
}
