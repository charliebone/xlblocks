namespace XlBlocks.AddIn.Dna;

using System.Collections;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using ExcelDna.Integration;
using ExcelDna.Registration;
using NLog;
using XlBlocks.AddIn.Types;

[AttributeUsage(AttributeTargets.ReturnValue, AllowMultiple = false)]
internal class ArrayReturnAttribute : Attribute
{
    public RangeOrientation RangeOrientation { get; set; }
}

internal class ArrayReturnConversionException : Exception
{
    public object ReturnValue { get; }

    public ArrayReturnConversionException(Exception exception, object returnValue) :
        base($"Output of type '{returnValue.GetType().Name}' with value '{returnValue}' could not be converted to an array", exception)
    {
        ReturnValue = returnValue;
    }
}

internal static class ArrayReturnConversion
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(ArrayReturnConversion).FullName);

    private static bool IsArrayable(Type returnType) => typeof(IXlBlockArrayableObject).IsAssignableFrom(returnType);

    private static bool IsEnumerable(Type returnType) => returnType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(returnType);

    public static Func<Type, ExcelReturnRegistration, LambdaExpression?> GetArrayReturnParameterConversion()
    {
        return (type, returnRegistration) => ArrayParameterConversion(type, returnRegistration);
    }

    public static LambdaExpression? ArrayParameterConversion(Type returnType, ExcelReturnRegistration returnRegistration)
    {
        if (returnRegistration.CustomAttributes.OfType<CacheContentsAttribute>().Any())
            return null;

        try
        {
            if (IsArrayable(returnType))
            {
                return ArrayableParameterConversion(returnRegistration);
            }
            else if (IsEnumerable(returnType))
            {
                return EnumerableConversion(returnRegistration);
            }
            return null;
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to register arrayable return parameter conversion", ex);
            return null;
        }
    }

    private static LambdaExpression? ArrayableParameterConversion(ExcelReturnRegistration returnRegistration)
    {
        var arrayableParamExpr = Expression.Parameter(typeof(IXlBlockArrayableObject), "arrayable");
        var returnAttribute = returnRegistration.CustomAttributes.OfType<ArrayReturnAttribute>().FirstOrDefault();
        var orientationExpr = Expression.Constant(returnAttribute?.RangeOrientation ?? RangeOrientation.ByColumn, typeof(RangeOrientation));

        var conversionExpr = Expression.Call(arrayableParamExpr, nameof(IXlBlockArrayableObject.AsArray), null, orientationExpr);
        return Expression.Lambda(conversionExpr, false, new[] { arrayableParamExpr });
    }

    private static LambdaExpression? EnumerableConversion(ExcelReturnRegistration returnRegistration)
    {
        var enumerableParamExpr = Expression.Parameter(typeof(IEnumerable), "enumerable");
        var returnAttribute = returnRegistration.CustomAttributes.OfType<ArrayReturnAttribute>().FirstOrDefault();
        var orientationExpr = Expression.Constant(returnAttribute?.RangeOrientation ?? RangeOrientation.ByColumn, typeof(RangeOrientation));

        var conversionExpr = Expression.Call(typeof(ArrayReturnConversion), nameof(EnumerableTo2DArray), null, enumerableParamExpr, orientationExpr);
        return Expression.Lambda(conversionExpr, false, new[] { enumerableParamExpr });
    }

    private static object[,] EnumerableTo2DArray(IEnumerable enumerable, RangeOrientation rangeOrientation)
    {
        try
        {
            if (enumerable is null)
                return new object[,] { { ExcelError.ExcelErrorNA } };

            if (enumerable is object[,] objectArray)
                return objectArray;

            if (enumerable is IDictionary dictionary)
            {
                var iDict = 0;
                var dictArray = new object[dictionary.Count, 2];
                foreach (DictionaryEntry kvp in dictionary)
                {
                    dictArray[iDict, 0] = kvp.Key;
                    dictArray[iDict, 1] = kvp.Value ?? ExcelError.ExcelErrorNA;
                    iDict++;
                }
                return dictArray;
            }

            var i = 0;
            var enumerableArray = enumerable.Cast<object>().ToArray();
            var outputArray = rangeOrientation == RangeOrientation.ByColumn ?
                new object[enumerableArray.Length, 1] : new object[1, enumerableArray.Length];
            foreach (var item in enumerable)
            {
                if (rangeOrientation == RangeOrientation.ByColumn)
                {
                    outputArray[i, 0] = item;
                }
                else
                {
                    outputArray[0, i] = item;
                }
                i++;
            }
            return outputArray;
        }
        catch (Exception ex)
        {
            throw new ArrayReturnConversionException(ex, enumerable);
        }
    }
}
