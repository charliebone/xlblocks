namespace XlBlocks.AddIn.Dna;

using System;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Registration;
using NLog;
using XlBlocks.AddIn.Types;

internal class RequiredParameterException : Exception
{
    public string ParameterName { get; set; }
    public override string Message => $"Parameter '{ParameterName}' cannot be missing or empty";

    public RequiredParameterException(string parameterName)
    {
        ParameterName = parameterName;
    }
}

internal static class BaseParameterConversion
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(BaseParameterConversion).FullName);

    public static Func<Type, ExcelParameterRegistration, LambdaExpression?> GetBaseInputParameterConversion()
    {
        return (type, inputParamRegistration) => BaseInputParameterConversion(type, inputParamRegistration);
    }

    private static LambdaExpression? BaseInputParameterConversion(Type inputType, ExcelParameterRegistration inputParamRegistration)
    {
        // ignore params style args here
        if (inputParamRegistration.CustomAttributes.OfType<ParamArrayAttribute>().Any())
            return null;

        var defaultAttribute = inputParamRegistration.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
        if (defaultAttribute is not null || inputParamRegistration.CustomAttributes.OfType<OptionalAttribute>().Any())
        {
            // this parameter is to be considered optional so Missing or Empty values are okay and will be converted into the default
            var defaultValue = defaultAttribute is null || defaultAttribute.Value is System.Reflection.Missing ?
                ParamTypeConverter.GetDefault(inputType) :
                defaultAttribute.Value;
            return OptionalInputParameterConversion(inputType, inputParamRegistration, defaultValue);
        }
        else
        {
            // this parameter is to be considered required and we will throw on Missing or Empty values (error values are okay)
            return RequiredInputParameterConversion(inputType, inputParamRegistration);
        }
    }

    private static LambdaExpression? RequiredInputParameterConversion(Type inputType, ExcelParameterRegistration inputParamRegistration)
    {
        try
        {
            if (Nullable.GetUnderlyingType(inputType) != null)
                throw new Exception("Nullable parameters must be optional");

            // if ExcelMissing or ExcelEmpty, throw exception (will eventually return NA)
            var paramExpr = Expression.Parameter(typeof(object), "input");
            var paramNameExpr = Expression.Constant(inputParamRegistration.ArgumentAttribute.Name);
            var expr = Expression.Condition(
                Expression.Call(typeof(BaseParameterConversion), nameof(CheckIfMissing), null, paramExpr),
                Expression.Call(typeof(BaseParameterConversion), nameof(ThrowRequiredParameterException), new[] { inputType }, paramExpr, paramNameExpr),
                GetConversionExpression(paramExpr, inputType));

            return Expression.Lambda(expr, false, new[] { paramExpr });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Failed to register required input parameter conversion for parameter {inputParamRegistration.ArgumentAttribute.Name}");
            return null;
        }
    }

    private static LambdaExpression? OptionalInputParameterConversion(Type inputType, ExcelParameterRegistration inputParamRegistration, object? defaultValue)
    {
        try
        {
            // if ExcelMissing or ExcelEmpty, return the default value for the type
            var paramExpr = Expression.Parameter(typeof(object), "input");
            var expr = Expression.Condition(
                Expression.Call(typeof(BaseParameterConversion), nameof(CheckIfMissing), null, paramExpr),
                Expression.Constant(defaultValue, inputType),
                GetConversionExpression(paramExpr, inputType));

            return Expression.Lambda(expr, false, new[] { paramExpr });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Failed to register optional input parameter conversion for parameter {inputParamRegistration.ArgumentAttribute.Name}");
            return null;
        }
    }

    private static Expression GetConversionExpression(Expression expression, Type inputType)
    {
        // handle XlBlockRange conversions here
        if (inputType != typeof(XlBlockRange))
            return ParamTypeConverter.GetConversionExpression(expression, inputType);

        var arrayExpr = ParamTypeConverter.GetConversionExpression(expression, typeof(object[,]));
        return Expression.Call(typeof(XlBlockRange), nameof(XlBlockRange.Build), null, arrayExpr);
    }

    private static bool CheckIfMissing(object input)
    {
        if (input is object[] inputArray)
        {
            if (inputArray.Length == 0)
                return true;

            if (inputArray.Length == 1)
                input = inputArray[0];
        }
        return input is ExcelMissing || input is ExcelEmpty;
    }

    internal static T ThrowRequiredParameterException<T>(object input, string paramName)
    {
        throw new RequiredParameterException(paramName);
    }

    public static Func<Type, ExcelReturnRegistration, LambdaExpression?> GetBaseReturnParameterConversion()
    {
        return (type, _) => BaseReturnParameterConversion(type);
    }

    public static LambdaExpression? BaseReturnParameterConversion(Type returnType)
    {
        /* this is the outermost return parameter conversion, and it is subsequently wrapped by the function execution handler registration
         * if a function errors in execution, we will catch it and want to return an ExcelError.NA our execution handler via args.ReturnValue
         * however, the generated lambda that wraps our handler will try to cast FunctionExecutionArgs.ReturnValue to this lambda's return type
         * in order to for that cast to always succeed, we must return an object from our FunctionLambda prior to registering execution handlers
         * so, we just box all non-object return types before giving the lambdas off to the execution handler
         */

        if (returnType == typeof(object))
            return null;

        try
        {
            var paramExpr = Expression.Parameter(returnType, "returnValue");
            var conversionExpr = Expression.Convert(paramExpr, typeof(object));
            return Expression.Lambda(conversionExpr, false, new[] { paramExpr });
        }
        catch (Exception ex)
        {
            _logger.Error(ex, $"Failed to register base return parameter conversion");
            return null;
        }
    }
}
