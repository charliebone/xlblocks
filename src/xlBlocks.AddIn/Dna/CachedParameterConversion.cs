namespace XlBlocks.AddIn.Dna;

using System.Linq.Expressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Registration;
using NLog;
using XlBlocks.AddIn.Cache;
using XlBlocks.AddIn.Types;

[AttributeUsage(AttributeTargets.Parameter | AttributeTargets.ReturnValue, Inherited = false, AllowMultiple = false)]
internal class CacheContentsAttribute : Attribute
{
    /// <summary>
    /// If missing, return the default value for the cached type instead of throwing an exception
    /// </summary>
    public bool DefaultIfMissing { get; set; }

    /// <summary>
    /// If retrieved cached object doesn't match parameter type, return the default value instead of throwing an exception
    /// </summary>
    public bool DefaultIfIncorrectType { get; set; }

    /// <summary>
    /// Fetch copyable objects from the cache by reference instead of copying by value, improving performance. 
    /// Safe to use with operations that do not mutate a cached object
    /// </summary>
    public bool AsReference { get; set; }

    /// <summary>
    /// Include null params elements (bad or missing key) in the array passed to the function
    /// </summary>
    public bool IncludeNullParams { get; set; }

    /// <summary>
    /// Cache the elements of the returned collection and return a collection containing handles to individual cached objects instead
    /// </summary>
    public bool CacheCollectionElements { get; set; }
}

internal class InvalidCacheKeyException : Exception
{
    public string ParameterName { get; }
    public override string Message => $"The input to '{ParameterName}' is invalid";

    public InvalidCacheKeyException(string parameterName)
    {
        ParameterName = parameterName;
    }
}

internal class MissingFromCacheException : Exception
{
    public string ParameterName { get; }
    public string CacheKey { get; }
    public override string Message => $"The input to '{ParameterName}' with handle '{CacheKey}' was not found in the XlBlocks cache";

    public MissingFromCacheException(string parameterName, string cacheKey)
    {
        ParameterName = parameterName;
        CacheKey = cacheKey;
    }
}

internal class IncorrectTypeCacheException : Exception
{
    public string ParameterName { get; }
    public string CacheKey { get; }
    public Type ExpectedType { get; }
    public override string Message => $"The input to '{ParameterName}' with handle '{CacheKey}' was not of the expected type '{ExpectedType.Name}'";

    public IncorrectTypeCacheException(string parameterName, string cacheKey, Type expectedType)
    {
        ParameterName = parameterName;
        CacheKey = cacheKey;
        ExpectedType = expectedType;
    }
}

internal static class CachedParameterConversion
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(CachedParameterConversion).FullName);

    public static Func<Type, ExcelParameterRegistration, LambdaExpression?> GetCachedInputParameterConversion()
    {
        return (type, inputParamRegistration) => CachedInputParameterConversion(type, inputParamRegistration);
    }

    internal static LambdaExpression? CachedInputParameterConversion(Type inputType, ExcelParameterRegistration inputParamRegistration)
    {
        if (!inputParamRegistration.CustomAttributes.OfType<CacheContentsAttribute>().Any())
            return null;

        // ignore params style args, they will be registered later
        if (inputParamRegistration.CustomAttributes.OfType<ParamArrayAttribute>().Any())
            return null;

        // when an input parameter is marked as CacheContents, use the input string to fetch it from the cache and return the cached object
        var cachedParamAttribute = inputParamRegistration.CustomAttributes.OfType<CacheContentsAttribute>().First();
        var defaultAttribute = inputParamRegistration.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
        var defaultValue = defaultAttribute == null || defaultAttribute.Value is System.Reflection.Missing ?
                ParamTypeConverter.GetDefault(inputType) :
                defaultAttribute.Value;

        var paramExpr = Expression.Parameter(typeof(object), "inputCacheKey");
        var cacheAttributeExpr = Expression.Constant(cachedParamAttribute);
        var paramNameExpr = Expression.Constant(inputParamRegistration.ArgumentAttribute.Name);
        var defaultValueExpr = Expression.Constant(defaultValue, inputType);

        var expr = Expression.Call(typeof(CachedParameterConversion), nameof(GetCacheValueOrDefault), new[] { inputType },
                paramExpr, cacheAttributeExpr, paramNameExpr, defaultValueExpr);
        return Expression.Lambda(expr, false, new[] { paramExpr });
    }

    private static T? GetCacheValueOrDefault<T>(object? inputCacheKey, CacheContentsAttribute cachedParamAttribute, string paramName, T? defaultValue)
    {
        if (inputCacheKey is null || inputCacheKey is ExcelMissing || inputCacheKey is ExcelEmpty || inputCacheKey is ExcelError)
        {
            if (!cachedParamAttribute.DefaultIfMissing)
                throw new InvalidCacheKeyException(paramName);

            return defaultValue;
        }
        var hexKey = inputCacheKey.ToString();
        if (hexKey is null)
            return defaultValue;

        if (XlBlocksAddIn.Instance.Cache.TryGetValue(hexKey, out var value))
        {
            if (value is T tVal)
                return (!cachedParamAttribute.AsReference && tVal is IXlBlockCopyableObject<T> copyableTVal) ?
                    copyableTVal.Copy() : tVal;

            if (!cachedParamAttribute.DefaultIfIncorrectType)
                throw new IncorrectTypeCacheException(paramName, hexKey, typeof(T));

            return defaultValue;
        }

        if (!cachedParamAttribute.DefaultIfMissing)
            throw new MissingFromCacheException(paramName, hexKey);

        return defaultValue;
    }

    public static Func<Type, ExcelReturnRegistration, LambdaExpression?> GetCachedReturnParameterConversion()
    {
        return (type, returnParamRegistration) => CachedReturnParameterConversion(type, returnParamRegistration);
    }

    internal static LambdaExpression? CachedReturnParameterConversion(Type _, ExcelReturnRegistration returnParamRegistration)
    {
        var cachedParamAttribute = returnParamRegistration.CustomAttributes.OfType<CacheContentsAttribute>().FirstOrDefault();
        if (cachedParamAttribute is null)
            return null;

        // when a return attribute is marked as a CachedParameter, cache it using the calling reference as a key and then return the cache key
        return (object outputToCache) => SetCacheValueOrDefault(outputToCache, cachedParamAttribute.CacheCollectionElements);
    }

    private static object SetCacheValueOrDefault(object outputToCache, bool cacheCollectionElements)
    {
        // if calling the excel function from a macro, generate a GUID to use as ref string and use that to set the cache
        var reference = CacheHelper.GetCallingReference(() => Guid.NewGuid().ToString());

        if (cacheCollectionElements && outputToCache is IXlBlockCacheableCollection cacheableCollection)
        {
            var elementCacher = XlBlocksAddIn.Instance.Cache.GetElementCacher(reference);
            var collection = cacheableCollection.CacheCollectionValues(elementCacher);

            XlBlocksAddIn.Instance.Cache.AddOrReplace(reference, collection, out var hexKey);
            return hexKey;
        }
        else
        {
            XlBlocksAddIn.Instance.Cache.AddOrReplace(reference, outputToCache, out var hexKey);
            return hexKey;
        }
    }

    /*
     * NOTE: the below functionality is almost entirely taken from the brilliant ProcessParamsRegistrations example in ExcelDna.Registration.
     * It can be found at https://github.com/Excel-DNA/Registration/blob/master/Source/ExcelDna.Registration/ParamsRegistration.cs
     */
    public static IEnumerable<ExcelFunctionRegistration> ProcessCacheAwareParamsRegistrations(this IEnumerable<ExcelFunctionRegistration> registrations)
    {
        foreach (var reg in registrations)
        {
            try
            {
                if (IsParamsMethod(reg))
                {
                    var lastParam = reg.ParameterRegistrations.Last();
                    var cacheContentsAttribute = lastParam.CustomAttributes.OfType<CacheContentsAttribute>().FirstOrDefault();
                    var isOptional = lastParam.CustomAttributes.OfType<OptionalAttribute>().FirstOrDefault() is not null;

                    reg.FunctionLambda = WrapCacheAwareMethodParams(reg.FunctionLambda, cacheContentsAttribute, lastParam.ArgumentAttribute.Name, isOptional);

                    // Clean out ParamArray attribute for the last parameter (there will be one)
                    lastParam.CustomAttributes.RemoveAll(att => att is ParamArrayAttribute);

                    // Add more attributes for the 'params' arguments
                    // Adjust the first one from myInput to myInput1
                    var paramsArgAttrib = lastParam.ArgumentAttribute;
                    paramsArgAttrib.Name += "1";

                    // Add the ellipse argument
                    reg.ParameterRegistrations.Add(
                        new ExcelParameterRegistration(
                            new ExcelArgumentAttribute
                            {
                                Name = "...",
                                Description = paramsArgAttrib.Description,
                                AllowReference = paramsArgAttrib.AllowReference
                            }));

                    // And the rest with no Name, but copying the description
                    var restCount = reg.FunctionLambda.Parameters.Count - reg.ParameterRegistrations.Count;
                    for (var i = 0; i < restCount; i++)
                    {
                        reg.ParameterRegistrations.Add(
                            new ExcelParameterRegistration(
                                new ExcelArgumentAttribute
                                {
                                    Name = string.Empty,
                                    Description = paramsArgAttrib.Description,
                                    AllowReference = paramsArgAttrib.AllowReference
                                }));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Exception while registering method {reg.FunctionAttribute.Name}");
                continue;
            }
            yield return reg;
        }
    }

    public delegate TResult CustomFunc29<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19, T20, T21, T22, T23, T24, T25, T26, T27, T28, T29, TResult>
                                    (T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6, T7 arg7, T8 arg8, T9 arg9, T10 arg10, T11 arg11, T12 arg12, T13 arg13, T14 arg14, T15 arg15, T16 arg16, T17 arg17, T18 arg18, T19 arg19, T20 arg20, T21 arg21, T22 arg22, T23 arg23, T24 arg24, T25 arg25, T26 arg26, T27 arg27, T28 arg28, T29 arg29);
    public delegate TResult CustomFunc125<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19, T20, T21, T22, T23, T24, T25, T26, T27, T28, T29, T30, T31, T32, T33, T34, T35, T36, T37, T38, T39, T40, T41, T42, T43, T44, T45, T46, T47, T48, T49, T50, T51, T52, T53, T54, T55, T56, T57, T58, T59, T60, T61, T62, T63, T64, T65, T66, T67, T68, T69, T70, T71, T72, T73, T74, T75, T76, T77, T78, T79, T80, T81, T82, T83, T84, T85, T86, T87, T88, T89, T90, T91, T92, T93, T94, T95, T96, T97, T98, T99, T100, T101, T102, T103, T104, T105, T106, T107, T108, T109, T110, T111, T112, T113, T114, T115, T116, T117, T118, T119, T120, T121, T122, T123, T124, T125, TResult>
                                    (T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5 arg5, T6 arg6, T7 arg7, T8 arg8, T9 arg9, T10 arg10, T11 arg11, T12 arg12, T13 arg13, T14 arg14, T15 arg15, T16 arg16, T17 arg17, T18 arg18, T19 arg19, T20 arg20, T21 arg21, T22 arg22, T23 arg23, T24 arg24, T25 arg25, T26 arg26, T27 arg27, T28 arg28, T29 arg29, T30 arg30, T31 arg31, T32 arg32, T33 arg33, T34 arg34, T35 arg35, T36 arg36, T37 arg37, T38 arg38, T39 arg39, T40 arg40, T41 arg41, T42 arg42, T43 arg43, T44 arg44, T45 arg45, T46 arg46, T47 arg47, T48 arg48, T49 arg49, T50 arg50, T51 arg51, T52 arg52, T53 arg53, T54 arg54, T55 arg55, T56 arg56, T57 arg57, T58 arg58, T59 arg59, T60 arg60, T61 arg61, T62 arg62, T63 arg63, T64 arg64, T65 arg65, T66 arg66, T67 arg67, T68 arg68, T69 arg69, T70 arg70, T71 arg71, T72 arg72, T73 arg73, T74 arg74, T75 arg75, T76 arg76, T77 arg77, T78 arg78, T79 arg79, T80 arg80, T81 arg81, T82 arg82, T83 arg83, T84 arg84, T85 arg85, T86 arg86, T87 arg87, T88 arg88, T89 arg89, T90 arg90, T91 arg91, T92 arg92, T93 arg93, T94 arg94, T95 arg95, T96 arg96, T97 arg97, T98 arg98, T99 arg99, T100 arg100, T101 arg101, T102 arg102, T103 arg103, T104 arg104, T105 arg105, T106 arg106, T107 arg107, T108 arg108, T109 arg109, T110 arg110, T111 arg111, T112 arg112, T113 arg113, T114 arg114, T115 arg115, T116 arg116, T117 arg117, T118 arg118, T119 arg119, T120 arg120, T121 arg121, T122 arg122, T123 arg123, T124 arg124, T125 arg125);

    private static LambdaExpression WrapCacheAwareMethodParams(LambdaExpression functionLambda, CacheContentsAttribute? cacheAttribute, string paramName, bool isOptional)
    {
        /* We are converting:
            *     [ExcelFunction(...)]
            *     public static string myFunc(string input, int otherInput, params object[] args)
            *     {    
            *          ...
            *     }
            * 
            * into:
            *     [ExcelFunction(...)]
            *     public static string myFunc(string input, int otherInput, object arg3, object arg4, object arg5, object arg6, {...until...}, object arg125)
            *     {
            *         // First we figure where in the list to stop building the param array
            *         int lastArgToAdd = 0;
            *         if (!(arg3 is ExcelMissing)) lastArgToAdd = 3;
            *         if (!(arg4 is ExcelMissing)) lastArgToAdd = 4;
            *         ...
            *         if (!(arg125 is ExcelMissing)) lastArgToAdd = 125;
            *     
            *         // Then add until we get there
            *         List<object> args = new List<object>();
            *         if (lastArgToAdd >= 3) args.Add(arg3);
            *         if (lastArgToAdd >= 4) args.Add(arg4);
            *         ...
            *         if (lastArgToAdd >= 125) args.Add(arg125);
            *        
            *         Array<object> argsArray = args.ToArray();
            *         return myFunc(input, otherInput, argsArray);
            *     }
            * 
            * 
            */

        var maxArguments = ExcelDnaUtil.ExcelVersion >= 12.0 ? 125 : 29;
        var normalParams = functionLambda.Parameters.Take(functionLambda.Parameters.Count - 1).ToList();
        var normalParamCount = normalParams.Count;
        var paramsParamCount = maxArguments - normalParamCount;
        var allParamExprs = new List<ParameterExpression>(normalParams);
        var blockExprs = new List<Expression>();
        var blockVars = new List<ParameterExpression>();
        var cacheAttributeExpr = Expression.Constant(cacheAttribute);
        var includeNullParamsExpr = Expression.Constant(cacheAttribute?.IncludeNullParams ?? false, typeof(bool));
        var paramNameExpr = Expression.Constant(paramName);

        // Run through the arguments looking for the position of the last non-ExcelMissing argument
        var lastArgVarExpr = Expression.Variable(typeof(int));
        blockVars.Add(lastArgVarExpr);
        blockExprs.Add(Expression.Assign(lastArgVarExpr, Expression.Constant(0)));
        for (var i = normalParamCount + 1; i <= maxArguments; i++)
        {
            allParamExprs.Add(Expression.Parameter(typeof(object), $"arg{i}"));

            var lenTestParam = Expression.IfThen(Expression.Not(Expression.TypeIs(allParamExprs[i - 1], typeof(ExcelMissing))),
                                Expression.Assign(lastArgVarExpr, Expression.Constant(i)));
            blockExprs.Add(lenTestParam);
        }

        // We know that last parameter is an array type
        // Create a new list to hold the values
        var argsArrayType = functionLambda.Parameters.Last().Type;
        var argsType = argsArrayType.GetElementType() ?? throw new Exception($"invalid params array element type");
        var nullableArgsType = argsType.IsValueType ? typeof(Nullable<>).MakeGenericType(argsType) : argsType;
        var defaultValueExpr = Expression.Constant(ParamTypeConverter.GetDefault(argsType), nullableArgsType);
        var argsTypeArray = new[] { argsType };
        var nullableArgsTypeArray = new[] { nullableArgsType };
        var argsListType = typeof(List<>).MakeGenericType(argsType);

        var argsListVarExpr = Expression.Variable(argsListType);
        blockVars.Add(argsListVarExpr);

        var argListAssignExpr = Expression.Assign(argsListVarExpr, Expression.New(argsListType));
        blockExprs.Add(argListAssignExpr);

        for (var i = normalParamCount + 1; i <= maxArguments; i++)
        {
            if (!isOptional && i == 1)
            {
                // if params parameter is not marked optional, ensure we have at least one non-missing arg
                var reqParamExpr = Expression.IfThen(Expression.Equal(lastArgVarExpr, Expression.Constant(0)),
                    Expression.Call(typeof(BaseParameterConversion), nameof(BaseParameterConversion.ThrowRequiredParameterException),
                        argsTypeArray, allParamExprs[i - 1], paramNameExpr));
                blockExprs.Add(reqParamExpr);
            }

            Expression testParam;
            if (cacheAttribute is not null)
            {
                var resultVar = Expression.Variable(nullableArgsType, "result");
                var getValueCall = Expression.Call(typeof(CachedParameterConversion), nameof(GetCacheValueOrDefault), nullableArgsTypeArray,
                                        allParamExprs[i - 1], cacheAttributeExpr, paramNameExpr, defaultValueExpr);
                var assignResultExpr = Expression.Assign(resultVar, getValueCall);

                var tryGetExpr = Expression.Block(
                    new[] { resultVar },
                    assignResultExpr,
                    Expression.IfThen(
                        Expression.OrElse(
                            includeNullParamsExpr,
                            Expression.NotEqual(resultVar, Expression.Constant(null, nullableArgsType))),
                        Expression.Call(argsListVarExpr, "Add", null,
                            Expression.Convert(argsType.IsValueType ? Expression.Property(resultVar, "Value") : resultVar, argsType))));

                testParam = Expression.IfThen(
                    Expression.GreaterThanOrEqual(lastArgVarExpr, Expression.Constant(i)),
                    tryGetExpr);
            }
            else
            {
                testParam = Expression.IfThen(
                    Expression.GreaterThanOrEqual(lastArgVarExpr, Expression.Constant(i)),
                    Expression.Call(argsListVarExpr, "Add", null, ParamTypeConverter.GetConversionExpression(allParamExprs[i - 1], argsType)));
            }
            blockExprs.Add(testParam);
        }

        var argArrayVarExpr = Expression.Variable(argsArrayType);
        blockVars.Add(argArrayVarExpr);

        var argArrayAssignExpr = Expression.Assign(argArrayVarExpr, Expression.Call(argsListVarExpr, "ToArray", null));
        blockExprs.Add(argArrayAssignExpr);

        var innerParams = new List<Expression>(normalParams) { argArrayVarExpr };
        var callInner = Expression.Invoke(functionLambda, innerParams);
        blockExprs.Add(callInner);

        var blockExpr = Expression.Block(blockVars, blockExprs);

        // Build the delegate type to return
        var allParamTypes = normalParams.Select(pi => pi.Type).ToList();
        var toAdd = maxArguments - allParamTypes.Count;
        for (var i = 0; i < toAdd; i++)
        {
            allParamTypes.Add(typeof(object));
        }
        allParamTypes.Add(functionLambda.ReturnType);

        var delegateType = maxArguments == 125 ?
            typeof(CustomFunc125<,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,>)
                                .MakeGenericType(allParamTypes.ToArray()) :
            typeof(CustomFunc29<,,,,,,,,,,,,,,,,,,,,,,,,,,,,,>)
                                .MakeGenericType(allParamTypes.ToArray());

        return Expression.Lambda(delegateType, blockExpr, allParamExprs);
    }

    static bool IsParamsMethod(ExcelFunctionRegistration reg)
    {
        var lastParam = reg.ParameterRegistrations.LastOrDefault();
        return lastParam != null && lastParam.CustomAttributes.Any(att => att is ParamArrayAttribute)
            && reg.FunctionLambda.Parameters.Last().Type.IsArray;
    }
}
