namespace XlBlocks.AddIn.Excel.Tests;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;

public static class CachedParameterConversionTests
{
    #region CachedParameterConversion Defaults
    [ExcelFunction(Description = "Test default behavior for cached object input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_Default_Object(
        [ExcelArgument(Description = "The cached object"), CacheContents] object cachedObject)
    {
        return cachedObject;
    }

    [ExcelFunction(Description = "Test default behavior for cached string input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_Default_String(
        [ExcelArgument(Description = "The cached object"), CacheContents] string cachedString)
    {
        return cachedString;
    }

    [ExcelFunction(Description = "Test default behavior for cached double input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_Default_Double(
        [ExcelArgument(Description = "The cached object"), CacheContents] double cachedDouble)
    {
        return cachedDouble;
    }
    #endregion

    #region CachedParameterConversion DefaultIfMissing tests
    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfMissing_Object(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfMissing = true)] object cachedObject)
    {
        return cachedObject;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfMissing_String(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfMissing = true)] string cachedString)
    {
        return cachedString;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfMissing_Double(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfMissing = true)] double cachedDouble)
    {
        return cachedDouble;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfMissing_StringWithDefault(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfMissing = true)] string cachedString = "default_string")
    {
        return cachedString;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfMissing_DoubleWithDefault(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfMissing = true)] double cachedDouble = 123d)
    {
        return cachedDouble;
    }
    #endregion

    #region CachedParameterConversion DefaultIfIncorrectType tests
    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfIncorrectType_String(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfIncorrectType = true)] string cachedString)
    {
        return cachedString;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached input parameter in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfIncorrectType_Double(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfIncorrectType = true)] double cachedDouble)
    {
        return cachedDouble;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached input parameter with default value in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfIncorrectType_StringWithDefault(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfIncorrectType = true)] string cachedString = "default_string")
    {
        return cachedString;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached input parameter with default valuein CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_DefaultIfIncorrectType_DoubleWithDefault(
        [ExcelArgument(Description = "The cached object"), CacheContents(DefaultIfIncorrectType = true)] double cachedDouble = 123d)
    {
        return cachedDouble;
    }
    #endregion

    #region CachedParameterConversion params style parameter tests
    [ExcelFunction(Description = "Test required parameter check for non-cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_NonCachedParams_Required(
        [ExcelArgument(Description = "The params argument")] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test optional parameter behavior for non-cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_NonCachedParams_Optional(
        [ExcelArgument(Description = "The params argument"), Optional] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test default behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_Default_Object(
        [ExcelArgument(Description = "The params argument"), CacheContents] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test default behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_Default_String(
        [ExcelArgument(Description = "The params argument"), CacheContents] params string[] stringArray)
    {
        return stringArray;
    }

    [ExcelFunction(Description = "Test default behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_Default_Double(
        [ExcelArgument(Description = "The params argument"), CacheContents] params double[] doubleArray)
    {
        return doubleArray;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for required cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(
        [ExcelArgument(Description = "The params argument"), CacheContents(DefaultIfMissing = true)] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test DefaultIfMissing=true behavior for optional cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(
        [ExcelArgument(Description = "The params argument"), CacheContents(DefaultIfMissing = true), Optional] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_Object(
        [ExcelArgument(Description = "The params argument"), CacheContents(DefaultIfIncorrectType = true)] params object[] objArray)
    {
        return objArray;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_String(
        [ExcelArgument(Description = "The params argument"), CacheContents(DefaultIfIncorrectType = true)] params string[] stringArray)
    {
        return stringArray;
    }

    [ExcelFunction(Description = "Test DefaultifIncorrectType=true behavior for cached params args in CachedParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_Double(
        [ExcelArgument(Description = "The params argument"), CacheContents(DefaultIfIncorrectType = true)] params double[] doubleArray)
    {
        return doubleArray;
    }
    #endregion
}
