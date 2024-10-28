namespace XlBlocks.AddIn.Excel.Tests;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlBlocks.AddIn.Dna;

public static class BaseParameterConversionTests
{
    [ExcelFunction(Description = "Test required parameter check in BaseParameterConverion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_BaseParameterConversion_RequiredParameter(
        [ExcelArgument(Description = "Required parameter")] object requiredParameter)
    {
        return requiredParameter;
    }

    [ExcelFunction(Description = "Test optional parameter check in BaseParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_BaseParameterConversion_OptionalParameterByAttr(
        [ExcelArgument(Description = "Optional parameter"), Optional] string optionalParameter)
    {
        return optionalParameter;
    }

    [ExcelFunction(Description = "Test optional parameter check in BaseParameterConversion", IsThreadSafe = true), IntegrationTestExcelFunction]
    public static object XBTests_BaseParameterConversion_OptionalParameterByDefault(
        [ExcelArgument(Description = "Optional parameter")] string paramOptDef = "default_value")
    {
        return paramOptDef;
    }
}
