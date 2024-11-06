namespace XlBlocks.AddIn.Integration.Tests;

using ExcelDna.Integration;

public partial class IntegrationTests
{
    public class BaseParameterConversionTests : IntegrationFunctionTests
    {
        /// <summary>
        /// Test required parameters (the default case) for BaseParameterConversions
        /// </summary>
        [ExcelFact]
        public void BaseParameterConversion_RequiredParameter()
        {
            var ws = _testWorkbook.Sheets.Add();

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=XBTests_BaseParameterConversion_RequiredParameter()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A2"].Value);
            ws.Range["A3"].Formula = "=XBTests_BaseParameterConversion_RequiredParameter(B3)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_BaseParameterConversion_RequiredParameter()";
            Assert.Equal("Parameter 'requiredParameter' cannot be missing or empty", ws.Range["A4"].Value);
            ws.Range["A5"].Formula = "=XBTests_BaseParameterConversion_RequiredParameter(B5)";
            Assert.Equal("Parameter 'requiredParameter' cannot be missing or empty", ws.Range["A5"].Value);

            ws.Range["A6"].Value = "parameter_value";
            ws.Range["B6"].Formula = "=XBTests_BaseParameterConversion_RequiredParameter(A6)";
            Assert.Equal(ws.Range["A6"].Value, ws.Range["B6"].Value);
        }

        /// <summary>
        /// Test optional parameters indicated via OptionalAttribute on argument for BaseConversions
        /// </summary>
        [ExcelFact]
        public void BaseParameterConversion_OptionalParameterByAttr()
        {
            var ws = _testWorkbook.Sheets.Add();

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByAttr()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A2"].Value);
            ws.Range["A3"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByAttr(B3)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByAttr()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A4"].Value);
            ws.Range["A5"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByAttr(B5)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A5"].Value);

            ws.Range["A6"].Value = "parameter_value";
            ws.Range["B6"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByAttr(A6)";
            Assert.Equal(ws.Range["A6"].Value, ws.Range["B6"].Value);
        }

        /// <summary>
        /// Test optional parameters indicated via default argument for BaseConversions
        /// </summary>
        [ExcelFact]
        public void BaseParameterConversion_OptionalParameterByDefault()
        {
            var ws = _testWorkbook.Sheets.Add();

            var parameterDefaultValue = "default_value";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByDefault()";
            Assert.Equal(parameterDefaultValue, ws.Range["A2"].Value);
            ws.Range["A3"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByDefault(B3)";
            Assert.Equal(parameterDefaultValue, ws.Range["A3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByDefault()";
            Assert.Equal(parameterDefaultValue, ws.Range["A4"].Value);
            ws.Range["A5"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByDefault(B5)";
            Assert.Equal(parameterDefaultValue, ws.Range["A5"].Value);

            ws.Range["A6"].Value = "another_value";
            ws.Range["B6"].Formula = "=XBTests_BaseParameterConversion_OptionalParameterByDefault(A6)";
            Assert.Equal(ws.Range["A6"].Value, ws.Range["B6"].Value);
        }
    }
}
