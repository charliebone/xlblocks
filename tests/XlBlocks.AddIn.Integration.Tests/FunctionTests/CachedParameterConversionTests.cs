namespace XlBlocks.AddIn.Integration.Tests;

using ExcelDna.Integration;

public partial class IntegrationTests
{
    public class CachedParameterConversionTests : IntegrationFunctionTests
    {
        /// <summary>
        /// Test DefaultIfMissing=FALSE on missing from cache (the default case) for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_ThrowIfMissing()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["A1"].Value = badHandle;

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=XBTests_CachedParameterConversion_Default_Object()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A2"].Value);
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_Default_Object(A1)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_CachedParameterConversion_Default_Object()";
            Assert.Equal("Parameter 'cachedObject' cannot be missing or empty", ws.Range["A4"].Value);
            ws.Range["A5"].Formula = "=XBTests_CachedParameterConversion_Default_Object(A1)";
            Assert.Equal($"The input to 'cachedObject' with handle '{badHandle}' was not found in the XlBlocks cache", ws.Range["A5"].Value);

            var cacheValue = "some_value";
            ws.Range["A6"].Value = cacheValue;
            ws.Range["B6"].Formula = "=XBCache_Set(A6)";
            ws.Range["C6"].Formula = "=XBTests_CachedParameterConversion_Default_Object(B6)";
            Assert.Equal(cacheValue, ws.Range["C6"].Value);
        }

        /// <summary>
        /// Test DefaultIfMissing=TRUE on missing from cache for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_DefaultIfMissing()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["A1"].Value = badHandle;

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_Double()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A2"].Value);
            ws.Range["B2"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_Double(A1)";
            Assert.Equal(0d, ws.Range["B2"].Value);
            ws.Range["C2"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_String(A1)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["C2"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_Double()";
            Assert.Equal("Parameter 'cachedDouble' cannot be missing or empty", ws.Range["A3"].Value);
            ws.Range["B3"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_Double(A1)";
            Assert.Equal(0, ws.Range["B3"].Value);
            ws.Range["C3"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_String(A1)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["C3"].Value);

            var doubleParameterDefault = 123d;
            var stringParameterDefault = "default_string";
            ws.Range["A4"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_DoubleWithDefault(A1)";
            Assert.Equal(doubleParameterDefault, ws.Range["A4"].Value);
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_StringWithDefault(A1)";
            Assert.Equal(stringParameterDefault, ws.Range["B4"].Value);

            var cacheValue = 123d;
            ws.Range["A6"].Value = cacheValue;
            ws.Range["B6"].Formula = "=XBCache_Set(A6)";
            ws.Range["C6"].Formula = "=XBTests_CachedParameterConversion_DefaultIfMissing_Double(B6)";
            Assert.Equal(cacheValue, ws.Range["C6"].Value);
        }

        /// <summary>
        /// Test DefaultIfIncorrectType=FALSE (the default case) for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_ThrowIfIncorrectType()
        {
            var ws = _testWorkbook.Sheets.Add();

            var stringValue = "string_value";
            var doubleValue = 123d;
            ws.Range["A1"].Value = stringValue;
            ws.Range["B1"].Formula = $"={doubleValue}*1";
            ws.Range["A2"].Formula = "=XBCache_Set(A1)";
            ws.Range["B2"].Formula = "=XBCache_Set(B1)";

            string handleString = ws.Range["A2"].Value;
            string handleDouble = ws.Range["B2"].Value;

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_Default_Double(A2)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A3"].Value);
            ws.Range["B3"].Formula = "=XBTests_CachedParameterConversion_Default_String(B2)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_CachedParameterConversion_Default_Double(A2)";
            Assert.Equal($"The input to 'cachedDouble' with handle '{handleString}' was not of the expected type 'Double'", ws.Range["A4"].Value);
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_Default_String(B2)";
            Assert.Equal($"The input to 'cachedString' with handle '{handleDouble}' was not of the expected type 'String'", ws.Range["B4"].Value);
        }

        /// <summary>
        /// Test DefaultIfIncorrectType=TRUE for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_DefaultIfIncorrectType()
        {
            var ws = _testWorkbook.Sheets.Add();

            var stringValue = "string_value";
            var doubleValue = 123d;
            ws.Range["A1"].Value = stringValue;
            ws.Range["B1"].Formula = $"={doubleValue}*1";
            ws.Range["A2"].Formula = "=XBCache_Set(A1)";
            ws.Range["B2"].Formula = "=XBCache_Set(B1)";

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_Double(A2)";
            Assert.Equal(0d, ws.Range["A3"].Value);
            ws.Range["B3"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_String(B2)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A4"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_Double(A2)";
            Assert.Equal(0d, ws.Range["A4"].Value);
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_String(B2)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B4"].Value);

            var doubleParameterDefault = 123d;
            var stringParameterDefault = "default_string";
            ws.Range["A5"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_DoubleWithDefault(A2)";
            Assert.Equal(doubleParameterDefault, ws.Range["A5"].Value);
            ws.Range["B5"].Formula = "=XBTests_CachedParameterConversion_DefaultIfIncorrectType_StringWithDefault(B2)";
            Assert.Equal(stringParameterDefault, ws.Range["B5"].Value);
        }

        /// <summary>
        /// Test required (the default case) non-cached params for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_NonCachedParams_Required()
        {
            var ws = _testWorkbook.Sheets.Add();

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A1"].Formula = "=XBTests_CachedParameterConversion_NonCachedParams_Required()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A1"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A2"].Formula = "=XBTests_CachedParameterConversion_NonCachedParams_Required()";
            Assert.Equal("Parameter 'objArray' cannot be missing or empty", ws.Range["A2"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A3"].Value = cachedValues[0];
            ws.Range["A4"].Value = cachedValues[1];
            ws.Range["A5"].Value = cachedValues[2];
            ws.Range["A6"].Value = cachedValues[3];

            ws.Range["B3:B6"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_NonCachedParams_Required(A3,A4,A5,A6))";
            Assert.Equal(ws.Range["A3"].Value, ws.Range["B3"].Value);
            Assert.Equal(ws.Range["A4"].Value, ws.Range["B4"].Value);
            Assert.Equal(ws.Range["A5"].Value, ws.Range["B5"].Value);
            Assert.Equal(ws.Range["A6"].Value2, ws.Range["B6"].Value2);
        }

        /// <summary>
        /// Test optional non-cached params for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_NonCachedParams_Optional()
        {
            var ws = _testWorkbook.Sheets.Add();

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A1"].Formula = "=XBTests_CachedParameterConversion_NonCachedParams_Optional()";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["A1"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A2"].Formula = "=XBTests_CachedParameterConversion_NonCachedParams_Optional()";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["A2"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A3"].Value = cachedValues[0];
            ws.Range["A4"].Value = cachedValues[1];
            ws.Range["A5"].Value = cachedValues[2];
            ws.Range["A6"].Value = cachedValues[3];

            ws.Range["B3:B6"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_NonCachedParams_Optional(A3,A4,A5,A6))";
            Assert.Equal(ws.Range["A3"].Value, ws.Range["B3"].Value);
            Assert.Equal(ws.Range["A4"].Value, ws.Range["B4"].Value);
            Assert.Equal(ws.Range["A5"].Value, ws.Range["B5"].Value);
            Assert.Equal(ws.Range["A6"].Value2, ws.Range["B6"].Value2);
        }

        /// <summary>
        /// Test default behavior for cached params for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_CachedParams_ThrowIfMissing()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A1"].Formula = "=XBTests_CachedParameterConversion_CachedParams_Default_Object()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A1"].Value);
            ws.Range["A2"].Value = badHandle;
            ws.Range["B2"].Formula = "=XBTests_CachedParameterConversion_CachedParams_Default_Object(A2)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B2"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_CachedParams_Default_Object()";
            Assert.Equal("Parameter 'objArray' cannot be missing or empty", ws.Range["A3"].Value);
            ws.Range["A4"].Value = badHandle;
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_CachedParams_Default_Object(A4)";
            Assert.Equal($"The input to 'objArray' with handle '{badHandle}' was not found in the XlBlocks cache", ws.Range["B4"].Value);

            var goodValue = "good_value";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A5"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A6"].Value = badHandle;
            ws.Range["B5:B6"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_Object(A5,A6))";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B5"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B6"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A7"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A8"].Value = badHandle;
            ws.Range["B7:B8"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_Object(A7,A8))";
            Assert.Equal($"The input to 'objArray' with handle '{badHandle}' was not found in the XlBlocks cache", ws.Range["B7"].Value);
            Assert.Equal($"The input to 'objArray' with handle '{badHandle}' was not found in the XlBlocks cache", ws.Range["B8"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A9"].Value = cachedValues[0];
            ws.Range["B9"].Formula = "=XBCache_Set(A9)";
            ws.Range["A10"].Value = cachedValues[1];
            ws.Range["B10"].Formula = "=XBCache_Set(A10)";
            ws.Range["A11"].Value = cachedValues[2];
            ws.Range["B11"].Formula = "=XBCache_Set(A11)";
            ws.Range["A12"].Value = cachedValues[3];
            ws.Range["B12"].Formula = "=XBCache_Set(A12)";

            ws.Range["C9:C12"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_Object(B9,B10,B11,B12))";
            Assert.Equal(ws.Range["A9"].Value, ws.Range["C9"].Value);
            Assert.Equal(ws.Range["A10"].Value, ws.Range["C10"].Value);
            Assert.Equal(ws.Range["A11"].Value, ws.Range["C11"].Value);
            Assert.Equal(ws.Range["A12"].Value2, ws.Range["C12"].Value2);
        }

        /// <summary>
        /// Test DefaultIfMissing=true with required cached params arg for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_CachedParams_DefaultIfMissing_Required()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A1"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required()";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A1"].Value);
            ws.Range["A2"].Value = badHandle;
            ws.Range["B2"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(A2)";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["B2"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required()";
            Assert.Equal("Parameter 'objArray' cannot be missing or empty", ws.Range["A3"].Value);
            ws.Range["A4"].Value = badHandle;
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(A4)";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["B4"].Value);

            var goodValue = "good_value";
            var goodValue2 = "another_value";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A5"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A6"].Value = badHandle;
            ws.Range["A7"].Formula = $"=XBCache_Set(\"{goodValue2}\")";
            ws.Range["B5:B7"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(A5,A6,A7))";
            Assert.Equal(goodValue, ws.Range["B5"].Value);
            Assert.Equal(goodValue2, ws.Range["B6"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B7"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A8"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A9"].Value = badHandle;
            ws.Range["A10"].Formula = $"=XBCache_Set(\"{goodValue2}\")";
            ws.Range["B8:B10"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(A8,A9,A10))";
            Assert.Equal(goodValue, ws.Range["B8"].Value);
            Assert.Equal(goodValue2, ws.Range["B9"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B10"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A11"].Value = cachedValues[0];
            ws.Range["B11"].Formula = "=XBCache_Set(A11)";
            ws.Range["A12"].Value = cachedValues[1];
            ws.Range["B12"].Formula = "=XBCache_Set(A12)";
            ws.Range["A13"].Value = cachedValues[2];
            ws.Range["B13"].Formula = "=XBCache_Set(A13)";
            ws.Range["A14"].Value = cachedValues[3];
            ws.Range["B14"].Formula = "=XBCache_Set(A14)";

            ws.Range["C11:C14"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Required(B11,B12,B13,B14))";
            Assert.Equal(ws.Range["A11"].Value, ws.Range["C11"].Value);
            Assert.Equal(ws.Range["A12"].Value, ws.Range["C12"].Value);
            Assert.Equal(ws.Range["A13"].Value, ws.Range["C13"].Value);
            Assert.Equal(ws.Range["A14"].Value2, ws.Range["C14"].Value2);
        }

        /// <summary>
        /// Test DefaultIfMissing=true with optional cached params arg for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_CachedParams_DefaultIfMissing_Optional()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A1"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional()";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["A1"].Value);
            ws.Range["A2"].Value = badHandle;
            ws.Range["B2"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(A2)";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["B2"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A3"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional()";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["A3"].Value);
            ws.Range["A4"].Value = badHandle;
            ws.Range["B4"].Formula = "=XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(A4)";
            AssertIsExcelError(ExcelError.ExcelErrorValue, ws.Range["B4"].Value);

            var goodValue = "good_value";
            var goodValue2 = "another_value";
            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["A5"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A6"].Value = badHandle;
            ws.Range["A7"].Formula = $"=XBCache_Set(\"{goodValue2}\")";
            ws.Range["B5:B7"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(A5,A6,A7))";
            Assert.Equal(goodValue, ws.Range["B5"].Value);
            Assert.Equal(goodValue2, ws.Range["B6"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B7"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["A8"].Formula = $"=XBCache_Set(\"{goodValue}\")";
            ws.Range["A9"].Value = badHandle;
            ws.Range["A10"].Formula = $"=XBCache_Set(\"{goodValue2}\")";
            ws.Range["B8:B10"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(A8,A9,A10))";
            Assert.Equal(goodValue, ws.Range["B8"].Value);
            Assert.Equal(goodValue2, ws.Range["B9"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["B10"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A11"].Value = cachedValues[0];
            ws.Range["B11"].Formula = "=XBCache_Set(A11)";
            ws.Range["A12"].Value = cachedValues[1];
            ws.Range["B12"].Formula = "=XBCache_Set(A12)";
            ws.Range["A13"].Value = cachedValues[2];
            ws.Range["B13"].Formula = "=XBCache_Set(A13)";
            ws.Range["A14"].Value = cachedValues[3];
            ws.Range["B14"].Formula = "=XBCache_Set(A14)";

            ws.Range["C11:C14"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfMissing_Optional(B11,B12,B13,B14))";
            Assert.Equal(ws.Range["A11"].Value, ws.Range["C11"].Value);
            Assert.Equal(ws.Range["A12"].Value, ws.Range["C12"].Value);
            Assert.Equal(ws.Range["A13"].Value, ws.Range["C13"].Value);
            Assert.Equal(ws.Range["A14"].Value2, ws.Range["C14"].Value2);
        }

        /// <summary>
        /// Test DefaultIfIncorrectType=true for cached params arg for CachedParameterConversions
        /// </summary>
        [ExcelFact]
        public void CachedParameterConversion_CachedParams_DefaultIfIncorrectType()
        {
            var ws = _testWorkbook.Sheets.Add();

            var stringValue = "string_value";
            var doubleValue = 123d;
            ws.Range["A2"].Value = stringValue;
            ws.Range["A3"].Formula = $"={doubleValue}*1";
            ws.Range["B2"].Formula = "=XBCache_Set(A2)";
            ws.Range["B3"].Formula = "=XBCache_Set(A3)";

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(FALSE)";
            ws.Range["C2:C3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_Object(B2,B3))";
            Assert.Equal(stringValue, ws.Range["C2"].Value);
            Assert.Equal(doubleValue, ws.Range["C3"].Value);
            ws.Range["D2:D3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_String(B2,B3))";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["D2"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["D3"].Value);
            ws.Range["E2:E3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_Default_Double(B2,B3))";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["E2"].Value);
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["E3"].Value);

            ws.Range["C1"].Formula = "=XBUtils_SetErrorOutput(TRUE)";
            ws.Range["F2:F3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_Object(B2,B3))";
            Assert.Equal(stringValue, ws.Range["F2"].Value);
            Assert.Equal(doubleValue, ws.Range["F3"].Value);
            ws.Range["G2:G3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_String(B2,B3))";
            Assert.Equal(stringValue, ws.Range["G2"].Value);
            Assert.Equal(stringValue, ws.Range["G3"].Value);
            ws.Range["H2:H3"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_Double(B2,B3))";
            Assert.Equal(0d, ws.Range["H2"].Value);
            Assert.Equal(doubleValue, ws.Range["H3"].Value);

            var cachedValues = new object[] { "string_value", 123d, "another_string", new DateTime(2023, 1, 23) };
            ws.Range["A4"].Value = cachedValues[0];
            ws.Range["B4"].Formula = "=XBCache_Set(A4)";
            ws.Range["A5"].Value = cachedValues[1];
            ws.Range["B5"].Formula = "=XBCache_Set(A5)";
            ws.Range["A6"].Value = cachedValues[2];
            ws.Range["B6"].Formula = "=XBCache_Set(A6)";
            ws.Range["A7"].Value = cachedValues[3];
            ws.Range["B7"].Formula = "=XBCache_Set(A7)";

            ws.Range["C4:C7"].FormulaArray = "=TRANSPOSE(XBTests_CachedParameterConversion_CachedParams_DefaultIfIncorrectType_Object(B4,B5,B6,B7))";
            Assert.Equal(ws.Range["A4"].Value, ws.Range["C4"].Value);
            Assert.Equal(ws.Range["A5"].Value, ws.Range["C5"].Value);
            Assert.Equal(ws.Range["A6"].Value, ws.Range["C6"].Value);
            Assert.Equal(ws.Range["A7"].Value2, ws.Range["C7"].Value2);
        }
    }
}
