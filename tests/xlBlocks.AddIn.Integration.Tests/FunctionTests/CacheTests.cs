namespace XlBlocks.AddIn.Integration.Tests;

using ExcelDna.Integration;

public partial class IntegrationTests
{
    public partial class CacheTests : IntegrationFunctionTests
    {
        /// <summary>
        /// Test basic cache operations
        /// </summary>
        [ExcelFact]
        public void CacheGetSetWorks()
        {
            var ws = _testWorkbook.Sheets.Add();

            var testStr = "Cache test string";
            ws.Range["A1"].Value = testStr;
            ws.Range["A2"].Formula = "=xbCache_Set(A1)";
            ws.Range["A3"].Formula = "=xbCache_Get(A2)";

            Assert.Equal(testStr, ws.Range["A3"].Value);

            ws.Range["A4"].Formula = "=xbCache_Exists(A2)";
            Assert.Equal(true, ws.Range["A4"].Value);

            ws.Range["A5"].Formula = "=xbCache_Exists(A1)";
            Assert.Equal(false, ws.Range["A5"].Value);
        }

        /// <summary>
        /// Test cache miss error behavior
        /// </summary>
        [ExcelFact]
        public void CacheMissFlagsError()
        {
            var ws = _testWorkbook.Sheets.Add();

            var badHandle = "bad_cache_handle";
            ws.Range["A1"].Formula = badHandle;

            ws.Range["C1"].Formula = "=xbUtils_setErrorOutput(FALSE)";
            ws.Range["A2"].Formula = "=xbCache_Get(A1)";
            AssertIsExcelError(ExcelError.ExcelErrorNA, ws.Range["A2"].Value);

            ws.Range["C1"].Formula = "=xbUtils_setErrorOutput(TRUE)";
            ws.Range["A3"].Formula = "=xbCache_Get(A1)";
            Assert.Equal($"The input to 'obj' with handle '{badHandle}' was not found in the xlBlocks cache", ws.Range["A3"].Value);
        }
    }
}
