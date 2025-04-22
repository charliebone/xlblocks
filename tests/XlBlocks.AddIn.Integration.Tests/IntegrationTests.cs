namespace XlBlocks.AddIn.Integration.Tests;

using System;
using ExcelDna.Integration;
using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;

public partial class IntegrationTests
{
    private static object AsExcelError(object value)
    {
        if (value is int comError)
        {
            switch (comError)
            {
                case -2146826288: return ExcelError.ExcelErrorNull;
                case -2146826281: return ExcelError.ExcelErrorDiv0;
                case -2146826265: return ExcelError.ExcelErrorRef;
                case -2146826259: return ExcelError.ExcelErrorName;
                case -2146826252: return ExcelError.ExcelErrorNum;
                case -2146826246: return ExcelError.ExcelErrorNA;
                case -2146826273: return ExcelError.ExcelErrorValue;
            }
        }
        return value;
    }

    public static void AssertIsExcelError(ExcelError expectedError, object value)
    {
        Assert.Equal<object>(expectedError, AsExcelError(value));
    }

    [ExcelTestSettings(AddIn = @"..\..\XlBlocks.Addin\debug\XlBlocks-AddIn")]
    public class IntegrationFunctionTests : IDisposable
    {
        readonly protected Workbook _testWorkbook;

        public IntegrationFunctionTests()
        {
            var app = Util.Application;
            _testWorkbook = app.Workbooks.Add();
        }

        public void Dispose()
        {
            _testWorkbook.Close(SaveChanges: false);
            GC.SuppressFinalize(this);
        }
    }

    [ExcelTestSettings(AddIn = @"..\..\XlBlocks.AddIn\debug\XlBlocks-AddIn", Workbook = "XlBlocksTestBook.xlsb")]
    public class IntegrationWorkbookTests : IDisposable
    {
        readonly Workbook _testWorkbook;

        public IntegrationWorkbookTests()
        {
            var app = Util.Application;
            _testWorkbook = app.Workbooks.Add();
        }

        public void Dispose()
        {
            _testWorkbook.Close(SaveChanges: false);
            GC.SuppressFinalize(this);
        }
    }
}
