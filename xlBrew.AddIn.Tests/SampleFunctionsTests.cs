namespace xlBrew.AddIn.Tests;

using System;
using Xunit;
using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;

public class SampleFunctionsTests
{
    // The path give here is relative to the output directory of the test project.
    // Setting an AddIn options here will request the test runner to load this add-in into Excel before the tests start.
    // The name here excludes the ".xll" or "64.xll" suffix. The test runner will choose according to the Excel bitness where it runs.
    [ExcelTestSettings(AddIn = @"..\..\..\..\xlBrew.Addin\bin\Debug\net6.0-windows\xlBrew-AddIn")]
    public class ExcelTests : IDisposable
    {
        // This workbook will be available to all tests in the class
        Workbook _testWorkbook;

        // The test class constructor will configure the required environment for the tests in the class.
        // In this case it creates a new Workbook that will be shared by the tests
        public ExcelTests()
        {
            var app = Util.Application;
            _testWorkbook = app.Workbooks.Add();
        }

        // Clean-up for the class is in the IDisposable.Dispose implementation
        public void Dispose()
        {
            _testWorkbook.Close(SaveChanges: false);
        }

        // This test just interacts with Excel
        [ExcelFact]
        public void ExcelCanAddNumbers()
        {
            var ws = _testWorkbook.Sheets[1];

            ws.Range["A1"].Value = 2.0;
            ws.Range["A2"].Value = 3.0;
            ws.Range["A3"].Formula = "=A1 + A2";

            var result = ws.Range["A3"].Value;

            Assert.Equal(5.0, result);
        }

        // This test depends on the AddIn value set in the class's ExcelTestSettings attributes
        // With the Sample-AddIn loaded, the function should work correctly.
        [ExcelFact]
        public void AddInCanAddNumbers()
        {
            var ws = _testWorkbook.Sheets[1];

            ws.Range["A1"].Value = 2.0;
            ws.Range["A2"].Value = 3.0;
            ws.Range["A3"].Formula = "=AddThem(A1, A2)";

            var result = ws.Range["A3"].Value;

            Assert.Equal(5.0, result);
        }

        // This test depends on the AddIn value set in the class's ExcelTestSettings attributes
        // With the Sample-AddIn loaded, the function should work correctly.
        [ExcelFact]
        public void AddInSaysHello()
        {
            var ws = _testWorkbook.Sheets[1];

            ws.Range["A1"].Value = "Charlie";
            ws.Range["A2"].Formula = "=SayHello(A1)";

            var result = ws.Range["A2"].Value;

            Assert.Equal("Hello Charlie", result);
        }

        // Before this test is run, a pre-created workbook will be loaded
        // It has been added to the test project and configured to always be copied to the output directory
        [ExcelFact(Workbook = "xlBrewTestBook.xlsx")]
        public void WorkbookCheckIsOK()
        {
            // Get the pre-loaded workbook using the Util.Workbook property
            var wb = Util.Workbook;
            var ws = wb.Sheets["Check"];
            Util.Application.CalculateFull();

            var result = ws.Range["A1"].Value;

            Assert.Equal("OK", result);
        }
    }
}