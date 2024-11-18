namespace XlBlocks.AddIn.Tests.Types;

using System;
using ExcelDna.Integration;
using XlBlocks.AddIn.Types;

public class XlBlockTableTests
{
    private static readonly XlBlockTable _employeeTable = XlBlockTable.Build(XlBlockRange.Build(
        new object[,]
        {
            { "ID", "Name", "Age" },
            { 1, "Alice", 30 },
            { 2, "Bob", 25 },
            { 3, "Charlie", 35 }
        }));

    private static readonly XlBlockTable _departmentTable = XlBlockTable.Build(XlBlockRange.Build(
        new object[,]
        {
            { "ID", "Department", "Age" },
            { 2, "HR", 3 },
            { 3, "Engineering", 4 },
            { 4, "Marketing", 5 }
        }));

    private static readonly XlBlockTable _assignmentTable = XlBlockTable.Build(XlBlockRange.Build(
        new object[,]
        {
            { "DeptID", "StartDate", "EmployeeID" },
            { 2, new DateTime(2023, 02, 01), 2 },
            { 3, new DateTime(2020, 12, 01), 1 },
            { 2, new DateTime(2022, 07, 10), 5 }
        }));

    private static readonly XlBlockTable _logDataTable = XlBlockTable.BuildWithTypes(XlBlockRange.Build(
        new object[,]
        {
            { 0, "Trace", ExcelError.ExcelErrorNA, 38.83 },
            { 1, "Warning", 25, ExcelError.ExcelErrorNA },
            { 2, ExcelError.ExcelErrorNA, 21, 83.45 },
            { 3, "Critical", 2, 1.77 },
            { 4, "Debug", 62, 53.67 },
            { 5, ExcelError.ExcelErrorNA, 22, ExcelError.ExcelErrorNA },
            { 6, "Warning", 11, 33.32 },
            { 7, "Critical", 45, 0.82 },
            { 8, "Info", 0, ExcelError.ExcelErrorNA },
            { 9, "Info", 101, ExcelError.ExcelErrorNA },
            { 10, "Debug", 62, 6.34 },
            { 11, "Warning", 45, ExcelError.ExcelErrorNA }
        }),
        XlBlockRange.Build(new[] { "int", "string", "int", "double" }),
        XlBlockRange.Build(new[] { "Id", "Category", "ErrorCount", "Average" }));

    private static void AssertTableMatchesExpected(object[,] expected, XlBlockTable? actual)
    {
        Assert.NotNull(actual);
        Assert.Equal(expected.GetLength(0) - 1, actual.RowCount);
        Assert.Equal(expected.GetLength(1), actual.ColumnCount);

        for (var col = 0; col < actual.ColumnCount; col++)
        {
            Assert.Equal(expected[0, col], actual.ColumnNames[col]);

            for (var row = 0; row < actual.RowCount; row++)
            {
                if (expected[row + 1, col] is double dblExpected && actual[row, col] is double dblActual)
                    Assert.Equal(dblExpected, dblActual, 0.0001);
                else
                    Assert.Equal(expected[row + 1, col], actual[row, col]);
            }
        }
    }

    #region Build tests

    [Fact]
    public void Build_ValidDataRange_CreatesXlBlockTable()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var table = XlBlockTable.Build(dataRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "Column1", "Column2" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(double), typeof(double) }, table.ColumnTypes);
    }

    [Fact]
    public void Build_EmptyHeaderRow_ThrowsArgumentException()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { ExcelEmpty.Value, ExcelEmpty.Value },
            { 1, 2 },
            { 3, 4 }
        });

        Assert.Throws<ArgumentException>(() => XlBlockTable.Build(dataRange));
    }

    [Fact]
    public void Build_MissingHeaderRow_ThrowsArgumentException()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { ExcelMissing.Value, ExcelMissing.Value },
            { 1, 2 },
            { 3, 4 }
        });

        Assert.Throws<ArgumentException>(() => XlBlockTable.Build(dataRange));
    }

    [Fact]
    public void Build_ErrorHeaderRow_ThrowsArgumentException()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { ExcelError.ExcelErrorNA, ExcelError.ExcelErrorNA },
            { 1, 2 },
            { 3, 4 }
        });

        Assert.Throws<ArgumentException>(() => XlBlockTable.Build(dataRange));
    }

    [Fact]
    public void Build_DataWithErrorsAndMissingValues_SuccessfullyConvertsToDBNull()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, ExcelError.ExcelErrorNA },
            { ExcelMissing.Value, 4 }
        });

        var table = XlBlockTable.Build(dataRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "Column1", "Column2" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(double), typeof(double) }, table.ColumnTypes);

        Assert.Null(table[0, "Column2"]);
        Assert.Null(table[1, "Column1"]);
    }

    [Fact]
    public void BuildWithTypes_ValidDataRange_CreatesXlBlockTableWithSpecifiedTypes()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, 2 },
            { 3, 4 }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "int" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" }
        });

        var table = XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "Column1", "Column2" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(int), typeof(int) }, table.ColumnTypes);

        Assert.Equal(1, table[0, 0]);
        Assert.Equal(2, table[0, 1]);
        Assert.Equal(3, table[1, 0]);
        Assert.Equal(4, table[1, 1]);
    }

    [Fact]
    public void BuildWithTypes_InvalidType_ThrowsArgumentException()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, "invalid" },
            { 3, 4 }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "int" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" }
        });

        Assert.Throws<ArgumentException>(() => XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange));
    }

    [Fact]
    public void BuildWithTypes_DataWithErrorsAndMissingValues_SuccessfullyConvertsToDBNull()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, ExcelError.ExcelErrorNA },
            { ExcelMissing.Value, 4 }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "int" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" }
        });

        var table = XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "Column1", "Column2" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(int), typeof(int) }, table.ColumnTypes);

        Assert.Null(table[0, "Column2"]);
        Assert.Null(table[1, "Column1"]);
    }

    [Fact]
    public void BuildWithTypes_MixedTypes_CreatesXlBlockTable()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, "test", 3.14, DateTime.Now },
            { 2, "another", 2.718, DateTime.Now.AddDays(1) }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "string", "double", "datetime" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "IntColumn", "StringColumn", "DoubleColumn", "DateTimeColumn" }
        });

        var table = XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
        Assert.Equal(new string[] { "IntColumn", "StringColumn", "DoubleColumn", "DateTimeColumn" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(int), typeof(string), typeof(double), typeof(DateTime) }, table.ColumnTypes);

        Assert.Equal(1, table[0, 0]);
        Assert.Equal("test", table[0, 1]);
        Assert.Equal(3.14, table[0, 2]);
        Assert.True(table[0, 3] is DateTime);
    }

    [Fact]
    public void BuildWithTypes_TypeRangeErrors_ThrowsArgumentException()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, "test", 3.14 },
            { 2, "another", 2.718 }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", ExcelError.ExcelErrorNA, "double" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "IntColumn", "StringColumn", "DoubleColumn" }
        });

        Assert.Throws<ArgumentException>(() => XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange));
    }

    [Fact]
    public void BuildWithTypes_DataWithMixedTypesAndErrors_SuccessfullyConvertsToExpectedTypes()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, "test", 3.14 },
        { ExcelMissing.Value, "another", ExcelError.ExcelErrorNA }
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "string", "double" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "IntColumn", "StringColumn", "DoubleColumn" }
        });

        var table = XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange);

        Assert.Equal(2, table.RowCount);
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(new string[] { "IntColumn", "StringColumn", "DoubleColumn" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(int), typeof(string), typeof(double) }, table.ColumnTypes);

        Assert.Null(table[1, "IntColumn"]);
        Assert.Null(table[1, "DoubleColumn"]);
    }

    [Fact]
    public void BuildFromDictionary_StringKey_DoubleValues()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Item1"] = 123,
            ["key2"] = 45.67,
            ["third key"] = 90.22
        }, typeof(string));

        var table = XlBlockTable.BuildFromDictionary(dictionary, "The Key", "Values");

        Assert.Equal(3, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "The Key", "Values" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(string), typeof(double) }, table.ColumnTypes);
    }

    [Fact]
    public void BuildFromDictionary_StringKey_StringValues()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Item1"] = 123,
            ["key2"] = "asdf",
            ["third key"] = 90.22,
            ["date value"] = new DateTime(2022, 1, 1)
        }, typeof(string));

        var table = XlBlockTable.BuildFromDictionary(dictionary, "String Key", "String Value");

        Assert.Equal(4, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "String Key", "String Value" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(string), typeof(string) }, table.ColumnTypes);
    }

    [Fact]
    public void BuildFromDictionary_IntKey_ExplicitValueType()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<int, object>
        {
            [23] = new DateTime(2022, 1, 1),
            [45] = 45678,
            [1] = "2024-01-23"
        }, typeof(int));

        var table = XlBlockTable.BuildFromDictionary(dictionary, "int", "v", "date");

        Assert.Equal(3, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new string[] { "int", "v" }, table.ColumnNames);
        Assert.Equal(new Type[] { typeof(int), typeof(DateTime) }, table.ColumnTypes);
    }

    [Fact]
    public void BuildFromCsv_UnspecifiedTypes()
    {
        var csvPath = Path.Combine(Directory.GetCurrentDirectory(), "Types", "logData.csv");
        var table = XlBlockTable.BuildFromCsv(csvPath, ",", true);
        Assert.Equal(12, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
    }

    [Fact]
    public void BuildFromCsv_SpecifiedTypes()
    {
        var csvPath = Path.Combine(Directory.GetCurrentDirectory(), "Types", "logData.csv");
        var table = XlBlockTable.BuildFromCsv(csvPath, ",", true);
        Assert.Equal(12, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
    }

    #endregion

    #region Join and Union

    [Fact]
    public void Join_InnerJoin_CommonColumns()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "inner", null);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Department" }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_InnerJoin_CommonColumns_IncludeDuplicates()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "inner", null, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_InnerJoin_SpecifiedColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "inner", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age.left", "Department", "Age.right" },
            { 2d, "Bob", 25d, "HR", 3d },
            { 3d, "Charlie", 35d, "Engineering", 4d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_InnerJoin_SpecifiedColumns_IncludeDuplicates()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "inner", joinOn, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 2d, "Bob", 25d, 2d, "HR", 3d },
            { 3d, "Charlie", 35d, 3d, "Engineering", 4d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_InnerJoin_SpecifiedButDifferentColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID", "EmployeeID" } });
        var result = XlBlockTable.Join(_employeeTable, _assignmentTable, "inner", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "DeptID", "StartDate", "EmployeeID" },
            { 2d, "Bob", 25d, 2d, new DateTime(2023, 02, 01), 2d },
            { 1d, "Alice", 30d, 3d, new DateTime(2020, 12, 01), 1d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_LeftJoin_CommonColumns()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "left", null);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Department", },
            { 1d, "Alice", 30d, null! },
            { 2d, "Bob", 25d, null! },
            { 3d, "Charlie", 35d, null! }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_LeftJoin_CommonColumns_IncludeDuplicates()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "left", null, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 1d, "Alice", 30d, null!, null!, null! },
            { 2d, "Bob", 25d, null!, null!, null! },
            { 3d, "Charlie", 35d, null!, null!, null! }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_LeftJoin_SpecifiedColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "left", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age.left", "Department", "Age.right" },
            { 1d, "Alice", 30d, null!, null! },
            { 2d, "Bob", 25d, "HR", 3d },
            { 3d, "Charlie", 35d, "Engineering", 4d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_LeftJoin_SpecifiedColumns_IncludeDuplicates()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "left", joinOn, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 1d, "Alice", 30d, null!, null!, null! },
            { 2d, "Bob", 25d, 2d, "HR", 3d },
            { 3d, "Charlie", 35d, 3d, "Engineering", 4d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_LeftJoin_SpecifiedButDifferentColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID", "EmployeeID" } });
        var result = XlBlockTable.Join(_employeeTable, _assignmentTable, "left", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "DeptID", "StartDate", "EmployeeID" },
            { 1d, "Alice", 30d, 3d, new DateTime(2020, 12, 01), 1d },
            { 2d, "Bob", 25d, 2d, new DateTime(2023, 02, 01), 2d },
            { 3d, "Charlie", 35d, null!, null!, null! },
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_RightJoin_CommonColumns()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "right", null);

        object[,] expectedResult =
        {
            { "Name", "ID", "Department", "Age" },
            { null!, 2d, "HR", 3d },
            { null!, 3d, "Engineering", 4d },
            { null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_RightJoin_CommonColumns_IncludeDuplicates()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "right", null, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { null!, null!, null!, 2d, "HR", 3d },
            { null!, null!, null!, 3d, "Engineering", 4d },
            { null!, null!, null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_RightJoin_SpecifiedColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "right", joinOn);

        object[,] expectedResult =
        {
            { "Name", "Age.left", "ID", "Department", "Age.right" },
            { "Bob", 25d, 2d, "HR", 3d },
            { "Charlie", 35d, 3d, "Engineering", 4d },
            { null!, null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_RightJoin_SpecifiedColumns_IncludeDuplicates()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "right", joinOn, includeDuplicateJoinColumns: true);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 2d, "Bob", 25d, 2d, "HR", 3d },
            { 3d, "Charlie", 35d, 3d, "Engineering", 4d },
            { null!, null!, null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_RightJoin_SpecifiedButDifferentColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID", "EmployeeID" } });
        var result = XlBlockTable.Join(_employeeTable, _assignmentTable, "right", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "DeptID", "StartDate", "EmployeeID" },
            { 2d, "Bob", 25d, 2d, new DateTime(2023, 02, 01), 2d },
            { 1d, "Alice", 30d, 3d, new DateTime(2020, 12, 01), 1d },
            { null!, null!, null!, 2d, new DateTime(2022, 07, 10), 5d },
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_FullOuterJoin_CommonColumns()
    {
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "full", null);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 1d, "Alice", 30d, null!, null!, null! },
            { 2d, "Bob", 25d, null!, null!, null! },
            { 3d, "Charlie", 35d, null!, null!, null! },
            { null!, null!, null!, 2d, "HR", 3d },
            { null!, null!, null!, 3d, "Engineering", 4d },
            { null!, null!, null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_FullOuterJoin_SpecifiedColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID" } });
        var result = XlBlockTable.Join(_employeeTable, _departmentTable, "full", joinOn);

        object[,] expectedResult =
        {
            { "ID.left", "Name", "Age.left", "ID.right", "Department", "Age.right" },
            { 1d, "Alice", 30d, null!, null!, null! },
            { 2d, "Bob", 25d, 2d, "HR", 3d },
            { 3d, "Charlie", 35d, 3d, "Engineering", 4d },
            { null!, null!, null!, 4d, "Marketing", 5d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_FullOuterJoin_SpecifiedButDifferentColumns()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID", "EmployeeID" } });
        var result = XlBlockTable.Join(_employeeTable, _assignmentTable, "full", joinOn);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "DeptID", "StartDate", "EmployeeID" },
            { 1d, "Alice", 30d, 3d, new DateTime(2020, 12, 01), 1d },
            { 2d, "Bob", 25d, 2d, new DateTime(2023, 02, 01), 2d },
            { 3d, "Charlie", 35d, null!, null!, null! },
            { null!, null!, null!, 2d, new DateTime(2022, 07, 10), 5d },
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Join_SpecifiedMissingColumns_Error()
    {
        var joinOn = XlBlockRange.Build(new object[,] { { "ID", "NotAnID" } });
        Assert.Throws<ArgumentException>(() => XlBlockTable.Join(_employeeTable, _assignmentTable, "full", joinOn));
    }

    [Fact]
    public void Join_NoCommonColumns_Error()
    {
        Assert.Throws<ArgumentException>(() => XlBlockTable.Join(_departmentTable, _assignmentTable, "full", null));
    }

    [Fact]
    public void UnionAll_MatchingTables()
    {
        var result = XlBlockTable.UnionAll(_employeeTable, _employeeTable, _employeeTable);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void UnionAll_DifferingMatchingTables()
    {
        Assert.Throws<ArgumentException>(() => XlBlockTable.UnionAll(_employeeTable, _departmentTable));
    }

    [Fact]
    public void Union_MatchingTables_DuplicatesRemoved()
    {
        var moreEmployees = XlBlockTable.Build(XlBlockRange.Build(
        new object[,]
        {
            { "ID", "Name", "Age" },
            { 1, "Alice", 30 },
            { 4, "Drew", 22 },
            { 3, "Charlie", 35 }
        }));

        var result = XlBlockTable.Union(_employeeTable, moreEmployees, _employeeTable);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d },
            { 4d, "Drew", 22d }
        };

        AssertTableMatchesExpected(expectedResult, result);
    }

    #endregion

    #region Sorting

    [Fact]
    public void Sort_SingleColumnWithNulls()
    {
        var expected = new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 3, "Critical", 2, 1.77 },
            { 7, "Critical", 45, 0.82 },
            { 4, "Debug", 62, 53.67 },
            { 10, "Debug", 62, 6.34 },
            { 8, "Info", 0, null! },
            { 9, "Info", 101, null! },
            { 0, "Trace", null!, 38.83 },
            { 1, "Warning", 25, null! },
            { 6, "Warning", 11, 33.32 },
            { 11, "Warning", 45, null! },
            { 2, null!, 21, 83.45 },
            { 5, null!, 22, null! },
        };

        var sorted = _logDataTable.Sort(
            XlBlockRange.Build(new object[] { "Category" }),
            XlBlockRange.Build(new object[] { false }),
            XlBlockRange.Build(new object[] { false }));
        AssertTableMatchesExpected(expected, sorted);
    }

    [Fact]
    public void Sort_MultiColumnWithNulls()
    {
        var expected = new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 7, "Critical", 45, 0.82 },
            { 3, "Critical", 2, 1.77 },
            { 10, "Debug", 62, 6.34 },
            { 4, "Debug", 62, 53.67 },
            { 8, "Info", 0, null! },
            { 9, "Info", 101, null! },
            { 0, "Trace", null!, 38.83 },
            { 6, "Warning", 11, 33.32 },
            { 1, "Warning", 25, null! },
            { 11, "Warning", 45, null! },
            { 2, null!, 21, 83.45 },
            { 5, null!, 22, null! },
        };

        var sorted = _logDataTable.Sort(
            XlBlockRange.Build(new object[] { "Category", "Average" }),
            XlBlockRange.Build(new object[] { false, false }),
            XlBlockRange.Build(new object[] { false, false }));
        AssertTableMatchesExpected(expected, sorted);
    }

    [Fact]
    public void Sort_MultiColumnWithNulls_Descending()
    {
        var expected = new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 8, "Info", 0, null! },
            { 3, "Critical", 2, 1.77 },
            { 6, "Warning", 11, 33.32 },
            { 2, null!, 21, 83.45 },
            { 5, null!, 22, null! },
            { 1, "Warning", 25, null! },
            { 11, "Warning", 45, null! },
            { 7, "Critical", 45, 0.82 },
            { 4, "Debug", 62, 53.67 },
            { 10, "Debug", 62, 6.34 },
            { 9, "Info", 101, null! },
            { 0, "Trace", null!, 38.83 },
        };

        var sorted = _logDataTable.Sort(
            XlBlockRange.Build(new object[] { "ErrorCount", "Category", "Average" }),
            XlBlockRange.Build(new object[] { false, true, true }),
            XlBlockRange.Build(new object[] { false, false, false }));
        AssertTableMatchesExpected(expected, sorted);
    }

    [Fact]
    public void Sort_MultiColumnWithNulls_Descending_NullsFirst_Irrelevant()
    {
        var expected = new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 9, "Info", 101, null! },
            { 10, "Debug", 62, 6.34 },
            { 4, "Debug", 62, 53.67 },
            { 11, "Warning", 45, null! },
            { 7, "Critical", 45, 0.82 },
            { 1, "Warning", 25, null! },
            { 5, null!, 22, null! },
            { 2, null!, 21, 83.45 },
            { 6, "Warning", 11, 33.32 },
            { 3, "Critical", 2, 1.77 },
            { 8, "Info", 0, null! },
            { 0, "Trace", null!, 38.83 },
        };

        var sorted = _logDataTable.Sort(
            XlBlockRange.Build(new object[] { "ErrorCount", "Category", "Average" }),
            XlBlockRange.Build(new object[] { true, true, false }),
            XlBlockRange.Build(new object[] { false, true, true }));
        AssertTableMatchesExpected(expected, sorted);
    }

    [Fact]
    public void Sort_MultiColumnWithNulls_Descending_NullsFirst_Matters()
    {
        var expected = new object[,]
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 5, null!, 22, null! },
            { 2, null!, 21, 83.45 },
            { 11, "Warning", 45, null! },
            { 1, "Warning", 25, null! },
            { 6, "Warning", 11, 33.32 },
            { 0, "Trace", null!, 38.83 },
            { 9, "Info", 101, null! },
            { 8, "Info", 0, null! },
            { 10, "Debug", 62, 6.34 },
            { 4, "Debug", 62, 53.67 },
            { 7, "Critical", 45, 0.82 },
            { 3, "Critical", 2, 1.77 },
        };

        var sorted = _logDataTable.Sort(
            XlBlockRange.Build(new object[] { "Category", "ErrorCount", "Average" }),
            XlBlockRange.Build(new object[] { true, true, false }),
            XlBlockRange.Build(new object[] { true, false, false }));
        AssertTableMatchesExpected(expected, sorted);
    }

    #endregion

    #region Filtering

    [Fact]
    public void Filter_InclusiveFilter()
    {
        var result = _employeeTable.Filter("Age", 30, inclusive: true);
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_ExclusiveFilter()
    {
        var result = _employeeTable.Filter("Age", 30, inclusive: false);
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_Empty()
    {
        var result = _employeeTable.Filter("");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_GreaterThanOrEqual()
    {
        var result = _employeeTable.Filter("[Age] >= 30");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 3d, "Charlie", 35d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_Equality()
    {
        var result = _employeeTable.Filter("[Name] == 'Alice'");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_Inequality()
    {
        var result = _employeeTable.Filter("[Name] <> 'Alice'");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_IsNull()
    {
        var result = _logDataTable.Filter("[Category] is null");
        object[,] expectedResult =
        {
            { "Id", "Category", "ErrorCount", "Average" },
            { 2, null!, 21, 83.45 },
            { 5, null!, 22, null! }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_IsNotNull()
    {
        var result = _employeeTable.Filter("[Name] is not null");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d },
            { 3d, "Charlie", 35d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_InClause()
    {
        var result = _employeeTable.Filter("[Name] IN ('Alice','Bob')");
        object[,] expectedResult =
        {
            { "ID", "Name", "Age" },
            { 1d, "Alice", 30d },
            { 2d, "Bob", 25d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Filter_Expression_IIF()
    {
        var result = _logDataTable.Filter("IIF([Category] IN ('Debug', 'Info'), [Average] > 5, [Average] > 10)");
        object[,] expectedResult =
        {
        { "Id", "Category", "ErrorCount", "Average" },
            { 0, "Trace", null!, 38.83 },
            { 2, null!, 21, 83.45 },
            { 4, "Debug", 62, 53.67 },
            { 6, "Warning", 11, 33.32 },
            { 10, "Debug", 62, 6.34 }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    #endregion

    #region Appending

    [Fact]
    public void AppendColumnsWith_AddAgeCategory()
    {
        var columnNames = XlBlockRange.Build(new object[,] { { "AgeCategory" } });
        var columnExpressions = XlBlockRange.Build(new object[,] { { "IIF([Age] >= 30, 'Senior', 'Junior')" } });

        var result = _employeeTable.AppendColumnsWith(columnNames, columnExpressions);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "AgeCategory" },
            { 1d, "Alice", 30d, "Senior" },
            { 2d, "Bob", 25d, "Junior" },
            { 3d, "Charlie", 35d, "Senior" }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnsWith_ISNULL()
    {
        var columnNames = XlBlockRange.Build(new object[,] { { "CategoryOrDefault" } });
        var columnExpressions = XlBlockRange.Build(new object[,] { { "ISNULL([Category], 'Unknown')" } });

        var result = _logDataTable.AppendColumnsWith(columnNames, columnExpressions);

        object[,] expectedResult =
        {
            { "Id", "Category", "ErrorCount", "Average", "CategoryOrDefault" },
            { 0, "Trace", null!, 38.83, "Trace" },
            { 1, "Warning", 25, null!, "Warning" },
            { 2, null!, 21, 83.45, "Unknown" },
            { 3, "Critical", 2, 1.77, "Critical" },
            { 4, "Debug", 62, 53.67, "Debug" },
            { 5, null!, 22, null!, "Unknown" },
            { 6, "Warning", 11, 33.32, "Warning" },
            { 7, "Critical", 45, 0.82, "Critical" },
            { 8, "Info", 0, null!, "Info" },
            { 9, "Info", 101, null!, "Info" },
            { 10, "Debug", 62, 6.34, "Debug" },
            { 11, "Warning", 45, null!, "Warning" }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnsWith_MultipleColumnsWithDependencies()
    {
        var columnNames = XlBlockRange.Build(new object[,]
        {
            { "CategoryOrDefault" },
            { "IsCritical" },
            { "CategoryDescription" }
        });
        var columnExpressions = XlBlockRange.Build(new object[,]
        {
            { "ISNULL([Category], 'Unknown')" },
            { "[CategoryOrDefault] == 'Critical'" },
            { "IIF([IsCritical], 'Urgent', 'Regular')" }
        });

        var result = _logDataTable.AppendColumnsWith(columnNames, columnExpressions);

        object[,] expectedResult =
        {
            { "Id", "Category", "ErrorCount", "Average", "CategoryOrDefault", "IsCritical", "CategoryDescription" },
            { 0, "Trace", null!, 38.83, "Trace", false, "Regular" },
            { 1, "Warning", 25, null!, "Warning", false, "Regular" },
            { 2, null!, 21, 83.45, "Unknown", false, "Regular" },
            { 3, "Critical", 2, 1.77, "Critical", true, "Urgent" },
            { 4, "Debug", 62, 53.67, "Debug", false, "Regular" },
            { 5, null!, 22, null!, "Unknown", false, "Regular" },
            { 6, "Warning", 11, 33.32, "Warning", false, "Regular" },
            { 7, "Critical", 45, 0.82, "Critical", true, "Urgent" },
            { 8, "Info", 0, null!, "Info", false, "Regular" },
            { 9, "Info", 101, null!, "Info", false, "Regular" },
            { 10, "Debug", 62, 6.34, "Debug", false, "Regular" },
            { 11, "Warning", 45, null!, "Warning", false, "Regular" }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromList_DefaultTypeDetection()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 10 }, { 20 }, { 30 } }), "drop");
        var result = _employeeTable.AppendColumnFromList(list, "BonusPoints");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "BonusPoints" },
            { 1d, "Alice", 30d, 10d },
            { 2d, "Bob", 25d, 20d },
            { 3d, "Charlie", 35d, 30d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromList_SpecifiedColumnType()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 100 }, { 200 }, { 300 } }), "drop");
        var result = _employeeTable.AppendColumnFromList(list, "Salaries", "double");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Salaries" },
            { 1d, "Alice", 30d, 100d },
            { 2d, "Bob", 25d, 200d },
            { 3d, "Charlie", 35d, 300d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromList_MismatchedLength_ThrowsException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 10 }, { 20 } }), "drop");

        Assert.Throws<ArgumentException>(() => _employeeTable.AppendColumnFromList(list, "BonusPoints"));
    }

    [Fact]
    public void AppendColumnFromList_DuplicateColumnName_ThrowsException()
    {
        var list = XlBlockList.Build(XlBlockRange.Build(new object[,] { { 10 }, { 20 }, { 30 } }), "drop");

        Assert.Throws<ArgumentException>(() => _employeeTable.AppendColumnFromList(list, "Age"));
    }

    [Fact]
    public void AppendColumnFromDictionary_DefaultTypeDetection()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = 1000,
            ["Bob"] = 2000,
            ["Charlie"] = 3000
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Salaries");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Salaries" },
            { 1d, "Alice", 30d, 1000d },
            { 2d, "Bob", 25d, 2000d },
            { 3d, "Charlie", 35d, 3000d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_SpecifiedColumnType()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = "Level 1",
            ["Bob"] = "Level 2",
            ["Charlie"] = "Level 3"
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Levels", "string");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Levels" },
            { 1d, "Alice", 30d, "Level 1" },
            { 2d, "Bob", 25d, "Level 2" },
            { 3d, "Charlie", 35d, "Level 3" }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_MissingKeys_BecomesNull()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = 1000,
            ["Charlie"] = 3000
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Salaries");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Salaries" },
            { 1d, "Alice", 30d, 1000d },
            { 2d, "Bob", 25d, null! },
            { 3d, "Charlie", 35d, 3000d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_MissingKeys_ValueOnMissing()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = 1000,
            ["Charlie"] = 3000
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Salaries", null, 1234d);

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "Salaries" },
            { 1d, "Alice", 30d, 1000d },
            { 2d, "Bob", 25d, 1234d },
            { 3d, "Charlie", 35d, 3000d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_DuplicateColumnName_ThrowsException()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = 1000,
            ["Bob"] = 2000,
            ["Charlie"] = 3000
        }, typeof(string));

        Assert.Throws<ArgumentException>(() => _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Age"));
    }

    [Fact]
    public void AppendColumnFromDictionary_CoerceDoubleToString()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = 10.5,
            ["Bob"] = 20.75,
            ["Charlie"] = 30.25
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "BonusPoints", "string");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "BonusPoints" },
            { 1d, "Alice", 30d, "10.5" },
            { 2d, "Bob", 25d, "20.75" },
            { 3d, "Charlie", 35d, "30.25" }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_CoerceStringToBoolean()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = "true",
            ["Bob"] = "false",
            ["Charlie"] = "true"
        }, typeof(string));

        var result = _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "IsActive", "bool");

        object[,] expectedResult =
        {
            { "ID", "Name", "Age", "IsActive" },
            { 1d, "Alice", 30d, true },
            { 2d, "Bob", 25d, false },
            { 3d, "Charlie", 35d, true }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void AppendColumnFromDictionary_CoerceNonNumericStringToDouble_ThrowsException()
    {
        var dictionary = new XlBlockDictionary(new Dictionary<string, object>
        {
            ["Alice"] = "ten",
            ["Bob"] = "twenty",
            ["Charlie"] = "thirty"
        }, typeof(string));

        Assert.Throws<ArgumentException>(() => _employeeTable.AppendColumnFromDictionary(dictionary, "Name", "Amounts", "double"));
    }

    #endregion

    #region Misc

    [Fact]
    public void Copy_ValidTable_CreatesIdenticalCopy()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var originalTable = XlBlockTable.Build(dataRange);
        var copiedTable = originalTable.Copy();

        Assert.Equal(originalTable.RowCount, copiedTable.RowCount);
        Assert.Equal(originalTable.ColumnCount, copiedTable.ColumnCount);
        Assert.Equal(originalTable.ColumnNames, copiedTable.ColumnNames);
        Assert.Equal(originalTable.ColumnTypes, copiedTable.ColumnTypes);

        for (int i = 0; i < originalTable.RowCount; i++)
        {
            for (int j = 0; j < originalTable.ColumnCount; j++)
            {
                Assert.Equal(originalTable[i, j], copiedTable[i, j]);
            }
        }
    }

    [Fact]
    public void Copy_TableWithMixedTypes_CreatesIdenticalCopy()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { 1, "test", 3.14, DateTime.Now },
            { 2, "another", 2.718, DateTime.Now.AddDays(1) },
            { 3, "third", 1.414, new DateTime(2023, 1, 23) },
        });

        var typeRange = XlBlockRange.Build(new object[,]
        {
            { "int", "string", "double", "datetime" }
        });

        var headerRange = XlBlockRange.Build(new object[,]
        {
            { "IntColumn", "StringColumn", "DoubleColumn", "DateTimeColumn" }
        });

        var originalTable = XlBlockTable.BuildWithTypes(dataRange, typeRange, headerRange);
        var copiedTable = originalTable.Copy();

        Assert.Equal(originalTable.RowCount, copiedTable.RowCount);
        Assert.Equal(originalTable.ColumnCount, copiedTable.ColumnCount);
        Assert.Equal(originalTable.ColumnNames, copiedTable.ColumnNames);
        Assert.Equal(originalTable.ColumnTypes, copiedTable.ColumnTypes);

        for (int i = 0; i < originalTable.RowCount; i++)
        {
            for (int j = 0; j < originalTable.ColumnCount; j++)
            {
                Assert.Equal(originalTable[i, j], copiedTable[i, j]);
            }
        }
    }

    [Fact]
    public void AsArray_IncludeHeaderByColumn_ReturnsArrayWithHeader()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var table = XlBlockTable.Build(dataRange);
        var result = table.AsArray(true, orientation: RangeOrientation.ByColumn);

        var expected = new object[,]
        {
            { "Column1", "Column2" },
            { 1d, 2d },
            { 3d, 4d }
        };

        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ExcludeHeaderByColumn_ReturnsArrayWithoutHeader()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var table = XlBlockTable.Build(dataRange);
        var result = table.AsArray(false, orientation: RangeOrientation.ByColumn);

        var expected = new object[,]
        {
            { 1d, 2d },
            { 3d, 4d }
        };

        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_IncludeHeaderByRow_ReturnsArrayWithHeader()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var table = XlBlockTable.Build(dataRange);
        var result = table.AsArray(true, orientation: RangeOrientation.ByRow);

        var expected = new object[,]
        {
            { "Column1", 1d, 3d },
            { "Column2", 2d, 4d }
        };

        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ExcludeHeaderByRow_ReturnsArrayWithoutHeader()
    {
        var dataRange = XlBlockRange.Build(new object[,]
        {
            { "Column1", "Column2" },
            { 1, 2 },
            { 3, 4 }
        });

        var table = XlBlockTable.Build(dataRange);
        var result = table.AsArray(false, orientation: RangeOrientation.ByRow);

        var expected = new object[,]
        {
            { 1d, 3d },
            { 2d, 4d }
        };

        Assert.Equal(expected, result);
    }

    [Fact]
    public void AsArray_ByColumn_ConvertNullToExcelErrorNA()
    {
        object[,] data =
        {
            { "ID", "Value" },
            { 1, null! },
            { 2, ExcelError.ExcelErrorNA },
            { 3, ExcelMissing.Value },
            { 4, ExcelEmpty.Value }
        };

        var table = XlBlockTable.Build(XlBlockRange.Build(data));
        var result = table.AsArray(true, orientation: RangeOrientation.ByColumn);

        object[,] expectedResult =
        {
            { "ID", "Value" },
            { 1d, ExcelError.ExcelErrorNA },
            { 2d, ExcelError.ExcelErrorNA },
            { 3d, ExcelError.ExcelErrorNA },
            { 4d, ExcelError.ExcelErrorNA }
        };

        Assert.Equal(expectedResult.GetLength(0), result.GetLength(0));
        Assert.Equal(expectedResult.GetLength(1), result.GetLength(1));

        for (int i = 0; i < result.GetLength(0); i++)
        {
            for (int j = 0; j < result.GetLength(1); j++)
            {
                Assert.Equal(expectedResult[i, j], result[i, j]);
            }
        }
    }

    [Fact]
    public void AsArray_ByRow_ConvertNullToExcelErrorNA()
    {
        object[,] data =
        {
            { "ID", "Value" },
            { 1, null! },
            { 2, ExcelError.ExcelErrorNA },
            { 3, ExcelMissing.Value },
            { 4, ExcelEmpty.Value }
        };

        var table = XlBlockTable.Build(XlBlockRange.Build(data));
        var result = table.AsArray(true, orientation: RangeOrientation.ByRow);

        object[,] expectedResult =
        {
            { "ID", 1d, 2d, 3d, 4d },
            { "Value", ExcelError.ExcelErrorNA, ExcelError.ExcelErrorNA, ExcelError.ExcelErrorNA, ExcelError.ExcelErrorNA }
        };

        Assert.Equal(expectedResult.GetLength(0), result.GetLength(0));
        Assert.Equal(expectedResult.GetLength(1), result.GetLength(1));

        for (int i = 0; i < result.GetLength(0); i++)
        {
            for (int j = 0; j < result.GetLength(1); j++)
            {
                Assert.Equal(expectedResult[i, j], result[i, j]);
            }
        }
    }

    [Fact]
    public void LookupValue_SingleMatch()
    {
        var result = _employeeTable.LookupValue("Name", "Alice", "ID");

        Assert.Equal(1d, result);
    }

    [Fact]
    public void LookupValue_MultipleMatches_Error()
    {
        var multipleMatchTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        Assert.Throws<ArgumentException>(() => multipleMatchTable.LookupValue("Name", "Alice", "ID"));
    }

    [Fact]
    public void LookupValue_MultipleMatches_First()
    {
        var multipleMatchTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        var result = multipleMatchTable.LookupValue("Name", "Alice", "ID", "first");

        Assert.Equal(1d, result);
    }

    [Fact]
    public void LookupValue_MultipleMatches_Last()
    {
        var multipleMatchTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        var result = multipleMatchTable.LookupValue("Name", "Alice", "ID", "last");

        Assert.Equal(3d, result);
    }

    [Fact]
    public void LookupValue_DifferentColumn()
    {
        var result = _employeeTable.LookupValue("Name", "Alice", "Age");

        Assert.Equal(30d, result);
    }

    [Fact]
    public void ToDictionary_ValidKeyValueColumns()
    {
        var result = _employeeTable.ToDictionary("Name", "Age");

        Assert.Equal(3, result.Count);
        Assert.Equal(typeof(string), result.KeyType);
        Assert.True(result.ContainsKey("Alice"));
        Assert.True(result.ContainsKey("Bob"));
        Assert.True(result.ContainsKey("Charlie"));
        Assert.Equal(30d, result["Alice"]);
        Assert.Equal(25d, result["Bob"]);
        Assert.Equal(35d, result["Charlie"]);
    }

    [Fact]
    public void ToDictionary_MissingKeyColumn_ThrowsException()
    {
        Assert.Throws<ArgumentException>(() => _employeeTable.ToDictionary("InvalidKeyColumn", "Age"));
    }

    [Fact]
    public void ToDictionary_MissingValueColumn_ThrowsException()
    {
        Assert.Throws<ArgumentException>(() => _employeeTable.ToDictionary("Name", "InvalidValueColumn"));
    }

    [Fact]
    public void ToDictionary_DuplicateKeys_ThrowsException()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        Assert.Throws<ArgumentException>(() => duplicateKeyTable.ToDictionary("Name", "Age"));
    }

    [Fact]
    public void ToDictionary_DuplicateKeys_TakesFirst()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        var result = duplicateKeyTable.ToDictionary("Name", "Age", "first");

        Assert.Equal(2, result.Count);
        Assert.Equal(typeof(string), result.KeyType);
        Assert.True(result.ContainsKey("Alice"));
        Assert.True(result.ContainsKey("Bob"));
        Assert.Equal(30d, result["Alice"]);
        Assert.Equal(25d, result["Bob"]);
    }

    [Fact]
    public void ToDictionary_DuplicateKeys_TakesLast()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 3, "Alice", 35 } // Duplicate "Name"
            }));

        var result = duplicateKeyTable.ToDictionary("Name", "Age", "last");

        Assert.Equal(2, result.Count);
        Assert.Equal(typeof(string), result.KeyType);
        Assert.True(result.ContainsKey("Alice"));
        Assert.True(result.ContainsKey("Bob"));
        Assert.Equal(35d, result["Alice"]);
        Assert.Equal(25d, result["Bob"]);
    }

    [Fact]
    public void ToDictionaryOfDictionaries_ValidKeyColumn()
    {
        var result = _employeeTable.ToDictionaryOfDictionaries("ID");

        Assert.Equal(3, result.Count);
        Assert.Equal(typeof(double), result.KeyType);
        Assert.True(result.ContainsKey(1d));
        Assert.True(result.ContainsKey(2d));
        Assert.True(result.ContainsKey(3d));

        var aliceDict = result[1d] as XlBlockDictionary;
        Assert.NotNull(aliceDict);
        Assert.Equal(3, aliceDict.Count);
        Assert.Equal("Alice", aliceDict["Name"]);
        Assert.Equal(30d, aliceDict["Age"]);
    }

    [Fact]
    public void ToDictionaryOfDictionaries_MissingKeyColumn_ThrowsException()
    {
        Assert.Throws<ArgumentException>(() => _employeeTable.ToDictionaryOfDictionaries("InvalidKeyColumn"));
    }

    [Fact]
    public void ToDictionaryOfDictionaries_DuplicateKeys_ThrowsException()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 1, "Charlie", 35 } // Duplicate "ID"
            }));

        Assert.Throws<ArgumentException>(() => duplicateKeyTable.ToDictionaryOfDictionaries("ID"));
    }

    [Fact]
    public void ToDictionaryOfDictionaries_DuplicateKeys_First()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 1, "Charlie", 35 } // Duplicate "ID"
            }));

        var result = duplicateKeyTable.ToDictionaryOfDictionaries("ID", "first");

        Assert.Equal(2, result.Count);
        Assert.Equal(typeof(double), result.KeyType);
        Assert.True(result.ContainsKey(1d));
        Assert.True(result.ContainsKey(2d));

        var dict1 = result[1d] as XlBlockDictionary;
        Assert.NotNull(dict1);
        Assert.Equal(3, dict1.Count);
        Assert.Equal(1d, dict1["ID"]);
        Assert.Equal("Alice", dict1["Name"]);
        Assert.Equal(30d, dict1["Age"]);
    }

    [Fact]
    public void ToDictionaryOfDictionaries_DuplicateKeys_Last()
    {
        var duplicateKeyTable = XlBlockTable.Build(XlBlockRange.Build(
            new object[,]
            {
                { "ID", "Name", "Age" },
                { 1, "Alice", 30 },
                { 2, "Bob", 25 },
                { 1, "Charlie", 35 } // Duplicate "ID"
            }));

        var result = duplicateKeyTable.ToDictionaryOfDictionaries("ID", "last");

        Assert.Equal(2, result.Count);
        Assert.Equal(typeof(double), result.KeyType);
        Assert.True(result.ContainsKey(1d));
        Assert.True(result.ContainsKey(2d));

        var dict1 = result[1d] as XlBlockDictionary;
        Assert.NotNull(dict1);
        Assert.Equal(3, dict1.Count);
        Assert.Equal(1d, dict1["ID"]);
        Assert.Equal("Charlie", dict1["Name"]);
        Assert.Equal(35d, dict1["Age"]);
    }

    [Fact]
    public void Project_SpecificColumns_SortTestTable()
    {
        var currentColumns = XlBlockRange.Build(new object[,] { { "Id" }, { "Category" }, { "Average" } });

        var result = _logDataTable.Project(currentColumns, null, null);

        object[,] expectedResult =
        {
            { "Id", "Category", "Average" },
            { 0, "Trace", 38.83 },
            { 1, "Warning", null! },
            { 2, null!, 83.45 },
            { 3, "Critical", 1.77 },
            { 4, "Debug", 53.67 },
            { 5, null!, null! },
            { 6, "Warning", 33.32 },
            { 7, "Critical", 0.82 },
            { 8, "Info", null! },
            { 9, "Info", null! },
            { 10, "Debug", 6.34 },
            { 11, "Warning", null! }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Project_RenameColumns_SortTestTable()
    {
        var currentColumns = XlBlockRange.Build(new object[,] { { "Id" }, { "Category" }, { "Average" } });
        var newColumns = XlBlockRange.Build(new object[,] { { "Identifier" }, { "CategoryType" }, { "AvgValue" } });

        var result = _logDataTable.Project(currentColumns, newColumns, null);

        object[,] expectedResult =
        {
            { "Identifier", "CategoryType", "AvgValue" },
            { 0, "Trace", 38.83 },
            { 1, "Warning", null! },
            { 2, null!, 83.45 },
            { 3, "Critical", 1.77 },
            { 4, "Debug", 53.67 },
            { 5, null!, null! },
            { 6, "Warning", 33.32 },
            { 7, "Critical", 0.82 },
            { 8, "Info", null! },
            { 9, "Info", null! },
            { 10, "Debug", 6.34 },
            { 11, "Warning", null! }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Project_ConvertColumnTypes_SortTestTable()
    {
        var currentColumns = XlBlockRange.Build(new object[,] { { "Id" }, { "ErrorCount" }, { "Average" } });
        var newTypes = XlBlockRange.Build(new object[,] { { "string" }, { "string" }, { "string" } });

        var result = _logDataTable.Project(currentColumns, null, newTypes);

        object[,] expectedResult =
        {
            { "Id", "ErrorCount", "Average" },
            { "0", null!, "38.83" },
            { "1", "25", null! },
            { "2", "21", "83.45" },
            { "3", "2", "1.77" },
            { "4", "62", "53.67" },
            { "5", "22", null! },
            { "6", "11", "33.32" },
            { "7", "45", "0.82" },
            { "8", "0", null! },
            { "9", "101", null! },
            { "10", "62", "6.34" },
            { "11", "45", null! }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void Project_TypeConversionError_SortTestTable_ThrowsException()
    {
        var currentColumns = XlBlockRange.Build(new object[,] { { "Category" } });
        var newTypes = XlBlockRange.Build(new object[,] { { "double" } });

        Assert.Throws<ArgumentException>(() => _logDataTable.Project(currentColumns, null, newTypes));
    }

    [Fact]
    public void GroupBy_SingleColumn_Count_AllNumericColumns()
    {
        var groupColumns = XlBlockRange.Build(new object[,] { { "Category" } });

        var result = _logDataTable.GroupBy(groupColumns, "Count", null, null);

        object[,] expectedResult =
        {
            { "Category", "Id", "ErrorCount", "Average" },
            { "Trace", 1L, 0L, 1L },
            { "Warning", 3L, 3L, 1L },
            { "Critical", 2L, 2L, 2L },
            { "Debug", 2L, 2L, 2L },
            { "Info", 2L, 2L, 0L }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void GroupBy_MultipleColumns_Sum()
    {
        var groupColumns = XlBlockRange.Build(new object[,] { { "Category" }, { "ErrorCount" } });
        var aggregateColumns = XlBlockRange.Build(new object[,] { { "Average" } });
        var result = _logDataTable.GroupBy(groupColumns, "Sum", aggregateColumns, null);

        object[,] expectedResult =
        {
            { "Category", "ErrorCount", "Average" },
            { "Trace", null!, 38.83 },
            { "Warning", 25, 0d },
            { null!, 21, 83.45 },
            { "Critical", 2, 1.77 },
            { "Debug", 62, 60.01 },
            { null!, 22, 0d },
            { "Warning", 11, 33.32 },
            { "Critical", 45, 0.82 },
            { "Info", 0, 0d },
            { "Info", 101, 0d },
            { "Warning", 45, 0d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void GroupBy_Average()
    {
        var groupColumns = XlBlockRange.Build(new object[,] { { "Category" } });

        var result = _logDataTable.GroupBy(groupColumns, "Average", XlBlockRange.Build(new object[,] { { "Average" } }), null);

        object[,] expectedResult =
        {
            { "Category", "Average" },
            { "Trace", 38.83 },
            { "Warning", 11.1067 },
            { "Critical", 1.295 },
            { "Debug", 30.005 },
            { "Info", 0d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void GroupBy_Max()
    {
        var groupColumns = XlBlockRange.Build(new object[,] { { "Category" } });

        var result = _logDataTable.GroupBy(groupColumns, "Max", XlBlockRange.Build(new object[,] { { "Average" } }), null);

        object[,] expectedResult =
        {
            { "Category", "Average" },
            { "Trace", 38.83 },
            { "Warning", 33.32 },
            { "Critical", 1.77 },
            { "Debug", 53.67 },
            { "Info", 0d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    [Fact]
    public void GroupBy_SpecifiedNames()
    {
        var groupColumns = XlBlockRange.Build(new object[,] { { "Category" } });

        var result = _logDataTable.GroupBy(groupColumns, "Max", XlBlockRange.Build(new object[,] { { "ErrorCount", "Average" } }),
            XlBlockRange.Build(new object[,] { { "ErrorMax", "MeanMax" } }));

        object[,] expectedResult =
        {
            { "Category", "ErrorMax", "MeanMax" },
            { "Trace", 0, 38.83 },
            { "Warning", 45, 33.32 },
            { "Critical", 45, 1.77 },
            { "Debug", 62, 53.67 },
            { "Info", 101, 0d }
        };
        AssertTableMatchesExpected(expectedResult, result);
    }

    #endregion
}
