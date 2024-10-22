namespace XlBlocks.AddIn.Tests.Cache;

using XlBlocks.AddIn.Cache;

public class CacheHelperTests
{

    [Fact]
    public void CheckColumnToLetter()
    {
        var testCases = new Dictionary<string, int>()
        {
            ["A"] = 1,
            ["CV"] = 100,
            ["ABC"] = 731,
            ["XFD"] = 16384
        };

        foreach (var testCase in testCases)
        {
            var referenceString = CacheHelper.ColumnNumberToLetter(testCase.Value);
            Assert.Equal(testCase.Key, referenceString);
        }
    }
}
