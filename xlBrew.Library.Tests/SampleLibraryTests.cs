namespace xlBrew.Library.Tests;

using xlBrew.Library;

public class SampleLibraryTests
{
    [Fact]
    public void AddThemTests()
    {
        var result = SampleLibrary.AddNumbers(1, 2);
        Assert.Equal(3d, result);
    }
}