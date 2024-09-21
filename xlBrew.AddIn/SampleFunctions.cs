namespace xlBrew.AddIn;

using ExcelDna.Integration;
using xlBrew.Library;

public static class SampleFunctions
{
    [ExcelFunction(Description = "Say hello")]
    public static string SayHello(string name)
    {
        return $"Hello {name}";
    }

    [ExcelFunction(Description = "Add numbers")]
    public static double AddThem(double a, double b)
    {
        return SampleLibrary.AddNumbers(a, b);
    }
}