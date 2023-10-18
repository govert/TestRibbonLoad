using ExcelDna.Integration;

public static class Functions
{
    [ExcelFunction]
    public static object SayHello()
    {
        return "Hello from TestRibbonLoad add-in";
    }
}