using SpreadSheetLightOther.Classes;

namespace SpreadSheetLightOther;

internal class Program
{
    static void Main(string[] args)
    {
        var (success, exception) = ExcelOperations.InsertColumnData(
            "SomeFile.xlsx", 4, 5, new List<string>()
            {
                "A", "B", "C", "D",
            });

        if (success)
        {
            Console.WriteLine("Done");
        }
        else if (exception is not null)
        {
            Console.WriteLine(exception.Message);
        }

        Console.ReadLine();
    }
}
