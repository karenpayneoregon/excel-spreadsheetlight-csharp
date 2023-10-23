using SpreadsheetLight;

namespace ExampleApp1;

internal partial class Program
{
    static void Main(string[] args)
    {
        SetStringCellValue("Customers.xlsx", "Example", "A1", "Hello");
        SetIntCellValue("Customers.xlsx", "Example", "A2", 1);
        Console.ReadLine();
    }
    // set string value for a cell
    public static void SetStringCellValue(string excelFileName, string sheetName, string cell, string cellValue)
    {
        if (!File.Exists(excelFileName)) return;
        try
        {
            using SLDocument document = new(excelFileName, sheetName);
            if (!document.GetSheetNames(false).Contains(sheetName)) return;
            document.SetCellValue(cell, cellValue);
            document.Save();
            Console.WriteLine("Example 1 done");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Example 1 someone has the file open outside of this app");
        }

    }
    // Set int value for a cell
    public static void SetIntCellValue(string excelFileName, string sheetName, string cell, int cellValue)
    {
        if (!File.Exists(excelFileName)) return;
        try
        {
            using SLDocument document = new(excelFileName, sheetName);
            if (!document.GetSheetNames(false).Contains(sheetName)) return;
            document.SetCellValue(cell, cellValue);
            document.Save();
            Console.WriteLine("Example 2 done");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Example 2 someone has the file open outside of this app");
        }

    }
}