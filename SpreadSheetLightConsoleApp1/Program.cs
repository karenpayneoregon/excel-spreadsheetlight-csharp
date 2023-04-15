using System.Diagnostics;
using System.Runtime.CompilerServices;
using SpreadSheetLightConsoleApp.Classes;
using SpreadSheetLightConsoleApp.Models;

namespace SpreadSheetLightConsoleApp;

class Program
{
    static void Main(string[] args)
    {
        var value = "karen";

        SearchItem searchItem = new(
            "Demo1.xlsx", 
            "sheet1", 
            value, 
            StringComparison.InvariantCultureIgnoreCase);


        (IReadOnlyList<FoundItemImmutable> items, Exception exception) = ExcelOperations.FindText(searchItem);
            
        if (exception is null)
        {
            if (items.Count >0)
            {
                Console.WriteLine($"Found {value} {items.Count} times");
                foreach (var foundItem in items)
                {
                    Console.WriteLine($"{foundItem}");
                }
            }
            else
            {
                Console.WriteLine($"Did not find {value}");
            }

            Console.ReadLine();
        }
        else
        {
            Debug.WriteLine(exception.Message);
        }
           
            
    }

    [ModuleInitializer]
    public static void Init1()
    {
        Console.Title = "Working with immutable read from Excel";
    }
}