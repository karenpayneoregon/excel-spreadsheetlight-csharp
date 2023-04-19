using System.Diagnostics;
using Spectre.Console;
using SpreadSheetLightConsoleApp.Classes;
using SpreadSheetLightConsoleApp.Models;

namespace SpreadSheetLightConsoleApp;

partial class Program
{
    static void Main(string[] args)
    {
        var value = "karen";

        SearchItem searchItem = new("Demo1.xlsx", "sheet1", value, 
            StringComparison.InvariantCultureIgnoreCase);


        (IReadOnlyList<FoundItemImmutable> items, Exception exception) = 
            ExcelOperations.FindText(searchItem);
            
        if (exception is null)
        {
            if (items.Count >0)
            {
                AnsiConsole.MarkupLine($"[white]Found[/] [cyan]{value}[/] [white]{items.Count}[/] times");
                foreach (var foundItem in items)
                {
                    Console.WriteLine($"{foundItem}");
                }
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Did not find {value}[/]");
            }

            Console.ReadLine();
        }
        else
        {
            Debug.WriteLine(exception.Message);
        }
           
            
    }

}