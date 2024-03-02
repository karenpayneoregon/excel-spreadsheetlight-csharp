using ExcelMapperApp1.Data;
using ExcelMapperApp1.Models;
using Ganss.Excel;

namespace ExcelMapperApp1;

/// <summary>
/// Before running
/// 1. Create Examples database under .\SQLEXPRESS
/// 2. Run Populate.sql under the scripts folder
/// </summary>
internal partial class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            const string excelFile = "Customers.xlsx";
            var excel = new ExcelMapper();
            var customers = (await excel.FetchAsync<Customers>(excelFile,"Customers")).ToList();
            await using var context = new Context();
            context.Customers.AddRange(customers);
            var affected = await context.SaveChangesAsync();
            Console.WriteLine(affected > 0 ? $"Saved {affected} records" : "Failed");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        AnsiConsole.MarkupLine("[yellow]Hello[/]");
        Console.ReadLine();
    }
}