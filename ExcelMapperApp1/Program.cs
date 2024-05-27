using ExcelMapperApp1.Data;
using ExcelMapperApp1.Models;
using Ganss.Excel;
using ExcelMapperApp1.Classes;

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
        await ExcelMapperOperations.NestedReadPeople();
        await ExcelMapperOperations.ReadProductsAndUpdate();
        await ExcelMapperOperations.ReadProductsCreateCopyWithLessProperties();
        await ExcelMapperOperations.SingleColumnExample();
        await ExcelMapperOperations.CustomersToDatabase();

        var (rows, hasIssues) = LightOperations.Iterate();
        AnsiConsole.MarkupLine(hasIssues ? $"[red]Issues found in rows: {string.Join(",", rows)}[/]" : "[green]No issues found[/]");

        Console.ReadLine();
    }
    
}