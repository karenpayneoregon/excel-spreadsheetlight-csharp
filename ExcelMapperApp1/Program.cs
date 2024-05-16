using ExcelMapperApp1.Data;
using ExcelMapperApp1.Models;
using Ganss.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using Dapper;
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
        await SingleColumnExample();

        await CustomersToDatabase();

        Console.ReadLine();
    }

    /// <summary>
    /// There are two columns, here we ignore the second column
    /// </summary>
    private static async Task SingleColumnExample()
    {
        const string excelFile = "Excel1.xlsx";
        ExcelMapper excel = new();
        var list = (await excel.FetchAsync<Sheet1>(excelFile, nameof(Sheet1))).ToList();
    }

    private static async Task CustomersToDatabase()
    {
        try
        {
            DapperOperations operations = new();
            operations.Reset();

            const string excelFile = "Customers.xlsx";
            ExcelMapper excel = new();
            await using var context = new Context();

            var customers = (await excel.FetchAsync<Customers>(excelFile, 
                nameof(Customers))).ToList();

            context.Customers.AddRange(customers);
            var affected = await context.SaveChangesAsync();

            AnsiConsole.MarkupLine(affected > 0 ? $"[cyan]Saved[/] [b]{affected}[/] [cyan]records[/]" : "[red]Failed[/]");
        }
        catch (Exception ex)
        {
            ex.ColorWithCyanFuchsia();
        }

        AnsiConsole.MarkupLine("[yellow]Done[/]");
    }
}

internal class DapperOperations
{
    private IDbConnection _cn = new SqlConnection(ConnectionString());
    public void Reset()
    {
        _cn.Execute($"DELETE FROM dbo.{nameof(Customers)}");
        _cn.Execute($"DBCC CHECKIDENT ({nameof(Customers)}, RESEED, 0)");
    }
}