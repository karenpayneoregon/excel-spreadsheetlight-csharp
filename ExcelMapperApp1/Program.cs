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
        await Operations.ReadProductsAndUpdate();
        await Operations.ReadProductsCreateCopyWithLessProperties();
        await Operations.SingleColumnExample();
        await Operations.CustomersToDatabase();

        Console.ReadLine();
    }
    
}