using ExcelMapperApp1.Data;
using ExcelMapperApp1.Models;
using Ganss.Excel;
using static ExcelMapperApp1.Classes.LightHelpers;
using static ExcelMapperApp1.Classes.SpectreConsoleHelpers;

namespace ExcelMapperApp1.Classes;
internal class ExcelMapperOperations
{
    /// <summary>
    /// There are two columns, here we ignore the second column
    /// </summary>
    public static async Task SingleColumnExample()
    {
        PrintCyan();

        const string excelFile = "Excel1.xlsx";

        ExcelMapper excel = new();

        var list = (await excel.FetchAsync<Sheet1>(excelFile, nameof(Sheet1))).ToList();

    }

    /// <summary>
    /// Read nested data from Nested.xlsx where <see cref="Person"/> has an <see cref="Address"/>
    /// </summary>
    public static async Task NestedReadPeople()
    {
        PrintCyan();

        const string excelFile = "Nested.xlsx";

        ExcelMapper excel = new();

        var contactList =  (await excel.FetchAsync<Person>(excelFile, "Contacts")).ToList();
        
        AnsiConsole.MarkupLine(ObjectDumper.Dump(contactList)
            .Replace("{Person}", "[cyan]{Person}[/]")
            .Replace("Address:", "[cyan]Address:[/]"));

    }
    /// <summary>
    /// Read products from Products.xlsx as list of <see cref="Products"/> then update
    /// several products and save to a new file ProductsOut.xlsx
    /// </summary>
    public static async Task ReadProductsAndUpdate()
    {

        PrintCyan();

        const string excelReadFile = "Products.xlsx";
        const string excelWriteFile = "ProductsOut.xlsx";

        if (File.Exists(excelWriteFile))
        {
            try
            {
                File.Delete(excelWriteFile);
            }
            catch (Exception ex)
            {
                ex.ColorWithCyanFuchsia();
                return;
            }
        }

        ExcelMapper excel = new();

        var products = excel.Fetch<Products>(excelReadFile, nameof(Products)).OrderBy(x => x.ProductName).ToList();

        var p1 = products.FirstOrDefault(x => x.ProductName == "CÃ\u00b4te de Blaye");
        if (p1 is not null)
        {
            p1.ProductName = "Cafe de Blave";
        }

        var p2 = products.FirstOrDefault(x => x.Supplier == "Aux joyeux ecclÃ\u00a9siastiques");
        if (p2 is not null)
        {
            p2.Supplier = "Aux Joy";
        }   

        var p3 = products.FirstOrDefault(x => x.ProductID == 48);
        if (p3 is not null)
        {
            products.Remove(p3);
        }

        await excel.SaveAsync(excelWriteFile, products, "Products");

    }
    /// <summary>
    /// Read products from Products.xlsx as list of <see cref="Products"/> then write to a new
    /// file as <see cref="ProductItem"/> ProductsCopy.xlsx
    /// </summary>
    /// <returns></returns>
    public static async Task ReadProductsCreateCopyWithLessProperties()
    {

        PrintCyan();

        const string excelReadFile = "Products.xlsx";
        const string excelWriteFile = "ProductsCopy.xlsx";

        if (File.Exists(excelWriteFile))
        {
            try
            {
                File.Delete(excelWriteFile);
            }
            catch (Exception ex)
            {
                ex.ColorWithCyanFuchsia();
                return;
            }
        }

        ExcelMapper excel = new();

        var products = (await excel.FetchAsync<Products>(excelReadFile,
            nameof(Products))).ToList();

        var productItems = products.Select(p => new ProductItem
        {
            ProductID = p.ProductID,
            ProductName = p.ProductName,
            CategoryName = p.CategoryName,
            UnitPrice = p.UnitPrice
        }).ToList();

        await new ExcelMapper().SaveAsync("productsCopy.xlsx", productItems, "Products");
    }

    /// <summary>
    /// Read Customers.xlsx data as list of <see cref="Customers"/> then write to database
    /// using EF Core
    /// </summary>
    public static async Task CustomersToDatabase()
    {
        
        PrintCyan();
        const string excelFile = "Customers.xlsx";

        if (SheetExists(excelFile, nameof(Customers)) == false)
        {
            AnsiConsole.MarkupLine($"[red]Sheet {nameof(Customers)} not found in {excelFile}[/]");
            return;
        }

        try
        {
            DapperOperations operations = new();
            operations.Reset();

            ExcelMapper excel = new();
            await using var context = new Context();

            var customers = (await excel.FetchAsync<Customers>(excelFile, nameof(Customers))).ToList();

            context.Customers.AddRange(customers);
            var affected = await context.SaveChangesAsync();

            AnsiConsole.MarkupLine(affected > 0 ? $"[cyan]Saved[/] [b]{affected}[/] [cyan]records[/]" : "[red]Failed[/]");
        }
        catch (Exception ex)
        {
            ex.ColorWithCyanFuchsia();
        }
    }
}
