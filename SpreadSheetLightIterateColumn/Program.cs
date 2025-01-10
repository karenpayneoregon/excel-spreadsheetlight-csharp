
using System.Data;
using SpreadsheetLight;
using DataTable = System.Data.DataTable;

namespace SpreadSheetLightIterateColumn;

internal partial class Program
{
    static void Main()
    {

        using SLDocument document = new("Excel1.xlsx", "Sheet1");
        var dt = CreateDataTable();

        for (int index = 2; index < document.GetWorksheetStatistics().EndRowIndex + 1; index++)
        {
            dt.Rows.Add(null, document.GetCellValueAsString(index, 1));
        }

        var table = CreateTable();
        foreach (DataRow row in dt.Rows)
        {
            table.AddRow(row.Field<int>("Id").ToString(), row.Field<string>("FirstName"));
        }

        AnsiConsole.Write(table);
        Console.ReadLine();
    }

    private static Table CreateTable()
        => new Table().RoundedBorder().LeftAligned()
            .AddColumn("[cyan]Id[/]")
            .AddColumn("[cyan]First[/]")
            .BorderColor(Color.LightSlateGrey)
            .Title("[LightGreen]Excel data[/]");
    /// <summary>
    /// Creates and initializes a new <see cref="System.Data.DataTable"/> with predefined columns.
    /// </summary>
    /// <remarks>
    /// The created <see cref="System.Data.DataTable"/> contains the following columns:
    /// <list type="bullet">
    /// <item>
    /// <description><c>Id</c>: An auto-incrementing integer column starting from 1.</description>
    /// </item>
    /// <item>
    /// <description><c>FirstName</c>: A string column for storing first names.</description>
    /// </item>
    /// </list>
    /// </remarks>
    /// <returns>
    /// A new instance of <see cref="System.Data.DataTable"/> with the specified columns.
    /// </returns>
    private static DataTable CreateDataTable()
    {
        DataTable table = new();
        table.Columns.Add("Id", typeof(int));
        table.Columns["Id"].AutoIncrement = true;
        table.Columns["Id"].AutoIncrementSeed = 1;
        table.Columns.Add("FirstName", typeof(string));
        return table;
    }
}



