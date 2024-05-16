
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



