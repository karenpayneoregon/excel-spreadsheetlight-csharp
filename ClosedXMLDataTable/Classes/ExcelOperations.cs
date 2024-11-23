using System.Data;
using ClosedXML.Excel;

namespace ClosedXMLDataTable.Classes;

public class ExcelOperations
{
    /// <summary>
    /// Creates an Excel file with a worksheet populated with sample data.
    /// </summary>
    /// <param name="fileName">The name of the Excel file to create.</param>
    public static void Create(string fileName)
    {
        var workbook = new XLWorkbook();

        var dataTable = GetTable("Information");

        // Add a DataTable as a worksheet
        workbook.Worksheets.Add(dataTable);
        workbook.Worksheets.First().Columns().AdjustToContents();

        // this will throw a runtime exception if the file is open as with other libraries used here
        workbook.SaveAs(fileName);
    }

    /// <summary>
    /// Mock up some data
    /// </summary>
    private static DataTable GetTable(string tableName)
    {
        DataTable table = new ();
        table.TableName = tableName;
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
        table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
        table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));

        return table;
    }

    /// <summary>
    /// Writes a value to a specific cell in an existing Excel file.
    /// </summary>
    /// <param name="reportFilePath">The path to the Excel file.</param>
    /// <param name="row">The row number of the cell (1-based index).</param>
    /// <param name="col">The column number of the cell (1-based index).</param>
    /// <param name="value">The value to write to the cell.</param>
    public static void WriteToCell(string reportFilePath, int row, int col, string value)
    {
        using var workbook = new XLWorkbook(reportFilePath);
        var worksheet = workbook.Worksheets.Worksheet(1);
        worksheet.Cell(row, col).Value = value;
        workbook.SaveAs(reportFilePath);
    }

}