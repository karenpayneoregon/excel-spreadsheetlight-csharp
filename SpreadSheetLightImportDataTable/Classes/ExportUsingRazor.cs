using SpreadsheetLight;
using System.Data;
#pragma warning disable CS8619 // Nullability of reference types in value doesn't match target type.

namespace SpreadSheetLightImportDataTable.Classes;
public class ExportUsingRazor
{
    public static DataTable GetTable(string tableName = "whatever")
    {
        DataTable table = new();
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

    public static (bool success, Exception exception) Export()
    {
        string fileName = "C:\\OED\\Demo.xlsx";

        if (!Directory.Exists(Path.GetDirectoryName(fileName)))
        {
            return (false, new FileNotFoundException(fileName));
        }

        DataTable table = GetTable();
        try
        {
            ExportToExcel(table, @fileName, true, "Data");
            return (true, null);
        }
        catch (Exception localException)
        {
            return (false, localException);
        }

    }

    /// <summary>
    /// Exports the provided <see cref="DataTable"/> to an Excel file.
    /// </summary>
    /// <param name="table">The <see cref="DataTable"/> to export.</param>
    /// <param name="fileName">The name of the Excel file to save the data to.</param>
    /// <param name="includeHeader">A boolean indicating whether to include the header row in the export.</param>
    /// <param name="sheetName">The name of the worksheet within the Excel file.</param>
    public static void ExportToExcel(DataTable table, string fileName, bool includeHeader, string sheetName)
    {
        using var document = new SLDocument();

        // import to first row, first column
        document.ImportDataTable(1, SLConvert.ToColumnIndex("A"), table, includeHeader);

        // give sheet a useful name
        document.RenameWorksheet(SLDocument.DefaultFirstSheetName, sheetName);

        document.SaveAs(fileName);
    }

    /// <summary>
    /// Exports the provided <see cref="DataTable"/> to an Excel file.
    /// </summary>
    /// <param name="table">The <see cref="DataTable"/> to export.</param>
    /// <param name="fileName">The name of the Excel file to save the data to.</param>
    /// <param name="includeHeader">A boolean indicating whether to include the header row in the export.</param>
    /// <param name="sheetName">The name of the worksheet within the Excel file.</param>
    /// <param name="row">The starting row in the Excel sheet where the data should be imported.</param>
    public static void ExportToExcel(DataTable table, string fileName, bool includeHeader, string sheetName, int row)
    {
        using var document = new SLDocument();

                document.ImportDataTable(row, SLConvert.ToColumnIndex("A"), table, includeHeader);

        // give sheet a useful name
        document.RenameWorksheet(SLDocument.DefaultFirstSheetName, sheetName);

        document.SaveAs(fileName);
    }

    /// <summary>
    /// Exports the provided <see cref="DataTable"/> to an Excel file with additional formatting options.
    /// </summary>
    /// <param name="table">The <see cref="DataTable"/> to export.</param>
    /// <param name="fileName">The name of the Excel file to save the data to.</param>
    /// <param name="includeHeader">A boolean indicating whether to include the header row in the export.</param>
    /// <param name="sheetName">The name of the worksheet within the Excel file.</param>
    /// <param name="row">The starting row in the Excel sheet where the data should be imported.</param>
    /// <remarks>
    /// This method applies a specific date format ("mm-dd-yyyy") to a designated column in the Excel file.
    /// </remarks>
    public static void ExportToExcel1(DataTable table, string fileName, bool includeHeader, string sheetName, int row)
    {
        using var document = new SLDocument();


        document.ImportDataTable(row, SLConvert.ToColumnIndex("A"), table, includeHeader);

        // give sheet a useful name
        document.RenameWorksheet(SLDocument.DefaultFirstSheetName, sheetName);

        SLStyle dateStyle = document.CreateStyle();
        dateStyle.FormatCode = "mm-dd-yyyy";
        // format a specific column using above style
        int dateColumnIndex = 6;
        document.SetColumnStyle(dateColumnIndex, dateStyle);

        document.SaveAs(fileName);
    }
}
