﻿using SpreadsheetLight;
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
    public static void ExportToExcel(DataTable table, string fileName, bool includeHeader, string sheetName)
    {
        using var document = new SLDocument();

        // import to first row, first column
        document.ImportDataTable(1, SLConvert.ToColumnIndex("A"), table, includeHeader);

        // give sheet a useful name
        document.RenameWorksheet(SLDocument.DefaultFirstSheetName, sheetName);

        document.SaveAs(fileName);
    }
}
