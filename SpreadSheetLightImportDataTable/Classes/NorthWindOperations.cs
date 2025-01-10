
// For access SQL-Server database
using NorthWind2020Library.Models;


// For referencing a DataTable
using System.Data;

using DocumentFormat.OpenXml.Spreadsheet;

// Provides the ability to transform a list of CustomersForExcel to a DataTable
using FastMember;

// Provides classes to interact with Microsoft Excel 2007+
using SpreadsheetLight;
using SpreadSheetLightImportDataTable.LanguageExtensions;
using Color = System.Drawing.Color;
#pragma warning disable CS8602

namespace SpreadSheetLightImportDataTable.Classes;

public class NorthWindOperations
{
    /// <summary>
    /// Exports the provided <see cref="DataTable"/> to an Excel file.
    /// </summary>
    /// <param name="table">The <see cref="DataTable"/> to export.</param>
    /// <param name="fileName">The name of the file to save the Excel document as.</param>
    /// <param name="includeHeader">A boolean indicating whether to include the header row in the Excel file.</param>
    /// <param name="sheetName">The name to assign to the Excel worksheet.</param>
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
    /// Convert a list to a DataTable for SpreadSheetLight to import using one method.
    /// </summary>
    /// <param name="list">list to place into an Excel file</param>
    /// <param name="fileName">Where to save</param>
    /// <remarks>
    /// No try/catch, instead the caller uses catch-when
    /// </remarks>
    public static void CustomersToExcel(List<CustomersForExcel> list, string fileName)
    {
        var table = list.ToDataTable();
        
        // ordinal index to the Modified column/property in the model
        int dateColumnIndex = 6;
        
        /*
         * Rearrange visual order of data columns
         */
        table.Columns["Title"].SetOrdinal(1);
        table.Columns["Modified"].SetOrdinal(dateColumnIndex);
        table.Columns["id"].SetOrdinal(6);

        table.Columns["CompanyName"].ColumnName = "Company";

        // Create an instance of SpreadSheetLight document
        using var document = new SLDocument();

        // Setup first row style for worksheet
        var headerStyle = HeaderStyle(document);

        // Create a format/style for Modified data column
        SLStyle dateStyle = document.CreateStyle();
        dateStyle.FormatCode = "mm-dd-yyyy";

        /*
         * Import DataTable to first row, first column in Sheet1 and include column names
         */

        document.ImportDataTable(1, SLConvert.ToColumnIndex("A"), table, true);
        
        /*
         * Hide the primary key column
         */
        document.HideColumn(7, 7);
        document.SetColumnStyle(dateColumnIndex, dateStyle);

        for (int columnIndex = 1; columnIndex < table.Columns.Count; columnIndex++)
        {
            document.AutoFitColumn(columnIndex);
        }

        document.AutoFitColumn(dateColumnIndex + 1);
            
        /*
         * By default, the first sheet name is Sheet1, let's provide a meaningful name
         */
        document.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Customers");

        /*
         * Format the first row with column names
         */
        document.SetCellStyle(1, 1, 1, 6, headerStyle);

        // one row below header
        document.SetActiveCell("A2");
        //SLPageSettings pageSettings = new SLPageSettings() {ZoomScale = 80};
        document.SetPageSettings(new SLPageSettings() { ZoomScale = 80 });
        // ensure header is visible when scrolling down
        document.FreezePanes(1,6);

        document.SaveAs(fileName);
            
    }

    /// <summary>
    /// Create the first row format/style
    /// </summary>
    /// <param name="document">Instance of a <see cref="SLDocument"/></param>
    /// <returns>A <see cref="SLStyle"/></returns>
    public static SLStyle HeaderStyle(SLDocument document)
    {
            
        SLStyle headerStyle = document.CreateStyle();

        headerStyle.Font.Bold = true;
        headerStyle.Font.FontColor = Color.White;
        headerStyle.Fill.SetPattern(
            PatternValues.LightGray,
            SLThemeColorIndexValues.Accent1Color,
            SLThemeColorIndexValues.Accent5Color);

        return headerStyle;
    }
}