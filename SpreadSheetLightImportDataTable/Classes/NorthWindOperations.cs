
// For access SQL-Server database
using NorthWind2020Library.Models;

using System.Collections.Generic;

// For referencing a DataTable
using System.Data;

using DocumentFormat.OpenXml.Spreadsheet;

// Provides the ability to transform a list of CustomersForExcel to a DataTable
using FastMember;

// Provides classes to interact with Microsoft Excel 2007+
using SpreadsheetLight;

using Color = System.Drawing.Color;

namespace SpreadSheetLightImportDataTable.Classes
{
    public class NorthWindOperations
    {
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
            // SpreadSheetLight also has a DataTable so we must point to the correct class.
            DataTable table = new();

            // ordinal index to the Modified column/property in the model
            int dateColumnIndex = 6;

            /*
             * Creates an instance of ObjectReader for transforming a List of CustomersForExcel
             * to a DataTable
             */
            using var reader = ObjectReader.Create(list);

            // Load List of CustomersForExcel into a DataTable
            table.Load(reader);
            
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
            var headerStyle = HeaderStye(document);

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
             * By default the first sheet name is Sheet1, let's provide a meaningful name
             */
            document.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Customers");

            /*
             * Format the first row with column names
             */
            document.SetCellStyle(1, 1, 1, 6, headerStyle);

            // one row below header
            document.SetActiveCell("A2");

            // ensure header is visible when scrolling down
            document.FreezePanes(1,6);

            document.SaveAs(fileName);
            
        }

        /// <summary>
        /// Create the first row format/style
        /// </summary>
        /// <param name="document">Instance of a <see cref="SLDocument"/></param>
        /// <returns>A <see cref="SLStyle"/></returns>
        public static SLStyle HeaderStye(SLDocument document)
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
}
