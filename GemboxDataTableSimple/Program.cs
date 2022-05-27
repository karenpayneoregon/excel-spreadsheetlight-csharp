using System.Data;
using GemBox.Spreadsheet;

namespace DataTableToSheet
{
    class Program
    {
        static void Main()
        {
            // Karen has Professional version which has been used in several apps
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("DataTable to Sheet");

            var dataTable = new DataTable();

            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("FirstName", typeof(string));
            dataTable.Columns.Add("LastName", typeof(string));

            dataTable.Rows.Add(100, "John", "Doe");
            dataTable.Rows.Add(101, "Fred", "Nurk");
            dataTable.Rows.Add(103, "Hans", "Meier");
            dataTable.Rows.Add(104, "Ivan", "Horvat");
            dataTable.Rows.Add(105, "Jean", "Dupont");
            dataTable.Rows.Add(106, "Mario", "Rossi");

            worksheet.Cells[0, 0].Value = "DataTable insert example:";

            // add a header on row 2
            worksheet.Rows[2].Style = workbook.Styles[BuiltInCellStyleName.Heading1];

            // Insert DataTable to an Excel worksheet.
            worksheet.InsertDataTable(dataTable,
                new InsertDataTableOptions()
                {
                    ColumnHeaders = true,
                    StartRow = 2
                });

            worksheet.Columns[0].Width = 10 * 256;
            worksheet.Columns[1].Width = 20 * 256;
            worksheet.Columns[2].Width = 20 * 256;

            workbook.Save("DataTable to Sheet.xlsx");
        }
    }
}