using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus1.Classes
{
    class DataOperations
    {
        private static string ConnectionString =
            "Data Source=.\\SQLEXPRESS;Initial Catalog=NorthWind2020;" +
            "Integrated Security=True";

        public static void Contacts(string _excelBaseFolder)
        {
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _excelBaseFolder, "Contacts.xlsx");
            using var cn = new SqlConnection() { ConnectionString = ConnectionString };
            string selectStatement =
                "SELECT C.ContactId As Id,C.FirstName + ' ' + C.LastName AS [Contact Name], Countries.Name " + 
                "FROM Customers AS Cust INNER JOIN Contacts AS C ON Cust.ContactId = C.ContactId "+ 
                "INNER JOIN Countries ON Cust.CountryIdentifier = Countries.CountryIdentifier;";

            using var cmd = new SqlCommand() { Connection = cn, CommandText = selectStatement };
            cn.Open();
            var reader = cmd.ExecuteReader();
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Contacts");
            worksheet.Cells["A1"].LoadFromDataReader(reader, true);

            worksheet.Cells.AutoFitColumns();
            using (ExcelRange range = worksheet.Cells[$"A1:C{worksheet.Dimension.End.Row}"])
            {
                ExcelTableCollection tableCollection = worksheet.Tables;
                ExcelTable table = tableCollection.Add(range, "ContactsTable");
                table.TableStyle = TableStyles.Light1;
            }

            package.SaveAs(filePath);
        }
    }
}
