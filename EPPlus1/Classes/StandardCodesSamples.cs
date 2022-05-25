using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus1.Models;
using OfficeOpenXml;
using Spectre.Console;

namespace EPPlus1.Classes
{
    public class StandardCodesSamples
    {
        /// <summary>
        /// Took
        /// https://github.com/EPPlusSoftware/EPPlus.Sample.NetCore/blob/master/02-ReadWorkbook/ReadWorkbookSample.cs
        ///
        /// And modified for Customers.xlsx
        /// </summary>
        /// <remarks>
        /// C:\Users\paynek\.nuget\packages\epplus\6.0.4\readme.txt
        /// </remarks>
        public static void Sample1()
        {

            Table customerTable = ConsoleOperations.DisplayTable();
            

            var filePath = FileUtil.GetFileInfo("02-ReadWorkbook", "Customers.xlsx").FullName;
            Console.WriteLine("Reading {0}", filePath);
            Console.WriteLine();
            FileInfo existingFile = new FileInfo(filePath);
            using ExcelPackage package = new(existingFile);
            
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
            var lastRow = worksheet.Dimension.End.Row;
            var lastColumn = worksheet.Dimension.End.Column;

            Console.WriteLine($"Row: {lastRow} Col: {lastColumn}");

            List<CustomerExcelItem> list = new();

            for (int rowIndex = 2; rowIndex < lastRow; rowIndex++)
            {
                var modDateValue = worksheet.Cells[rowIndex, 6].Text;
                var idValue = worksheet.Cells[rowIndex, lastColumn].Text;

                if (DateTime.TryParse(modDateValue, out var modifiedDate) && int.TryParse(idValue, out var id))
                {
                    list.Add(new CustomerExcelItem()
                    {
                        RowIndex = rowIndex,
                        Id = id, 
                        CompanyName = worksheet.Cells[rowIndex, 1].Text,
                        Title = worksheet.Cells[rowIndex, 2].Text,
                        Contact = worksheet.Cells[rowIndex, 3].Text,
                        Country = worksheet.Cells[rowIndex, 4].Text,
                        Phone = worksheet.Cells[rowIndex, 5].Text,
                        ModifiedDate = modifiedDate
                    });
                }
            }


            foreach (var item in list)
            {

                customerTable.AddRow(
                    item.RowIndex.ToString(),
                    item.Id.ToString(),
                    item.CompanyName,
                    item.Title,
                    item.Contact,
                    item.Country,
                    item.Phone,
                    item.ModifiedDate?.ToString("d")!
                );
            }

            AnsiConsole.Write(customerTable);

        }


    }
}
