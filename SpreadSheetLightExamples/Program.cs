using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using NorthWind2020Library.Classes;
using NorthWind2020Library.Models;
using SpreadSheetLightImportDataTable.Classes;
using SpreadSheetLightLibrary.Classes;
using IO = System.IO;

namespace SpreadSheetLightExamples
{
    partial class Program
    {
        static void Main(string[] args)
        {
            //ImportTabTextFile();
            CreatePopulateCustomerData();
        }

        /// <summary>
        /// Import tab delimited text file with minor formatting
        /// </summary>
        private static void ImportTabTextFile()
        {
            var importFileName = "Products.txt";
            var excelFileName = "ProductsImported.xlsx";
            var sheetName = "Products";

            var success = Operations.ImportTabDelimitedTextFile(
                importFileName,
                excelFileName, 
                sheetName);

            
            if (success)
            {
                var count = Operations.GetWorkSheetLastRow(excelFileName, sheetName);
                Console.WriteLine($"Wrote {count} rows to {sheetName} in {excelFileName}");
            }
            else
            {
                Console.WriteLine("Failed");
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Create new Excel file, format data
        /// </summary>
        private static void CreatePopulateCustomerData()
        {
            List<CustomersForExcel> list = CustomerOperations.FromJson();

            try
            {
                NorthWindOperations.CustomersToExcel(list, IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers.xlsx"));
                Console.WriteLine("Done");
            }
            catch (Exception exception) when (exception.Message.Contains("The process cannot access the file"))
            {
                Console.WriteLine("Hey you have the spreadsheet open, can not save!!!");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Something went wrong '{exception.Message}'");
            }
        }
    }
}
