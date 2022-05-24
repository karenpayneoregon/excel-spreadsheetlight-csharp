﻿using System;
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
            ImportTabTextFile();
            Console.ReadLine();
        }

        /// <summary>
        /// Import tab delimited text file with minor formatting
        /// </summary>
        private static void ImportTabTextFile()
        {
            var success = Operations.ImportTabDelimitedTextFile(
                "Products.txt", 
                "ProductsImported.xlsx", 
                "Products");

            Console.WriteLine(success);
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
