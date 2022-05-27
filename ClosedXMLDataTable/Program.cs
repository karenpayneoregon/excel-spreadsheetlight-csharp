using System;
using System.IO;
using ClosedXMLDataTable.Classes;

namespace ClosedXMLDataTable
{
    partial class Program
    {
        static void Main(string[] args)
        {
            ExcelOperations
                .Create(
                    Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,"NewFile.xlsx"));
        }
    }
}
