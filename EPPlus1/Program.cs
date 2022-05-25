using System;
using System.IO;
using System.Runtime.CompilerServices;
using EPPlus1.Classes;
using OfficeOpenXml;

namespace EPPlus1
{
    partial class Program
    {
        static void Main(string[] args)
        {
            StandardCodesSamples.Sample1();
            Console.ReadLine();
        }

        [ModuleInitializer]
        public static void Init1()
        {
            Console.Title = "Working Excel";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
    }

}
