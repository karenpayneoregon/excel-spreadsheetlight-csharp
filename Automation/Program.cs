using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Automation.Classes;

namespace Automation
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Code sample - Excel automation";

            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "People.xlsx");
            var sheetName = "People";

            ExcelBaseExample example = new ExcelBaseExample
            {
                ReturnDictionary = new Dictionary<string, object>()
                {
                    { "A1", null },
                    { "B2", null },
                    { "B3", null },
                    { "B4", null },
                    { "C2", null },
                    { "C4", null },
                    { "C5", null }
                }
            };

            example.ReadCells(fileName,sheetName);

            var data = example.ReturnDictionary;

            foreach (var kvp in data)
            {
                Console.WriteLine($"{kvp.Key,-10}{kvp.Value}");
            }

            Console.ReadLine();

        }
 
    }
}
