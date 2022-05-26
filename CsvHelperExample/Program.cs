using System;
using CsvHelperExample.Classes;

namespace CsvHelperExample
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var table = Operations.ReadAccounts();
            // us the DataTable Visualizer to view rows or iterate rows
            Console.ReadLine();
        }
    }
}
