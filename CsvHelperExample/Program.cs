using System;
using System.IO;
using CsvHelperExample.Classes;

namespace CsvHelperExample
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var dt = Operations.ReadAccounts();
            //WellFormedData();
            //MalFormedData();
            Console.ReadLine();
        }
        private static void WellFormedData()
        {
            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Accounts.csv");
            var (success, accounts) = Operations.ReadAccounts1(fileName);
            Console.WriteLine(success ? "Do work" : "See error log");
            Console.ReadLine();
        }

        private static void MalFormedData()
        {
            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AccountsBad.csv");
            var (success, accounts) = Operations.ReadAccounts1(fileName);
            Console.WriteLine(success ? "Do work" : "See error log");
            Console.ReadLine();
        }
    }
}
