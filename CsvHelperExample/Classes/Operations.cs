using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelperExample.Models;
using FastMember;

namespace CsvHelperExample.Classes
{
    class Operations
    {
        public static DataTable ReadProducts()
        {
            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Products.csv");
            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Encoding = Encoding.UTF8,
                Delimiter = "," ,
                HasHeaderRecord = false
            };

            using (var fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var textReader = new StreamReader(fs, Encoding.UTF8))
                using (var csv = new CsvReader(textReader, configuration))
                {
                    DataTable table = new();
                    using var reader = ObjectReader.Create(csv.GetRecords<Products>().ToList());

                    table.Load(reader);

                    table.Columns["DiscontinuedDate"].SetOrdinal(4);
                    return table;
                }
            }
        }

        public static DataTable ReadAccounts()
        {
            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Accounts.csv");
            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Encoding = Encoding.UTF8,
                Delimiter = ",",
                HasHeaderRecord = false
            };

            using (var fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var textReader = new StreamReader(fs, Encoding.UTF8))
                using (var csv = new CsvReader(textReader, configuration))
                {
                    DataTable table = new();
                    using var reader = ObjectReader.Create(csv.GetRecords<Account>().ToList());

                    table.Load(reader);
                    return table;
                }
            }
        }
    }
}
