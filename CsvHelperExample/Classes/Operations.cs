using System.Data;
using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelperExample.Models;
using FastMember;

namespace CsvHelperExample.Classes;

/// <summary>
/// Provides an two examples for reading data known to be clean and one
/// example where data may be malformed
/// </summary>
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

                table.Columns["DiscontinuedDate"]!.SetOrdinal(4);
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
            HasHeaderRecord = false,
        };

        using (FileStream fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
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
    /// <summary>
    /// How to handle malformed lines
    /// </summary>
    public static (bool success, List<Account>) ReadAccounts1(string fileName)
    {
        List<Account> accounts = new();

        StringBuilder errorBuilder = new ();
            
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Encoding = Encoding.UTF8,
            Delimiter = ",",
            HasHeaderRecord = false,
        };

        using (var fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            using (var textReader = new StreamReader(fs, Encoding.UTF8))
            using (var csv = new CsvReader(textReader, configuration))
            {

                while (csv.Read())
                {
                    try
                    {
                        var record = csv.GetRecord<Account>();
                        accounts.Add(record);
                    }
                    catch (Exception ex)
                    {
                        errorBuilder.AppendLine(ex.Message);
                    }
                }
            }

            if (errorBuilder.Length >0)
            {
                errorBuilder.Insert(0, $"Errors for {fileName}\n");

                File.WriteAllText(
                    Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,"ParseErrors.txt"), 
                    errorBuilder.ToString());

                return (false, null);
            }
            else
            {
                return (true, accounts);
            }
        }
    }



}