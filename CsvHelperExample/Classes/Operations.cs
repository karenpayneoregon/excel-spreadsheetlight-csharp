using System.Data;
using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelperExample.Models;
using FastMember;

namespace CsvHelperExample.Classes;

/// <summary>
/// Provides two examples for reading data known to be clean and one
/// example where data may be malformed
/// </summary>
class Operations
{
    /// <summary>
    /// Reads product data from a CSV file and loads it into a <see cref="DataTable"/>.
    /// </summary>
    /// <remarks>
    /// The method reads data from a file named "Products.csv" located in the application's base directory.
    /// It uses the <see cref="CsvHelper"/> library to parse the CSV file and the <see cref="FastMember.ObjectReader"/> 
    /// to load the data into a <see cref="DataTable"/>. The "DiscontinuedDate" column is reordered to the 5th position.
    /// </remarks>
    /// <returns>
    /// A <see cref="DataTable"/> containing the product data read from the CSV file.
    /// </returns>
    /// <exception cref="FileNotFoundException">
    /// Thrown if the "Products.csv" file is not found in the application's base directory.
    /// </exception>
    /// <exception cref="IOException">
    /// Thrown if there is an issue accessing or reading the "Products.csv" file.
    /// </exception>
    /// <exception cref="CsvHelperException">
    /// Thrown if there is an error while parsing the CSV file.
    /// </exception>
    public static DataTable ReadProducts()
    {
        var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Products.csv");
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Encoding = Encoding.UTF8,
            Delimiter = "," ,
            HasHeaderRecord = false
        };

        using var fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
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

    /// <summary>
    /// Reads account data from a CSV file and loads it into a <see cref="DataTable"/>.
    /// </summary>
    /// <remarks>
    /// The method reads data from a file named "Accounts.csv" located in the application's base directory.
    /// It uses the <see cref="CsvHelper"/> library to parse the CSV file and the <see cref="FastMember.ObjectReader"/> 
    /// to load the data into a <see cref="DataTable"/>.
    /// </remarks>
    /// <returns>
    /// A <see cref="DataTable"/> containing the account data read from the CSV file.
    /// </returns>
    /// <exception cref="FileNotFoundException">
    /// Thrown if the "Accounts.csv" file is not found in the application's base directory.
    /// </exception>
    /// <exception cref="IOException">
    /// Thrown if there is an issue accessing or reading the "Accounts.csv" file.
    /// </exception>
    /// <exception cref="CsvHelperException">
    /// Thrown if there is an error while parsing the CSV file.
    /// </exception>
    public static DataTable ReadAccounts()
    {
         
        var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Accounts.csv");
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Encoding = Encoding.UTF8,
            Delimiter = ",",
            HasHeaderRecord = false,
        };

        using FileStream fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
        using (var textReader = new StreamReader(fs, Encoding.UTF8))
        using (var csv = new CsvReader(textReader, configuration))
        {
                    
            DataTable table = new();
            using var reader = ObjectReader.Create(csv.GetRecords<Account>().ToList());
            table.Load(reader);
            return table;
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

        using var fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
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