using System.Data.OleDb;

namespace OleDbDemoForm.Classes;

public static class OleDbHelpers
{
    /// <summary>
    /// Create connection string by file extension for Excel for worksheets with headers
    /// </summary>
    /// <param name="fileName">Excel file to create connection string for</param>
    /// <returns>connection string</returns>
    public static string ConnectionString(string fileName)
    {
        OleDbConnectionStringBuilder Builder = new();
            
        if (System.IO.Path.GetExtension(fileName).ToUpper() == ".XLS")
        {
            Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
            Builder.Add("Extended Properties", "Excel 8.0;IMEX=1;HDR=Yes;");
        }
        else
        {
            Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
            Builder.Add("Extended Properties", "Excel 12.0;IMEX=1;HDR=Yes;");
        }

        Builder.DataSource = fileName;

        return Builder.ToString();

    }
}