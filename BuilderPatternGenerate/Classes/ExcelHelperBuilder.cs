using System.Data.OleDb;
using BuilderPatternGenerate.Models;

namespace BuilderPatternGenerate.Classes;

public class ExcelHelperBuilder
{
    private string _fileName = "";
    private int _iMEX = 1;
    private bool _hasHeader = true;

    public ExcelHelper Build() =>
        new()
        {
            FileName = _fileName,
            IMEX = _iMEX,
            HasHeader = _hasHeader,
            ConnectionString = InternalConnectionString()
        };

    /// <summary>
    /// Excel file to work with
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExcelHelperBuilder WithFileName(string fileName)
    {
        _fileName = fileName;
        return this;
    }

    /// <summary>
    /// <code>
    /// 0 = Export mode
    /// 1 = intermix numbers
    /// 2 = Linked mode (full update capabilities)
    /// 3
    /// </code>
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public ExcelHelperBuilder WithIMEX(int value)
    {
        _iMEX = value;
        return this;
    }
    /// <summary>
    /// Set to true if first row contains column names
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public ExcelHelperBuilder WithHasHeader(bool value)
    {
        _hasHeader = value;
        return this;
    }

    public string InternalConnectionString()
    {
        if (string.IsNullOrWhiteSpace(_fileName))
        {
            throw new Exception("Must include a file name");
        }

        var header = _hasHeader ? "Yes" : "No";

        OleDbConnectionStringBuilder builder = new();

        if (Path.GetExtension(_fileName)!.ToUpper() == ".XLS")
        {
            builder.Provider = "Microsoft.Jet.OLEDB.4.0";
            builder.Add($"Extended Properties", $"Excel 8.0;IMEX={_iMEX};HDR={header};");
        }
        else
        {
            builder.Provider = "Microsoft.ACE.OLEDB.12.0";
            builder.Add("Extended Properties", $"Excel 12.0;IMEX={_iMEX};HDR={header};");
        }

        builder.DataSource = _fileName!;

        return builder.ConnectionString;
    }
}