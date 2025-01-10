using System.Data.OleDb;
using BuilderPatternGenerate.Models;

namespace BuilderPatternGenerate.Classes;

/// <summary>
/// Provides a builder for creating and configuring instances of the <see cref="BuilderPatternGenerate.Models.ExcelHelper"/> class.
/// </summary>
/// <remarks>
/// This class simplifies the creation of <see cref="BuilderPatternGenerate.Models.ExcelHelper"/> objects by using the builder pattern.
/// It allows setting properties such as the file name, IMEX mode, and whether the first row contains headers.
/// </remarks>
public class ExcelHelperBuilder
{
    private string _fileName = "";
    private int _iMEX = 1;
    private bool _hasHeader = true;

    /// <summary>
    /// Builds and returns a new instance of the <see cref="ExcelHelper"/> class
    /// configured with the current settings of the builder.
    /// </summary>
    /// <returns>
    /// A new instance of the <see cref="ExcelHelper"/> class
    /// with properties such as <see cref="ExcelHelper.FileName"/>,
    /// <see cref="ExcelHelper.IMEX"/>, 
    /// <see cref="ExcelHelper.HasHeader"/>, 
    /// and <see cref="ExcelHelper.ConnectionString"/> set accordingly.
    /// </returns>
    /// <exception cref="Exception">
    /// Thrown if the file name is not specified or is invalid when generating the connection string.
    /// </exception>
    public ExcelHelper Build() =>
        new()
        {
            FileName = _fileName,
            IMEX = _iMEX,
            HasHeader = _hasHeader,
            ConnectionString = InternalConnectionString()
        };

    /// <summary>
    /// Specifies the file name of the Excel file to be used by the <see cref="ExcelHelper"/> instance.
    /// </summary>
    /// <param name="fileName">
    /// The full path to the Excel file. This value is used to configure the 
    /// <see cref="ExcelHelper.FileName"/> property and to generate the connection string.
    /// </param>
    /// <returns>
    /// The current instance of <see cref="ExcelHelperBuilder"/> to allow method chaining.
    /// </returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the provided <paramref name="fileName"/> is null, empty, or consists only of white-space characters.
    /// </exception>
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

    /// <summary>
    /// Generates the connection string required to access the Excel file
    /// based on the current settings of the builder.
    /// </summary>
    /// <returns>
    /// A connection string configured with the file name, IMEX mode, and header settings.
    /// </returns>
    /// <exception cref="Exception">
    /// Thrown if the file name is not specified or is invalid.
    /// </exception>
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