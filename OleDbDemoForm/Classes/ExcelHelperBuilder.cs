using System;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OleDbDemoForm.Classes
{
    public class ExcelHelperBuilder
    {
        private string _fileName = "";
        private int _iMEX = 1;
        private bool _hasHeader = true;
        
        public ExcelHelper Build() =>
          new ExcelHelper
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
        public ExcelHelperBuilder UsingFileName(string fileName)
        {
            _fileName = fileName;
            return this;
        }

        /// <summary>
        /// <code>
        /// 0 = Export mode
        /// 1 = intermix numbers
        /// 2 = Linked mode (full update capabilities)
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
        public ExcelHelperBuilder HasHeader(bool value = true)
        {
            _hasHeader = value;
            return this;
        }

        private string[] validExtensions => new [] { ".xlsx", ".xls" };

        public string InternalConnectionString()
        {
            if (string.IsNullOrWhiteSpace(_fileName))
            {
                throw new Exception("Must include a file name");
            }
            
            var knownFileExtension = validExtensions.Contains(Path.GetExtension(_fileName), 
                StringComparer.OrdinalIgnoreCase);

            if (!knownFileExtension)
            {
                throw new Exception("unknown file extensions");
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
}
