using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OleDbDemoForm.Classes
{
    public static class OleDbHelpers
    {
        /// <summary>
        /// Create connection string by file extension for Excel for worksheets with headers
        /// </summary>
        /// <param name="FileName">Excel file to create connection string for</param>
        /// <returns>connection string</returns>
        public static string ConnectionString(string FileName)
        {
            OleDbConnectionStringBuilder Builder = new();
            if (System.IO.Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", "Excel 8.0;IMEX=1;HDR=Yes;");
            }
            else
            {
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", "Excel 12.0;IMEX=1;HDR=Yes;");
            }

            Builder.DataSource = FileName;

            return Builder.ToString();

        }
    }
}
