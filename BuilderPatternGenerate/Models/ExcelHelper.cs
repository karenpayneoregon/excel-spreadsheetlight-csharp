using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuilderPatternGenerate.Models
{
    public class ExcelHelper
    {
        public string FileName { get; set; }
        public int IMEX { get; set; }
        public bool HasHeader { get; set; }
        /// <summary>
        /// Connection string for connecting to an Excel file
        /// </summary>
        public string ConnectionString { get; set; }
    }

    public enum ExcelType
    {
        XLS,
        XLSX
    }
}
