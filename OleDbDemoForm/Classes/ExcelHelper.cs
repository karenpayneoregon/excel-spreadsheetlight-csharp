﻿namespace OleDbDemoForm.Classes
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


}
