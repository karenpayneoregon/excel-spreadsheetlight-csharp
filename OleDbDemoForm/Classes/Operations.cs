using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OleDbDemoForm.Classes
{
    public class Operations
    {
        public static string ConnectionString(string FileName)
        {
            OleDbConnectionStringBuilder Builder = new();
            if (System.IO.Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", "Excel 8.0;IMEX=1;HDR=No;");
            }
            else
            {
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", "Excel 12.0;IMEX=1;HDR=Yes;");
            }

            Builder.DataSource = FileName;

            return Builder.ToString();

        }
        public static DataTable GetData()
        {
            DataTable genderTable = new DataTable();
            DataTable personTable = new DataTable();

            string FileName = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "People.xlsx");
            using OleDbConnection cn = new() { ConnectionString = ConnectionString(FileName) };
            using OleDbCommand cmd = new() { Connection = cn };
            cmd.CommandText = "SELECT GenderId, Role FROM [Gender$]";
            cn.Open();

            genderTable.Load(cmd.ExecuteReader());

            cmd.CommandText = "SELECT Id, FirstName, LastName, GenderId,  BirthDay  FROM [People$]";
            personTable.Load(cmd.ExecuteReader());

            personTable.Columns.Add("Gender", typeof(string));
            personTable.Columns["Gender"].SetOrdinal(4);

            //personTable.Columns["Id"].ColumnMapping = MappingType.Hidden;
            //personTable.Columns["Gender"].ColumnMapping = MappingType.Hidden;

            foreach (DataColumn column in personTable.Columns)
            {
                Debug.WriteLine($"{column.ColumnName}  {column.DataType}");
            }


            foreach (DataRow row in personTable.Rows)
            {
                row.SetField("Gender", 
                    genderTable
                        .AsEnumerable()
                        .FirstOrDefault(dataRow => dataRow.Field<double>("GenderId") == 
                                                   row.Field<double>("GenderId"))!
                        .Field<string>("Role"));
            }

            return personTable;
        }

	}
}
