using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

using static OleDbDemoForm.Classes.OleDbHelpers;

namespace OleDbDemoForm.Classes
{
    public class Operations
    {

        public static DataTable GetPeopleFromExcel()
        {
            DataTable genderTable = new DataTable();
            DataTable personTable = new DataTable();

            string FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "People.xlsx");
            using OleDbConnection cn = new() { ConnectionString = ConnectionString(FileName) };
            using OleDbCommand cmd = new() { Connection = cn };
            cmd.CommandText = "SELECT GenderId, Role FROM [Gender$]";
            cn.Open();

            genderTable.Load(cmd.ExecuteReader());

            cmd.CommandText = "SELECT Id, FirstName, LastName, GenderId,  BirthDay  FROM [People$]";
            personTable.Load(cmd.ExecuteReader());

            personTable.Columns.Add("Gender", typeof(string));
            personTable.Columns["Gender"]!.SetOrdinal(4);


            // easy way to hide columns
            //personTable.Columns["Id"].ColumnMapping = MappingType.Hidden;
            //personTable.Columns["Gender"].ColumnMapping = MappingType.Hidden;

            /*
             * In many cases a developer gets column types wrong, this will provide
             * what .NET thinks it is vs what you think it is.
             */
            //foreach (DataColumn column in personTable.Columns)
            //{
            //    Debug.WriteLine($"{column.ColumnName}  {column.DataType}");
            //}

            /*
             * Cheap way to get gender 
             */
            foreach (DataRow row in personTable.Rows)
            {
                row.SetField("Gender", 
                    genderTable
                        .AsEnumerable()
                        .FirstOrDefault(dataRow => dataRow.Field<double>("GenderId") == 
                                                   row.Field<double>("GenderId"))!
                        .Field<string>("Role"));
            }

            // want to sort?
            // personTable.DefaultView.Sort = "LastName ASC";

            return personTable;

        }

	}
}
