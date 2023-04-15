using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using static Automation.Classes.ExcelBaseExample;
using Excel = Microsoft.Office.Interop.Excel;


namespace Automation.Classes
{
    public class ExcelOpenOperation : ExcelBase
    {
        public bool HasErrors { get; set; }
        public Dictionary<string, object> ReturnDictionary;
        public ExceptionInformation ExceptionInfo;

        public void ReadCells(string pFileName, string pSheetName)
        {
            // local method
            void ReleaseComObject(object pComObject)
            {
                try
                {
                    Marshal.ReleaseComObject(pComObject);
                    pComObject = null;
                }
                catch (Exception)
                {
                    pComObject = null;
                }
            }

            var annihilationList = new List<object>();
            var proceed = false;

            Excel.Application xlApp = null;
            Excel.Workbooks xlWorkBooks = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Sheets xlWorkSheets = null;
            Excel.Range xlCells = null;

            xlApp = new Excel.Application();
            annihilationList.Add(xlApp);

            xlApp.DisplayAlerts = false;

            xlWorkBooks = xlApp.Workbooks;
            annihilationList.Add(xlWorkBooks);

            xlWorkBook = xlWorkBooks.Open(pFileName);
            annihilationList.Add(xlWorkBook);

            xlApp.Visible = false;

            xlWorkSheets = xlWorkBook.Sheets;
            annihilationList.Add(xlWorkSheets);

            for (var sheetIndex = 1; sheetIndex <= xlWorkSheets.Count; sheetIndex++)
            {
                try
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkSheets[sheetIndex];

                    if (xlWorkSheet.Name == pSheetName)
                    {
                        proceed = true;
                        break;
                    }
                    else
                    {
                        ReleaseComObject(xlWorkSheet);
                    }
                }
                catch (Exception ex)
                {

                    HasErrors = true;
                    ExceptionInfo.UnKnownException = true;
                    ExceptionInfo.Message = $"Error finding sheet: '{ex.Message}'";
                    ExceptionInfo.FileNotFound = false;
                    ExceptionInfo.SheetNotFound = false;

                    proceed = false;
                    annihilationList.Add(xlWorkSheet);
                }
            }


            if (!proceed)
            {
                var firstSheet = (Excel.Worksheet)xlWorkSheets[1];
                xlWorkSheet = xlWorkSheets.Add(firstSheet);
                xlWorkSheet.Name = pSheetName;

                annihilationList.Add(firstSheet);
                annihilationList.Add(xlWorkSheet);

                xlWorkSheet.Name = pSheetName;

                proceed = true;
                ExceptionInfo.CreatedSheet = true;

            }
            else
            {
                if (!annihilationList.Contains(xlWorkSheet))
                {
                    annihilationList.Add(xlWorkSheet);
                }
            }

            if (proceed)
            {

                if (!annihilationList.Contains(xlWorkSheet))
                {
                    annihilationList.Add(xlWorkSheet);
                }


                foreach (var key in ReturnDictionary.Keys.ToArray())
                {
                    try
                    {
                        xlCells = xlWorkSheet.Range[key];
                        ReturnDictionary[key] = xlCells.Value;
                        annihilationList.Add(xlCells);
                    }
                    catch (Exception e)
                    {
                        HasErrors = true;
                        ExceptionInfo.Message = $"Error reading cell [{key}]: '{e.Message}'";
                        ExceptionInfo.FileNotFound = false;
                        ExceptionInfo.SheetNotFound = false;

                        annihilationList.Add(xlCells);

                        xlWorkBook.Close();
                        xlApp.UserControl = true;
                        xlApp.Quit();

                        annihilationList.Add(xlCells);

                        return;

                    }
                }
            }


            // this is debatable, should we save the file after adding a non-existing sheet?
            if (ExceptionInfo.CreatedSheet)
            {
                xlWorkSheet?.SaveAs(pFileName);
            }

            xlWorkBook.Close();
            xlApp.UserControl = true;
            xlApp.Quit();

            ReleaseObjects(annihilationList);

        }
    }
   
}
