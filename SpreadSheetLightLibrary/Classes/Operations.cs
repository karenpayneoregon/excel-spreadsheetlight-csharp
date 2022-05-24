
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using Color = System.Drawing.Color;

namespace SpreadSheetLightLibrary.Classes
{
    public class Operations
    {
        public static int GetWorkSheetLastRow(string pFileName, string pSheetName)
        {
            var lastRow = 0;

            using SLDocument document = new(pFileName, pSheetName);
            /*
             * get statistics, in this case we want the last used row so we
             * simply index into EndRowIndex yet there are more properties.
             */
            lastRow = document.GetWorksheetStatistics().EndRowIndex;

            return lastRow;
        }
        /// <summary>
        /// Create a new Excel file, rename the default sheet from
        /// Sheet1 to the value in pSheetName
        /// </summary>
        /// <param name="pFileName"></param>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public bool CreateNewFile(string pFileName, string pSheetName)
        {
            using SLDocument document = new();
            document.RenameWorksheet("Sheet1", pSheetName);
            document.SaveAs(pFileName);
            return true;
        }
        /// <summary>
        /// Create a new Excel file
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool CreateNewFile(string pFileName)
        {
            using SLDocument document = new();
            document.SaveAs(pFileName);
            return true;
        }
        /// <summary>
        /// Example for importing a tab delimited text file
        /// </summary>
        /// <param name="pTextFileName"></param>
        /// <param name="pExcelFileName"></param>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public static bool ImportTabDelimitedTextFile(string pTextFileName, string pExcelFileName, string pSheetName)
        {
            try
            {

                var line = System.IO.File.ReadAllLines(pExcelFileName).FirstOrDefault();
                /*
                 * Needed later for auto-fit columns
                 */
                var columnCount = line.Split('\t').Length;


                using SLDocument document = new();
                var headerStyle = HeaderStye(document);
                var sheets = document.GetSheetNames(false);
                if (sheets.Any((sheetName) => string.Equals(sheetName, pSheetName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    document.SelectWorksheet(pSheetName);
                    document.ClearCellContent();
                }
                else
                {
                    document.AddWorksheet(pSheetName);
                }

                var importOptions = new SLTextImportOptions();

                document.ImportText(pTextFileName, "A1", importOptions);

                // do not need Sheet1
                if (sheets.FirstOrDefault((sheetName) => sheetName.ToLower() == "sheet1") != null)
                {
                    if (pSheetName.ToLower() != "sheet1")
                    {
                        document.DeleteWorksheet("Sheet1");
                    }
                }

                document.SetCellStyle(1, 1, 1, columnCount, headerStyle);

                for (int columnIndex = 1; columnIndex < columnCount +1; columnIndex++)
                {
                    document.AutoFitColumn(columnIndex);
                }

                
                document.SetActiveCell("C2");

                // ensure header is visible when scrolling down
                document.FreezePanes(1, 6);

                document.SaveAs(pExcelFileName);

                return true;

            }
            catch (Exception exception)
            {
                return false;
            }
        }
        /// <summary>
        /// Example for formatting currency and dates
        /// 
        /// var ops = new SpreadSheetLightLibrary.Examples();
        /// var excelFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SpreadSheetLightFormatting.xlsx");
        /// if (File.Exists(excelFileName))
        /// {
        ///     File.Delete(excelFileName);
        /// }
        /// 
        /// ops.CreateNewFile(excelFileName);
        /// ops.SimpleFormatting(excelFileName);
        /// 
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool SimpleFormatting(string pFileName)
        {
            using SLDocument document = new(pFileName, "Sheet1");
            SLStyle currencyStyle = document.CreateStyle();
            currencyStyle.FormatCode = "$#,##0.000";

            document.SetCellValue("H3", 100.3);
            document.SetCellValue("I3", 200.5);
            document.SetCellStyle("H3", currencyStyle);
            document.SetCellStyle("I3", currencyStyle);

            SLStyle dateStyle = document.CreateStyle();
            dateStyle.FormatCode = "mm-dd-yyyy";


            Dictionary<string, DateTime> dictDates = new Dictionary<string, DateTime>()
            {
                {
                    "H4", new DateTime(2017,
                        1,
                        1)
                },
                {
                    "H5", new DateTime(2017,
                        1,
                        2)
                },
                {
                    "H6", new DateTime(2017,
                        1,
                        3)
                },
                {
                    "H7", new DateTime(2017,
                        1,
                        4)
                }
            };

            foreach (var dateItem in dictDates)
            {
                if (document.SetCellValue(dateItem.Key, dateItem.Value))
                {
                    document.SetCellStyle(dateItem.Key, dateStyle);
                    document.SetColumnWidth(dateItem.Key, 12);
                }

            }

            document.Save();

            return true;

        }
        /// <summary>
        /// demonstrate how to get used columns in the format a a letter rather than an integer
        /// </summary>
        /// <returns></returns>
        public string[] UsedCellsInWorkSheet(string pFileName, string pSheetName)
        {
            using SLDocument document = new(pFileName, pSheetName);
            SLWorksheetStatistics stats = document.GetWorksheetStatistics();

            IEnumerable<string> columnNames = Enumerable.Range(1, stats.EndColumnIndex)
                // ReSharper disable once ConvertClosureToMethodGroup
                .Select((cellIndex) => SLConvert.ToColumnName(cellIndex));

            return columnNames.ToArray();
        }

        /// <summary>
        /// Get sheet names in an Excel file
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public List<string> SheetNames(string pFileName)
        {
            using SLDocument document = new(pFileName);
            return document.GetSheetNames(false);
        }

        /// <summary>
        /// Remove a sheet if it exists.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public bool RemoveWorkSheet(string pFileName, string pSheetName)
        {
            using SLDocument document = new(pFileName);
            var workSheets = document.GetSheetNames(false);
            if (workSheets.Any((sheetName) => sheetName.ToLower() == pSheetName.ToLower()))
            {
                if (workSheets.Count > 1)
                {
                    document.SelectWorksheet(document.GetSheetNames().FirstOrDefault((sName) => sName.ToLower() != pSheetName.ToLower()));
                }
                else if (workSheets.Count == 1)
                {
                    throw new Exception("Can not delete the sole worksheet");
                }

                document.DeleteWorksheet(pSheetName);
                document.Save();

                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Add a new sheet if it does not currently exists.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public bool AddNewSheet(string pFileName, string pSheetName)
        {
            using SLDocument document = new(pFileName);
            if (!(document.GetSheetNames(false).Any((sheetName) => string.Equals(sheetName, pSheetName, StringComparison.CurrentCultureIgnoreCase))))
            {
                document.AddWorksheet(pSheetName);
                document.Save();
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Determine if a sheet exists in the specified excel file
        /// </summary>
        /// <param name="pFileName"></param>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public bool SheetExists(string pFileName, string pSheetName)
        {
            using SLDocument document = new(pFileName);
            return document.GetSheetNames(false).Any((sheetName) => sheetName.ToLower() == pSheetName.ToLower());
        }


        public static SLStyle HeaderStye(SLDocument document)
        {

            SLStyle headerStyle = document.CreateStyle();

            headerStyle.Font.Bold = true;
            headerStyle.Font.FontColor = Color.White;
            headerStyle.Fill.SetPattern(
                PatternValues.LightGray,
                SLThemeColorIndexValues.Accent1Color,
                SLThemeColorIndexValues.Accent5Color);

            return headerStyle;
        }
    }
}
