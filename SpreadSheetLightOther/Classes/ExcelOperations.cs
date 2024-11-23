using SpreadsheetLight;

#pragma warning disable CS8619

namespace SpreadSheetLightOther.Classes;
internal class ExcelOperations
{
    /// <summary>
    /// Inserts a list of string data into a specified column in an Excel file starting from a given row.
    /// </summary>
    /// <param name="fileName">The name of the Excel file to be modified.</param>
    /// <param name="row">The starting row number where the data will be inserted.</param>
    /// <param name="column">The column number where the data will be inserted.</param>
    /// <param name="list">The list of string data to be inserted into the column.</param>
    /// <returns>A tuple containing a boolean indicating success or failure, and an exception if an error occurred.</returns>
    public static (bool success, Exception exception) InsertColumnData(string fileName, int row, int column, List<string> list)
    {
        try
        {
            using var document = new SLDocument(fileName, "Sheet1");
            for (int index = 0; index < list.Count; index++)
            {
                document.SetCellValue(SLConvert.ToCellReference(row, column), list[index]);

                row++;
            }

            document.Save();

            return (true, null);
        }
        catch (Exception exception)
        {
            return (false, exception);
        }
    }
}

