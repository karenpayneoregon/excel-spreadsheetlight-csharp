using SpreadsheetLight;

#pragma warning disable CS8619

namespace SpreadSheetLightOther.Classes;
internal class ExcelOperations
{
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

