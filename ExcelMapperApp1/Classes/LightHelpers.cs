using SpreadsheetLight;

namespace ExcelMapperApp1.Classes;
public static class LightHelpers
{
    /// <summary>
    /// Get sheet names in an Excel file
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public static List<string> SheetNames(string fileName)
    {
        using SLDocument document = new(fileName);
        return document.GetSheetNames(false);
    }

    /// <summary>
    /// Determine if a sheet exists in the specified excel file
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="pSheetName"></param>
    /// <returns></returns>
    public static bool SheetExists(string fileName, string pSheetName)
    {

        using SLDocument document = new(fileName);
        return document.GetSheetNames(false).Any((sheetName) =>
            string.Equals(sheetName.ToLower(), pSheetName.ToLower(), StringComparison.OrdinalIgnoreCase));
    }

}
