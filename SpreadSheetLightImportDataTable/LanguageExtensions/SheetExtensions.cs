using SpreadsheetLight;

namespace SpreadSheetLightImportDataTable.LanguageExtensions;

/// <summary>
/// Common extensions 
/// </summary>
public static class SheetExtensions
{

    /// <summary>
    /// Checks if a sheet with the specified name exists in the given SLDocument.
    /// </summary>
    /// <param name="document">The SLDocument to check for the sheet.</param>
    /// <param name="sheetName">The name of the sheet to look for.</param>
    /// <returns>
    /// <c>true</c> if the sheet exists; otherwise, <c>false</c>.
    /// </returns>
    public static bool SheetExists(this SLDocument document, string sheetName) =>
        document.GetSheetNames(false).Any((name) =>
            string.Equals(name, sheetName, StringComparison.CurrentCultureIgnoreCase));

}