using SpreadsheetLight;
using static ExcelMapperApp1.Classes.SpectreConsoleHelpers;

namespace ExcelMapperApp1.Classes;
internal class LightOperations
{
    /// <summary>
    /// For article to show an example to test if the person's birthdate can be read as a date
    /// </summary>
    /// <returns>
    /// If there are issues, the list of rows with issues is returned
    /// </returns>
    public static (List<int> rows, bool hasIssues) Iterate()
    {

        PrintCyan();

        List<int> list = [];

        const string excelFile = "Nested1.xlsx";
        const int columnIndex = 4;

        using SLDocument document = new(excelFile);

        var stats = document.GetWorksheetStatistics();

        // skip header row
        for (int rowIndex = 2; rowIndex < stats.EndRowIndex + 1; rowIndex++)
        {
            var date = document.GetCellValueAsDateTime(rowIndex, columnIndex);
            if (date == new DateTime(1900,1,1))
            {
                list.Add(rowIndex);
            }
        }

        return (list, list.Any());
    }
}
