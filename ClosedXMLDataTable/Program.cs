
using ClosedXMLDataTable.Classes;

namespace ClosedXMLDataTable;

partial class Program
{
    static void Main(string[] args)
    {
        ExcelOperations.WriteToCell("DemoSetCellValue.xlsx", 1,1,"Hello");
    }

    private static void Create()
    {
        ExcelOperations
            .Create(
                Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "NewFile.xlsx"));
    }
}