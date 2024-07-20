Imports ConsoleHelperLibrary.Classes
Imports Ganss.Excel
Imports Spectre.Console

Module Program
    Sub Main()
        Console.Title = "Code sample"
        WindowUtility.SetConsoleWindowPosition(WindowUtility.AnchorWindow.Center)

        Const excelFile = "Products.xlsx"

        Dim excel As New ExcelMapper()

        ' Fetch data from the Excel file, and convert it to a list of Products
        ' From here use the data as you wish.
        ' You can set a breakpoint on Console.ReadLine() to inspect the data using the local window
        Dim products As List(Of Products) = excel.
                Fetch(Of Products)(excelFile, "Products").
                ToList()

        AnsiConsole.MarkupLine("[cyan]Done[/]")
        Console.ReadLine()
    End Sub
End Module
