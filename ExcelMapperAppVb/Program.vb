Imports ConsoleHelperLibrary.Classes
Imports Ganss.Excel
Imports Spectre.Console

Module Program
    Sub Main()
        Console.Title = "Code sample"
        WindowUtility.SetConsoleWindowPosition(WindowUtility.AnchorWindow.Center)

        Const excelFile = "Products.xlsx"

        Dim excel As New ExcelMapper()
        Dim products As List(Of Products) = excel.
                Fetch(Of Products)(excelFile, "Products").
                ToList()

        AnsiConsole.MarkupLine("[cyan]Done[/]")
        Console.ReadLine()
    End Sub
End Module
