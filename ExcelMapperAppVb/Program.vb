Imports Ganss.Excel

Module Program
    Sub Main()

        Const excelFile = "Products.xlsx"

        Dim excel As New ExcelMapper()
        Dim products As List(Of Products) = excel.Fetch(Of Products)(excelFile, "Products").ToList()


        Console.WriteLine("Done")
    End Sub
End Module
