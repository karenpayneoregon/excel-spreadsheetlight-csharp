# About

Based off Stackoverflow question [Microsoft.ACE.OLEDB.12.0 ignores data in a column](https://stackoverflow.com/questions/78482822/microsoft-ace-oledb-12-0-ignores-data-in-a-column).

They want to get data from one column using OleDb which is more work than its worth and is fragile.

In this code [SpreadSheetLight](https://www.nuget.org/packages/SpreadsheetLight.Cross.Platform/3.5.1?_src=template) NuGet package is used to iterate columns in a single row to a DataTable, a List&lt;T> is an option too but in the question they wanted a DataTable.

There are a handful of other Excel NuGet packages that can perform what is done here with SpreadSheetLight but here is super simple using SpreadSheetLight.

1. Open the Excel file to a sheet, one line of code
1. Iterate rows to last used row in first column
1. Add data to DataTable

## Why did I not reply?

Doubtful they wanted anything but OleDb
