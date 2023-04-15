# About

An example for using SpreadSheetLight with .NET Core which opens an existing Excel .xlsx file and writes to a single column then saves the file.

## Requires

NuGet package [SpreadsheetLight.Cross.Platform](https://www.nuget.org/packages/SpreadsheetLight.Cross.Platform/3.5.1?_src=template) which is for .NET Core 6 and 7.

For .NET Framework NuGet package [SpreadSheetLight](https://www.nuget.org/packages/SpreadsheetLight)

## Remarks

Seems that without looking at the source code for SpreadSheetLight that are issues saving files and/or updating cells and guess its from the version of DocumentFormat.OpenXml package.

All the other SpreadsheetLight projects work fine with the standard library.

