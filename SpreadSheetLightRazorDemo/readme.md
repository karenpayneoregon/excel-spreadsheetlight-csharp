# About

Shows how to export a DataTable using mocked data to an Excel file using the SpreadsheetLight.

## Notes

In a real application data would come from a List&lt;T> which needs to be convert to a DataTable first via.

```csharp
/// <summary>
/// Converts an <see cref="IEnumerable{T}"/> to a <see cref="DataTable"/>.
/// </summary>
/// <typeparam name="T">The type of the elements in the source collection.</typeparam>
/// <param name="sender">The source collection to convert.</param>
/// <returns>A <see cref="DataTable"/> representation of the source collection.</returns>
public static DataTable ToDataTable<T>(this IEnumerable<T> sender)
{
    DataTable table = new(typeof(T).Name);
    using var reader = ObjectReader.Create(sender);
    table.Load(reader);
    return table;
}
```