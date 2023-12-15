using FastMember;
using System.Data;

namespace SpreadSheetLightImportDataTable.LanguageExtensions;

public static class Extensions
{
    public static DataTable ToDataTable<T>(this IEnumerable<T> sender, bool allowDbNull = false)
    {
        DataTable table = new(typeof(T).Name);
        using var reader = ObjectReader.Create(sender);
        table.Load(reader);

        if (!allowDbNull) return table;
        foreach (DataColumn column in table.Columns)
        {
            column.AllowDBNull = true;
        }

        return table;
    }
}

