using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace SpreadSheetLightImportDataTable.LanguageExtensions
{
    /// <summary>
    /// Common extensions 
    /// </summary>
    public static class SheetExtensions
    {

        /// <summary>
        /// Same as in SheetHelpers while in this case it's an extension method
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool SheetExists(this SLDocument document, string sheetName) =>
            document.GetSheetNames(false).Any((name) =>
                string.Equals(name, sheetName, StringComparison.CurrentCultureIgnoreCase));

    }
}
