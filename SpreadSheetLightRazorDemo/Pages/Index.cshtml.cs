using Microsoft.AspNetCore.Mvc.RazorPages;
using Serilog;
using SpreadSheetLightImportDataTable.Classes;

namespace SpreadSheetLightRazorDemo.Pages;
public class IndexModel : PageModel
{
    public string? Message { get; set; }
    public void OnGet() { }
    public void OnPost()
    {
        var (success, exception) = ExportUsingRazor.Export();
        Message = success ? "Done" : exception.Message;
    }
}
