using System.Diagnostics;

namespace OleDbDemoForm.Classes
{
    public class Example
    {
        public static void Run(string fileName)
        {
            var connectionBuilder = new ExcelHelperBuilder()
                .UsingFileName(fileName)
                .HasHeader()
                .WithIMEX(1).Build();
            
        }
    }
}