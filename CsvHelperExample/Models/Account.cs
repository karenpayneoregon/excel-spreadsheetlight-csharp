using System;

namespace CsvHelperExample.Models
{
    public class Account
    {
        public string Column1 { get; set; }
        public long Column2 { get; set; }
        public DateTime Column3 { get; set; }
        public string Column4 { get; set; }
        public string Column5 { get; set; }

        public override string ToString() => Column1;

    }
}
