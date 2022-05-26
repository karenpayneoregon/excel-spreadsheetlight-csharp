using System;

namespace CsvHelperExample.Models
{
    public class Products
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; }
        public decimal UnitPrice { get; set; }
        public short UnitsInStock { get; set; }
        public DateTime DiscontinuedDate { get; set; }
        public override string ToString() => ProductName;

    }
}