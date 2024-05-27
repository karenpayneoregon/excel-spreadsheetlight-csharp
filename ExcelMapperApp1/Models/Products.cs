#nullable disable

namespace ExcelMapperApp1.Models;

public class Products
{
    public int ProductID { get; set; }

    public string ProductName { get; set; }

    public string CategoryName { get; set; }
    public int? SupplierID { get; set; }

    public int? CategoryID { get; set; }

    public string Supplier { get; set; }
    public string QuantityPerUnit { get; set; }

    public decimal? UnitPrice { get; set; }

    public short? UnitsInStock { get; set; }

    public short? UnitsOnOrder { get; set; }

    public short? ReorderLevel { get; set; }

    
}