
#nullable disable

namespace ExcelMapperApp1.Models;

public partial class Customers
{
    public int Id { get; set; }

    public string Company { get; set; }

    public string ContactType { get; set; }

    public string ContactName { get; set; }

    public string Country { get; set; }

    public DateOnly JoinDate { get; set; }
    public override string ToString() => Company;

}