namespace ExcelMapperApp1.Models;
public class Person
{
    public int Id { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public DateOnly BirthDate { get; set; }
    public Address Address { get; set; }
    public override string ToString() => $"{FirstName} {LastName} {Address}";
}