﻿// <auto-generated> This file has been auto generated by EF Core Power Tools. </auto-generated>

#nullable disable


using Microsoft.EntityFrameworkCore;
using NorthWind2020ConsoleApp.Models;
using NorthWind2020Library.Data.Configurations;
using NorthWind2020Library.Models;

namespace NorthWind2020Library.Data;

public partial class Context : DbContext
{
    public Context()
    {
    }

    public Context(DbContextOptions<Context> options)
        : base(options)
    {
    }

    public virtual DbSet<BusinessEntityPhone> BusinessEntityPhone { get; set; }
    public virtual DbSet<Categories> Categories { get; set; }
    public virtual DbSet<ContactDevices> ContactDevices { get; set; }
    public virtual DbSet<ContactType> ContactType { get; set; }
    public virtual DbSet<Contacts> Contacts { get; set; }
    public virtual DbSet<Countries> Countries { get; set; }
    public virtual DbSet<Customers> Customers { get; set; }
    public virtual DbSet<EmployeeTerritories> EmployeeTerritories { get; set; }
    public virtual DbSet<Employees> Employees { get; set; }
    public virtual DbSet<OrderDetails> OrderDetails { get; set; }
    public virtual DbSet<Orders> Orders { get; set; }
    public virtual DbSet<PhoneType> PhoneType { get; set; }
    public virtual DbSet<Products> Products { get; set; }
    public virtual DbSet<Region> Region { get; set; }
    public virtual DbSet<Shippers> Shippers { get; set; }
    public virtual DbSet<Suppliers> Suppliers { get; set; }
    public virtual DbSet<Territories> Territories { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        if (!optionsBuilder.IsConfigured)
        {
            // for this demo we are not concerned about security
            optionsBuilder.UseSqlServer("Data Source=.\\SQLEXPRESS;Initial Catalog=NorthWind2020;Integrated Security=True");
        }
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.HasAnnotation("Relational:Collation", "SQL_Latin1_General_CP1_CI_AS");

        modelBuilder.ApplyConfiguration(new BusinessEntityPhoneConfiguration());
        modelBuilder.ApplyConfiguration(new CategoriesConfiguration());
        modelBuilder.ApplyConfiguration(new ContactDevicesConfiguration());
        modelBuilder.ApplyConfiguration(new ContactTypeConfiguration());
        modelBuilder.ApplyConfiguration(new ContactsConfiguration());
        modelBuilder.ApplyConfiguration(new CountriesConfiguration());
        modelBuilder.ApplyConfiguration(new CustomersConfiguration());
        modelBuilder.ApplyConfiguration(new EmployeeTerritoriesConfiguration());
        modelBuilder.ApplyConfiguration(new EmployeesConfiguration());
        modelBuilder.ApplyConfiguration(new OrderDetailsConfiguration());
        modelBuilder.ApplyConfiguration(new OrdersConfiguration());
        modelBuilder.ApplyConfiguration(new PhoneTypeConfiguration());
        modelBuilder.ApplyConfiguration(new ProductsConfiguration());
        modelBuilder.ApplyConfiguration(new RegionConfiguration());
        modelBuilder.ApplyConfiguration(new ShippersConfiguration());
        modelBuilder.ApplyConfiguration(new SuppliersConfiguration());
        modelBuilder.ApplyConfiguration(new TerritoriesConfiguration());
        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}