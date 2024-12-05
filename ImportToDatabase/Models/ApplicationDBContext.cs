using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace ImportToDatabase.Models
{
    public class ApplicationDBContext : DbContext
    {
        //public DbSet<ProductCatalog> ProductCatalogs { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=DESKTOP-EB9K6GQ;Database=aspnet-BrokerWareLite3;Trusted_Connection=True;MultipleActiveResultSets=true");
        }
    }
}
