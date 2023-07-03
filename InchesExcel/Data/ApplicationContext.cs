
using InchesExcel.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;


namespace InchesExcel.Data
{
    public class ApplicationContext : DbContext
    {
        public ApplicationContext(DbContextOptions<ApplicationContext> options) : base(options)
        {
        }
        public DbSet<ClientExcel> ClientExcel { get; set; }

        public DbSet<Doctors> Doctors { get; set; }

        //public IQueryable<ClientExcel> SearchBetweenDates(DateTime start, DateTime end)
        //{
        //    SqlParameter Start = new SqlParameter("@start", start);
        //    SqlParameter End = new SqlParameter("@end", end);
        //    return this.ClientExcel.FromSql("EXECUTE Customers_SearchCustomers @Start @End", Start,End);
        //}


    }
}
