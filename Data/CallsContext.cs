using CallStatistic.Models;
using Microsoft.EntityFrameworkCore;

namespace CallStatistic.Data
{
    public class CallsContext: DbContext
    {
        public CallsContext(DbContextOptions<CallsContext> options) : base(options) { }
        public DbSet<Calls> Calls { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Calls>().ToTable("Calls");
        }
    }
}