using System.Data.Entity;

namespace _4335Project
{
    public class AppDbContext : DbContext
    {
        public DbSet<Service> Services { get; set; }
    }
}
