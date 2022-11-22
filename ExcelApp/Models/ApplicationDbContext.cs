using Microsoft.EntityFrameworkCore;

namespace ExcelApp.Models
{
    public class ApplicationDbContext:DbContext
    {
        public ApplicationDbContext(DbContextOptions options): base(options)
        {

        }
        public DbSet<ExcelModel> Excel { get; set; }
    }
}
