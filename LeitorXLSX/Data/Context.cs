using LeitorXLSX.Models;
using Microsoft.EntityFrameworkCore;

namespace LeitorXLSX.Data
{
    public class Context : DbContext
    {
        public Context(DbContextOptions<Context> options) : base(options)
        {

        }

        // Models;
        public DbSet<Voto>? Votos { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
    
        }
    }
}
