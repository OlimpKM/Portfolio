using Microsoft.EntityFrameworkCore;
using RestAPI_zadania.Models;

namespace RestAPI_zadania.Data
{
   public class BazaDbContext : DbContext
   {
      public BazaDbContext(DbContextOptions<BazaDbContext> options) : base(options) { }

      // To stworzy tabelę "Zadania" w bazie danych
      public DbSet<Zadanie> Zadania { get; set; }
      public DbSet<Uzytkownik> Uzytkownicy { get; set; }
   }
}
