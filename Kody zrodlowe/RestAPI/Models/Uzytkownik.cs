namespace RestAPI_zadania.Models
{
   public class Uzytkownik
   {
      public int Id { get; set; }
      public string Username { get; set; } = string.Empty;
      public string PasswordHash { get; set; } = string.Empty; // Szyfrowane hasło
      public string Rola { get; set; } = "User";
   }
}
