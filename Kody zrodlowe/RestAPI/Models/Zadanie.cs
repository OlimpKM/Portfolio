namespace RestAPI_zadania.Models
{
    public class Zadanie
    {
      public int Id { get; set; }
      public string Tytul { get; set; } = string.Empty;
      public bool CzyWykonane { get; set; }
      public DateTime DataUtworzenia { get; set; } = DateTime.Now;
   }
}
