using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using RestAPI_zadania.Security;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using RestAPI_zadania.Data;
using RestAPI_zadania.Models;

namespace RestAPI_zadania.Controllers
{
   [Route("api/[controller]")] // adres: api/zadania
   [ApiController]
   [Authorize]

   public class ZadaniaController : ControllerBase
   {
      private readonly BazaDbContext _context;

      public ZadaniaController(BazaDbContext context)
      {
         _context = context;
      }

      // pobieranie WSZYSTKICH zadań
      // adres: GET api/zadania
      [HttpGet]
      public async Task<IActionResult> GetAll()
      {
         var zadania = await _context.Zadania.ToListAsync();
         return Ok(zadania);
      }

      // pobieranie jednego zadania po ID
      // adres: GET api/zadania/{id} (np. api/zadania/1)
      [HttpGet("{id}")]
      public async Task<IActionResult> GetById(int id)
      {
         // Szukamy zadania w bazie danych po jego kluczu głównym
         var zadanie = await _context.Zadania.FindAsync(id);

         // Jeśli nie znaleziono zadania o takim ID, zwracamy 404 Not Found
         if (zadanie == null)
         {
            return NotFound(new { message = $"Nie znaleziono zadania o ID {id}" });
         }

         return Ok(zadanie);
      }

      // dodanie zadania
      // adres: POST api/zadania
      [HttpPost]
      public async Task<ActionResult<Zadanie>> PostZadanie(Zadanie zadanie)
      {
         _context.Zadania.Add(zadanie);
         await _context.SaveChangesAsync();
         return CreatedAtAction(nameof(GetById), new { id = zadanie.Id }, zadanie);
      }

      // usunięcie zadania po ID
      // adres: DELETE api/zadania/{id}
      [HttpDelete("{id}")]
      public async Task<IActionResult> DelById(int id)
      {
         var zadanie = await _context.Zadania.FindAsync(id);
         if (zadanie == null)
         {
            return NotFound(new { message = $"Nie znaleziono zadania o ID {id}" });
         }

         _context.Zadania.Remove(zadanie);
         await _context.SaveChangesAsync();

         return Ok(zadanie);
      }

      // aktualizacja zadania
      // adres: PUT api/zadania/3
      [HttpPut("{id}")]
      public async Task<IActionResult> UpdateZadanie(int id, Zadanie zadanieZmienione)
      {
         if (id != zadanieZmienione.Id)
         {
            return BadRequest(new { message = "ID w adresie i w obiekcie nie są zgodne" });
         }
         _context.Entry(zadanieZmienione).State = EntityState.Modified;

         try
         {
            await _context.SaveChangesAsync();
         }
         catch (DbUpdateConcurrencyException)
         {
            if (!_context.Zadania.Any(e => e.Id == id))
            {
               return NotFound();
            }
            else { throw; }
         }

         return NoContent(); // Zwracamy status 204 (sukces, brak treści do odesłania)
      }
   }
}
