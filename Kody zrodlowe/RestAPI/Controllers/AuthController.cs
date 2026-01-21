using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using RestAPI_zadania.Data;
using RestAPI_zadania.Models;
using RestAPI_zadania.Security;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

[ApiController]
[Route("api/[controller]")] // adres: api/auth
public class AuthController : ControllerBase
{
   private readonly BazaDbContext _context;
   public AuthController(BazaDbContext context)
   {
      _context = context;
   }

   [HttpPost("login")] // Adres: api/auth/login
   public IActionResult Login([FromBody] UserLogin login)
   {
      var user = _context.Uzytkownicy.FirstOrDefault(u => u.Username == login.Username);

      if (user == null || !BCrypt.Net.BCrypt.Verify(login.Password, user.PasswordHash))
      {
         return Unauthorized("Błędny użytkownik lub hasło.");
      }

      // dane są OK, generujemy token
      var token = GenerateJwtToken(user);
      return Ok(new { token });
   }

   [HttpPost("register")]
   [AllowAnonymous]
   public IActionResult Register([FromBody] UserLogin model)
   {
      // 1. Sprawdzamy, czy w bazie są już jacykolwiek użytkownicy
      bool czyBazaPusta = !_context.Uzytkownicy.Any();

      if (!czyBazaPusta)
      {
         // sprawdzamy bezpiecznie: czy Identity nie jest nullem ORAZ czy jest uwierzytelniony
         bool isAuthenticated = User?.Identity?.IsAuthenticated ?? false;
         bool isAdmin = User?.IsInRole("Admin") ?? false;

         if (!isAuthenticated || !isAdmin)
         {
            return Unauthorized("Tylko administrator może dodawać kolejnych użytkowników.");
         }
      }

      // sprawdzamy, czy login nie jest już zajęty
      if (_context.Uzytkownicy.Any(u => u.Username == model.Username))
      {
         return BadRequest("Użytkownik o takiej nazwie już istnieje.");
      }

      // tworzymy nowego użytkownika
      var newUser = new Uzytkownik
      {
         Username = model.Username,
         PasswordHash = BCrypt.Net.BCrypt.HashPassword(model.Password),
         Rola = czyBazaPusta ? "Admin" : "User"
      };

      _context.Uzytkownicy.Add(newUser);
      _context.SaveChanges();

      string komunikat = czyBazaPusta
          ? "Baza była pusta. Zarejestrowano pierwszego Administratora."
          : "Zarejestrowano nowego użytkownika.";

      return Ok(new { message = komunikat });
   }

   // generator tokena
   private string GenerateJwtToken(Uzytkownik user)
   {
      var tokenHandler = new JwtSecurityTokenHandler();
      var key = Encoding.ASCII.GetBytes(Const.SecurityKey);

      var tokenDescriptor = new SecurityTokenDescriptor
      {
         Subject = new ClaimsIdentity(new[]
          {
            new Claim(ClaimTypes.Name, user.Username),
            new Claim(ClaimTypes.Role, user.Rola)
        }),
         Expires = DateTime.UtcNow.AddDays(7),
         SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
      };

      var token = tokenHandler.CreateToken(tokenDescriptor);
      return tokenHandler.WriteToken(token);
   }
}