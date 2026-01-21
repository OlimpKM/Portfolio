using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using RestAPI_zadania.Data;
using RestAPI_zadania.Security;

var builder = WebApplication.CreateBuilder(args);

// konfiguracja klucza i autentykacji
var key = Encoding.ASCII.GetBytes(Const.SecurityKey);

// dodanie serwisu
builder.Services.AddControllers();

builder.Services.AddDbContext<BazaDbContext>(options => options.UseSqlite("Data Source=zadania.db"));

builder.Services.AddCors(options => {
   options.AddDefaultPolicy(policy => {
      policy.SetIsOriginAllowed(_ => true) // ka¿dy adres (nawet null)
            .AllowAnyHeader()
            .AllowAnyMethod();
   });
});

builder.Services.AddAuthentication(options =>
{
   options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
   options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
   options.TokenValidationParameters = new TokenValidationParameters
   {
      ValidateIssuerSigningKey = true,
      IssuerSigningKey = new SymmetricSecurityKey(key),
      ValidateIssuer = false,
      ValidateAudience = false
   };
});

builder.Services.AddAuthorization();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
   options.SwaggerDoc("v1", new() { Title = "Moje API Zadañ", Version = "v1" });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
if (app.Environment.IsDevelopment())
{
   app.UseSwagger();
   app.UseSwaggerUI(options =>
   {
      options.SwaggerEndpoint("/swagger/v1/swagger.json", "v1");
      options.RoutePrefix = string.Empty; // przekierowanie na Swagger(a)
   });
}
app.UseRouting();
app.UseCors();

// w³¹czenie autoryzacji w potoku (Pipeline)
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();
app.MapStaticAssets();


app.Run();
