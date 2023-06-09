using CallStatistic.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddMvc().AddSessionStateTempDataProvider();
builder.Services.AddSession();
builder.Services.AddDbContext<CallsContext>(options => options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));
builder.Services.AddMvc();

var app = builder.Build();

//app.MapGet("/", () => "Hello World!");
app.UseRouting();
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Calls}/{action=Index}");
app.Run();
