using Microsoft.AspNetCore.Http.Features;
using ProfileMigration.Application;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddRequestTimeouts();
// Excel files are CopyToOutputDirectory → resolve relative paths from BaseDirectory
builder.Services.AddProfileMigration(builder.Configuration, AppContext.BaseDirectory);

builder.Services.Configure<FormOptions>(o =>
{
    o.MultipartBodyLengthLimit = 200_000_000; // ~200 MB Excel uploads
});

builder.WebHost.ConfigureKestrel(o =>
{
    o.Limits.KeepAliveTimeout = TimeSpan.FromHours(2);
    o.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(5);
    o.Limits.MaxRequestBodySize = 200_000_000;
});

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();
app.UseRequestTimeouts();
app.MapControllers();

app.Run();
