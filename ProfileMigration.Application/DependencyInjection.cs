using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using ProfileMigration.Application.Analysis;
using ProfileMigration.Application.Areas;
using ProfileMigration.Application.Branches;
using ProfileMigration.Application.Constants;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Options;
using ProfileMigration.Application.Profiles;
using ProfileMigration.Application.Profiles.Commands;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application;

public static class DependencyInjection
{
    public static IServiceCollection AddProfileMigration(
        this IServiceCollection services,
        IConfiguration configuration,
        string? contentRoot = null)
    {
        services.Configure<MigrationOptions>(opts =>
        {
            opts.ConnectionString = configuration.GetConnectionString("OracleDb") ?? "";
            opts.DatabaseSchema = configuration["DatabaseSchema"] ?? opts.DatabaseSchema;
            opts.ExcelFilePath = configuration["ExcelFilePath"] ?? opts.ExcelFilePath;
            opts.IdCardExcelFilePath = configuration["IdCardExcelFilePath"] ?? opts.IdCardExcelFilePath;
            opts.AddressExcelFilePath = configuration["AddressExcelFilePath"] ?? opts.AddressExcelFilePath;
            opts.AreaExcelFilePath = configuration["AreaExcelFilePath"] ?? opts.AreaExcelFilePath;
            opts.ContentRoot = contentRoot ?? AppContext.BaseDirectory;

            configuration.GetSection(MigrationOptions.SectionName).Bind(opts);
            if (!string.IsNullOrWhiteSpace(configuration.GetConnectionString("OracleDb")))
                opts.ConnectionString = configuration.GetConnectionString("OracleDb")!;
        });

        services.AddSingleton<IOracleConnectionFactory, OracleConnectionFactory>();
        services.AddSingleton<MigrationPathResolver>();

        services.AddDbContext<SilaDbContext>((sp, b) =>
        {
            var opts = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<MigrationOptions>>().Value;
            SilaDbContext.DefaultSchema = opts.DatabaseSchema;
            b.UseOracle(opts.ConnectionString);
        });

        // Also register options for services that take DbContextOptions directly
        services.AddSingleton(sp =>
        {
            var opts = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<MigrationOptions>>().Value;
            SilaDbContext.DefaultSchema = opts.DatabaseSchema;
            return new DbContextOptionsBuilder<SilaDbContext>()
                .UseOracle(opts.ConnectionString)
                .Options;
        });

        services.AddScoped<BranchMigrationService>();
        services.AddScoped<ConstantMigrationService>();
        services.AddScoped<AreaMigrationService>();
        services.AddScoped<ConstantsPhaseService>();
        services.AddScoped<ProfileMigrationService>();
        services.AddScoped<WorkMigrationService>();
        services.AddScoped<BankMigrationService>();
        services.AddScoped<PartnerMigrationService>();
        services.AddScoped<ContactMigrationService>();
        services.AddScoped<ResetProfileMigrationDataCommandHandler>();
        services.AddScoped<ClientAnalysisService>();

        return services;
    }
}
