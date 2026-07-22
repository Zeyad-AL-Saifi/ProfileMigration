using Dapper;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application.Branches;

public sealed class BranchMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    ILogger<BranchMigrationService> logger)
{
    static readonly (int Id, string NameEn, bool IsCentral)[] SeedBranches =
    [
        (1, "Default", false), (2, "Tulkarm", false), (3, "Tubas", false),
        (4, "Nablus", false), (5, "Ramallah", true), (6, "Jericho", false),
        (7, "Bethlehem", false), (8, "Hebron", false), (9, "Gaza", false),
        (10, "Jenin", false),
    ];

    static readonly (int BranchId, string LangId, string Name)[] SeedLangs =
    [
        (1, "en", "Default"), (1, "ar", "Default"),
        (2, "en", "Tulkarm"), (2, "ar", "طولكرم"),
        (3, "en", "Tubas"), (3, "ar", "طوباس"),
        (4, "en", "Nablus"), (4, "ar", "نابلس"),
        (5, "en", "Ramallah"), (5, "ar", "رام الله"),
        (6, "en", "Jericho"), (6, "ar", "أريحا"),
        (7, "en", "Bethlehem"), (7, "ar", "بيت لحم"),
        (8, "en", "Hebron"), (8, "ar", "الخليل"),
        (9, "en", "Gaza"), (9, "ar", "غزة"),
        (10, "en", "Jenin"), (10, "ar", "جنين"),
    ];

    public async Task<MigrationReportDto> ValidateAsync(CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "branches", CanProceed = true };

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);

        int existingBranches = await conn.ExecuteScalarAsync<int>(
            $"SELECT COUNT(1) FROM {connectionFactory.QualifyTable("BRANCHES_TB")}");
        int existingLangs = await conn.ExecuteScalarAsync<int>(
            $"SELECT COUNT(1) FROM {connectionFactory.QualifyTable("BRANCH_LANGS_TB")}");

        report.Stats["expectedBranches"] = SeedBranches.Length;
        report.Stats["existingBranches"] = existingBranches;
        report.Stats["willDeleteBranches"] = existingBranches;
        report.Stats["willInsertBranches"] = SeedBranches.Length;
        report.Stats["expectedLangs"] = SeedLangs.Length;
        report.Stats["existingLangs"] = existingLangs;
        report.Stats["willDeleteLangs"] = existingLangs;
        report.Stats["willInsertLangs"] = SeedLangs.Length;

        ReportBuilder.AddIssue(
            report,
            "Info",
            "BRANCHES_REPLACE",
            "Existing branches and language rows will be replaced by the configured branches (1-10).",
            existingBranches + existingLangs,
            SeedBranches.Select(b => $"{b.Id}:{b.NameEn}"));

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(CancellationToken ct = default)
    {
        var log = new List<string>();
        const int CountryId = 1;
        var today = DateTime.Today;
        var now = DateTime.Now;

        await using var ctx = new SilaDbContext(dbOptions);
        await using var transaction = await ctx.Database.BeginTransactionAsync(ct);

        foreach (var (id, nameEn, isCentral) in SeedBranches)
        {
            ctx.BranchesTbs.Add(new BranchesTb
            {
                BranchId = id,
                BranchName = nameEn,
                CountryId = CountryId,
                StateId = id,
                CityId = id,
                IsCentralFlag = isCentral,
                CurrDate = today,
                DailyCloseFlag = false,
                CloseFlag = false,
                CreatedOn = now,
                IsHidden = false,
                GovernorateId = null,
            });
            log.Add($"[ADD] BRANCHES_TB BranchId={id} ({nameEn})");
        }

        foreach (var (branchId, langId, name) in SeedLangs)
        {
            ctx.BranchLangsTbs.Add(new BranchLangsTb
            {
                BranchId = branchId,
                LangId = langId,
                BranchName = name,
                CreatedOn = now,
            });
            log.Add($"[ADD] BRANCH_LANGS_TB ({branchId},{langId}) '{name}'");
        }

        try
        {
            int langsDeleted = await ctx.BranchLangsTbs.ExecuteDeleteAsync(ct);
            int branchesDeleted = await ctx.BranchesTbs.ExecuteDeleteAsync(ct);
            await ctx.SaveChangesAsync(ct);
            await transaction.CommitAsync(ct);

            logger.LogWarning(
                "Branches replaced. Deleted {BranchesDeleted} branches and {LangsDeleted} language rows; inserted {BranchesInserted} branches and {LangsInserted} language rows",
                branchesDeleted,
                langsDeleted,
                SeedBranches.Length,
                SeedLangs.Length);

            return new PhaseRunResultDto
            {
                Phase = "branches",
                Success = true,
                Message = $"Branches replaced: {branchesDeleted} deleted and {SeedBranches.Length} inserted. Langs: {langsDeleted} deleted and {SeedLangs.Length} inserted.",
                Stats =
                {
                    ["branchesDeleted"] = branchesDeleted,
                    ["branchesInserted"] = SeedBranches.Length,
                    ["langsDeleted"] = langsDeleted,
                    ["langsInserted"] = SeedLangs.Length,
                },
                Log = log,
            };
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync(CancellationToken.None);
            logger.LogError(ex, "Replacing branches failed; transaction rolled back");
            throw;
        }
    }
}
