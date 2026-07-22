using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application.Branches;

public sealed class BranchMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory)
{
    static readonly (int Id, string NameEn, bool IsCentral)[] SeedBranches =
    [
        (1, "Jenin", false), (2, "Tulkarm", false), (3, "Tubas", false),
        (4, "Nablus", false), (5, "Ramallah", true), (6, "Jericho", false),
        (7, "Bethlehem", false), (8, "Hebron", false), (9, "Gaza", false),
    ];

    static readonly (int BranchId, string LangId, string Name)[] SeedLangs =
    [
        (1, "en", "Jenin"), (1, "ar", "جنين"),
        (2, "en", "Tulkarm"), (2, "ar", "طولكرم"),
        (3, "en", "Tubas"), (3, "ar", "طوباس"),
        (4, "en", "Nablus"), (4, "ar", "نابلس"),
        (5, "en", "Ramallah"), (5, "ar", "رام الله"),
        (6, "en", "Jericho"), (6, "ar", "أريحا"),
        (7, "en", "Bethlehem"), (7, "ar", "بيت لحم"),
        (8, "en", "Hebron"), (8, "ar", "الخليل"),
        (9, "en", "Gaza"), (9, "ar", "غزة"),
    ];

    public async Task<MigrationReportDto> ValidateAsync(CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "branches", CanProceed = true };

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);

        var existingBranchIds = (await conn.QueryAsync<int>(
            $"SELECT BRANCH_ID FROM {connectionFactory.QualifyTable("BRANCHES_TB")} WHERE BRANCH_ID BETWEEN 1 AND 9"))
            .ToHashSet();

        var existingLangs = (await conn.QueryAsync<(int BranchId, string LangId)>(
            $"""
            SELECT BRANCH_ID AS BranchId, LOWER(LANG_ID) AS LangId
            FROM {connectionFactory.QualifyTable("BRANCH_LANGS_TB")}
            WHERE BRANCH_ID BETWEEN 1 AND 9
            """)).Select(x => (x.BranchId, x.LangId)).ToHashSet();

        int missingBranches = SeedBranches.Count(b => !existingBranchIds.Contains(b.Id));
        int missingLangs = SeedLangs.Count(l => !existingLangs.Contains((l.BranchId, l.LangId.ToLowerInvariant())));

        report.Stats["expectedBranches"] = SeedBranches.Length;
        report.Stats["existingBranches"] = existingBranchIds.Count;
        report.Stats["willInsertBranches"] = missingBranches;
        report.Stats["willSkipBranches"] = SeedBranches.Length - missingBranches;
        report.Stats["expectedLangs"] = SeedLangs.Length;
        report.Stats["willInsertLangs"] = missingLangs;
        report.Stats["willSkipLangs"] = SeedLangs.Length - missingLangs;

        ReportBuilder.AddIssue(report, "Info", "BRANCHES_MISSING",
            "Branches that will be inserted (1-9).", missingBranches,
            SeedBranches.Where(b => !existingBranchIds.Contains(b.Id)).Select(b => $"{b.Id}:{b.NameEn}"));

        ReportBuilder.AddIssue(report, "Info", "BRANCH_LANGS_MISSING",
            "Branch language rows that will be inserted.", missingLangs);

        report.Breakdown["missingBranchIds"] = SeedBranches
            .Where(b => !existingBranchIds.Contains(b.Id)).Select(b => b.Id).ToList();

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(CancellationToken ct = default)
    {
        var log = new List<string>();
        const int CountryId = 1;
        var today = DateTime.Today;
        var now = DateTime.Now;

        await using var ctx = new SilaDbContext(dbOptions);

        var existingIds = (await ctx.BranchesTbs.Select(b => b.BranchId).ToListAsync(ct)).ToHashSet();
        var existingLangKeys = (await ctx.BranchLangsTbs
                .Select(l => new { l.BranchId, l.LangId }).ToListAsync(ct))
            .Select(x => (x.BranchId, LangId: x.LangId.ToLowerInvariant()))
            .ToHashSet();

        int branchesInserted = 0, branchesSkipped = 0, langsInserted = 0, langsSkipped = 0;

        foreach (var (id, nameEn, isCentral) in SeedBranches)
        {
            if (existingIds.Contains(id))
            {
                branchesSkipped++;
                log.Add($"[SKIP] BRANCHES_TB BranchId={id} ({nameEn})");
                continue;
            }

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
                GovernorateId = id,
            });
            branchesInserted++;
            log.Add($"[ADD] BRANCHES_TB BranchId={id} ({nameEn})");
        }

        foreach (var (branchId, langId, name) in SeedLangs)
        {
            if (existingLangKeys.Contains((branchId, langId.ToLowerInvariant())))
            {
                langsSkipped++;
                log.Add($"[SKIP] BRANCH_LANGS_TB ({branchId},{langId})");
                continue;
            }

            ctx.BranchLangsTbs.Add(new BranchLangsTb
            {
                BranchId = branchId,
                LangId = langId,
                BranchName = name,
                CreatedOn = now,
            });
            langsInserted++;
            log.Add($"[ADD] BRANCH_LANGS_TB ({branchId},{langId}) '{name}'");
        }

        if (branchesInserted + langsInserted > 0)
            await ctx.SaveChangesAsync(ct);

        return new PhaseRunResultDto
        {
            Phase = "branches",
            Success = true,
            Message = $"Branches: {branchesInserted} inserted, {branchesSkipped} skipped. Langs: {langsInserted} inserted, {langsSkipped} skipped.",
            Stats =
            {
                ["branchesInserted"] = branchesInserted,
                ["branchesSkipped"] = branchesSkipped,
                ["langsInserted"] = langsInserted,
                ["langsSkipped"] = langsSkipped,
            },
            Log = log,
        };
    }
}
