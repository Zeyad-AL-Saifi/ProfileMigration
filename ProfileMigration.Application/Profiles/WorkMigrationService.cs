using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Profiles;

/// <summary>PROFILE_WORKS_TB — requires profiles already migrated (matched by IdNum).</summary>
public sealed class WorkMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths)
{
    public async Task<MigrationReportDto> ValidateAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "profilesWorks", CanProceed = true };
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        report.Stats["clientExcelPath"] = excelPaths.ClientPath;
        report.Stats["idCardExcelPath"] = excelPaths.IdCardPath;
        report.Stats["addressExcelPath"] = excelPaths.AddressPath;

        var sourceErrors = ValidateSourceFiles(
            excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        if (sourceErrors.Count > 0)
        {
            ReportBuilder.AddIssue(report, "Error", "PROFILE_EXCEL_SOURCE_INVALID",
                "Profile Excel source files, sheets, or required columns are invalid.",
                sourceErrors.Count,
                sourceErrors.Select(e => $"{e.Source}.{e.Kind}: {e.Name}"));
            return report;
        }

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var allGroups = BuildMergedGroups(loaded, []);
        int highConflict = allGroups.Count(g => g.HighConflict);
        var groups = allGroups.Where(g => !g.HighConflict).ToList();

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var profileByIdNum = (await conn.QueryAsync<(int ProfileId, string IdNum)>(
                "SELECT PROFILE_ID AS ProfileId, ID_NUM AS IdNum FROM RHODES_BANKING_SILA.PROFILES_TB WHERE ID_NUM IS NOT NULL"))
            .GroupBy(x => x.IdNum, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().ProfileId, StringComparer.OrdinalIgnoreCase);

        var existingWorkProfiles = (await conn.QueryAsync<int>(
                "SELECT DISTINCT PROFILE_ID FROM RHODES_BANKING_SILA.PROFILE_WORKS_TB"))
            .ToHashSet();

        int willInsert = 0, alreadyHasWork = 0, profileNotFound = 0, noWorkData = 0;
        var notFoundSamples = new List<string>();
        var alreadySamples = new List<string>();

        foreach (var g in groups)
        {
            if (g.Work is null)
            {
                noWorkData++;
                continue;
            }

            if (!profileByIdNum.TryGetValue(g.IdNum, out int profileId))
            {
                profileNotFound++;
                if (notFoundSamples.Count < SampleLimit)
                    notFoundSamples.Add($"IdNum={g.IdNum}");
                continue;
            }

            if (existingWorkProfiles.Contains(profileId))
            {
                alreadyHasWork++;
                if (alreadySamples.Count < SampleLimit)
                    alreadySamples.Add($"ProfileId={profileId} IdNum={g.IdNum}");
                continue;
            }

            willInsert++;
        }

        report.Stats["mergedGroups"] = groups.Count;
        report.Stats["willInsert"] = willInsert;
        report.Stats["alreadyHasWork"] = alreadyHasWork;
        report.Stats["profileNotFound"] = profileNotFound;
        report.Stats["noWorkData"] = noWorkData;
        report.Stats["highConflictSkipped"] = highConflict;
        report.Stats["profilesInDb"] = profileByIdNum.Count;

        ReportBuilder.AddIssue(report, "Warning", "PROFILE_NOT_FOUND",
            "Work rows whose IdNum has no matching PROFILES_TB row — run profiles first.", profileNotFound, notFoundSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_HAS_WORK",
            "Profiles that already have PROFILE_WORKS_TB — will be skipped.", alreadyHasWork, alreadySamples);
        ReportBuilder.AddIssue(report, "Info", "NO_WORK_DATA",
            "Merged groups with no WORK_PLACE / PROFESSION / MONTHLY_INCOME — nothing to insert.", noWorkData);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var groups = BuildMergedGroups(loaded, log).Where(g => !g.HighConflict && g.Work is not null).ToList();

        Dictionary<string, int> profileByIdNum;
        HashSet<int> existingWorkProfiles;
        int nextWorkId;

        await using (var ctx = new SilaDbContext(dbOptions))
        {
            nextWorkId = ((await ctx.ProfileWorksTbs.Select(w => (int?)w.WorkId).MaxAsync(ct)) ?? -1) + 1;
            existingWorkProfiles = (await ctx.ProfileWorksTbs.Select(w => w.ProfileId).Distinct().ToListAsync(ct)).ToHashSet();
            profileByIdNum = (await ctx.ProfilesTbs
                    .Where(p => p.IdNum != null)
                    .Select(p => new { p.ProfileId, p.IdNum })
                    .ToListAsync(ct))
                .GroupBy(x => x.IdNum!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().ProfileId, StringComparer.OrdinalIgnoreCase);
        }

        var batch = new List<ProfileWorksTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, skipped = 0, profileNotFound = 0, alreadyHas = 0, insertFailed = 0;

        foreach (var g in groups)
        {
            if (!profileByIdNum.TryGetValue(g.IdNum, out int profileId))
            {
                profileNotFound++;
                skipped++;
                continue;
            }

            if (existingWorkProfiles.Contains(profileId))
            {
                alreadyHas++;
                skipped++;
                continue;
            }

            var work = g.Work!;
            work.ProfileId = profileId;
            work.WorkId = nextWorkId++;
            existingWorkProfiles.Add(profileId);
            batch.Add(work);

            if (batch.Count >= MigrationBatchInserter.BatchSize)
            {
                int ok = await FlushAsync(batch, log, ct);
                inserted += ok;
                insertFailed += batch.Count - ok;
                batch.Clear();
            }
        }

        if (batch.Count > 0)
        {
            int ok = await FlushAsync(batch, log, ct);
            inserted += ok;
            insertFailed += batch.Count - ok;
        }

        log.Add($"Inserted={inserted} Skipped={skipped} ProfileNotFound={profileNotFound} AlreadyHas={alreadyHas} InsertFailed={insertFailed}");

        return new PhaseRunResultDto
        {
            Phase = "profilesWorks",
            Success = true,
            Message = $"Works: {inserted} inserted, {skipped} skipped, {insertFailed} rejected by DB.",
            Stats =
            {
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["profileNotFound"] = profileNotFound,
                ["alreadyHasWork"] = alreadyHas,
                ["insertFailed"] = insertFailed,
            },
            Log = log,
        };
    }

    async Task<int> FlushAsync(List<ProfileWorksTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            w => $"[SKIP WORK] WorkId={w.WorkId} ProfileId={w.ProfileId}: ",
            log, ct);
}
