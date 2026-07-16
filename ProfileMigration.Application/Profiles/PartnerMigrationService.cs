using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Profiles;

/// <summary>PROFILES_PARTNERS_TB — requires profiles already migrated (matched by IdNum).</summary>
public sealed class PartnerMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths)
{
    public async Task<MigrationReportDto> ValidateAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "profilesPartners", CanProceed = true };
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
        var log = new List<string>();
        var groups = BuildMergedGroups(loaded, log).Where(g => !g.HighConflict).ToList();

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var profileByIdNum = (await conn.QueryAsync<(int ProfileId, string IdNum)>(
                "SELECT PROFILE_ID AS ProfileId, ID_NUM AS IdNum FROM RHODES_BANKING_SILA.PROFILES_TB WHERE ID_NUM IS NOT NULL"))
            .GroupBy(x => x.IdNum, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().ProfileId, StringComparer.OrdinalIgnoreCase);

        var existingPartnerProfiles = (await conn.QueryAsync<int>(
                "SELECT DISTINCT PROFILE_ID FROM RHODES_BANKING_SILA.PROFILES_PARTNERS_TB"))
            .ToHashSet();

        int willInsert = 0, alreadyHasPartner = 0, profileNotFound = 0, noPartnerData = 0;
        var notFoundSamples = new List<string>();
        var alreadySamples = new List<string>();

        foreach (var g in groups)
        {
            if (g.Partner is null)
            {
                noPartnerData++;
                continue;
            }

            if (!profileByIdNum.TryGetValue(g.IdNum, out int profileId))
            {
                profileNotFound++;
                if (notFoundSamples.Count < SampleLimit)
                    notFoundSamples.Add($"IdNum={g.IdNum}");
                continue;
            }

            if (existingPartnerProfiles.Contains(profileId))
            {
                alreadyHasPartner++;
                if (alreadySamples.Count < SampleLimit)
                    alreadySamples.Add($"ProfileId={profileId} IdNum={g.IdNum}");
                continue;
            }

            willInsert++;
        }

        report.Stats["mergedGroups"] = groups.Count;
        report.Stats["willInsert"] = willInsert;
        report.Stats["alreadyHasPartner"] = alreadyHasPartner;
        report.Stats["profileNotFound"] = profileNotFound;
        report.Stats["noPartnerData"] = noPartnerData;
        report.Stats["partnerSkippedMissingId"] = log.Count(l => l.Contains("SKIP PARTNER", StringComparison.Ordinal));
        report.Stats["profilesInDb"] = profileByIdNum.Count;

        ReportBuilder.AddIssue(report, "Warning", "PROFILE_NOT_FOUND",
            "Partner rows whose IdNum has no matching PROFILES_TB row — run profiles first.", profileNotFound, notFoundSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_HAS_PARTNER",
            "Profiles that already have a partner — will be skipped.", alreadyHasPartner, alreadySamples);
        ReportBuilder.AddIssue(report, "Info", "NO_PARTNER_DATA",
            "Merged groups with no partner (missing name or national id) — nothing to insert.", noPartnerData);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var groups = BuildMergedGroups(loaded, log).Where(g => !g.HighConflict && g.Partner is not null).ToList();

        Dictionary<string, int> profileByIdNum;
        HashSet<int> existingPartnerProfiles;
        int nextPartnerId;

        await using (var ctx = new SilaDbContext(dbOptions))
        {
            nextPartnerId = ((await ctx.ProfilesPartnersTbs.Select(p => (int?)p.PartnerId).MaxAsync(ct)) ?? -1) + 1;
            existingPartnerProfiles = (await ctx.ProfilesPartnersTbs.Select(p => p.ProfileId).Distinct().ToListAsync(ct)).ToHashSet();
            profileByIdNum = (await ctx.ProfilesTbs
                    .Where(p => p.IdNum != null)
                    .Select(p => new { p.ProfileId, p.IdNum })
                    .ToListAsync(ct))
                .GroupBy(x => x.IdNum!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().ProfileId, StringComparer.OrdinalIgnoreCase);
        }

        var batch = new List<ProfilesPartnersTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, skipped = 0, profileNotFound = 0, alreadyHas = 0, insertFailed = 0;

        foreach (var g in groups)
        {
            if (!profileByIdNum.TryGetValue(g.IdNum, out int profileId))
            {
                profileNotFound++;
                skipped++;
                continue;
            }

            if (existingPartnerProfiles.Contains(profileId))
            {
                alreadyHas++;
                skipped++;
                continue;
            }

            var partner = g.Partner!;
            partner.ProfileId = profileId;
            partner.PartnerId = nextPartnerId++;
            existingPartnerProfiles.Add(profileId);
            batch.Add(partner);

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
            Phase = "profilesPartners",
            Success = true,
            Message = $"Partners: {inserted} inserted, {skipped} skipped, {insertFailed} rejected by DB.",
            Stats =
            {
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["profileNotFound"] = profileNotFound,
                ["alreadyHasPartner"] = alreadyHas,
                ["insertFailed"] = insertFailed,
            },
            Log = log,
        };
    }

    async Task<int> FlushAsync(List<ProfilesPartnersTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            p => $"[SKIP PARTNER] PartnerId={p.PartnerId} ProfileId={p.ProfileId}: ",
            log, ct);
}
