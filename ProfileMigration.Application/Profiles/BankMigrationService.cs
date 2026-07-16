using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Profiles;

/// <summary>PROFILE_BANK_INFORMATION_TB — requires profiles already migrated (matched by ID number and type).</summary>
public sealed class BankMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths)
{
    public async Task<MigrationReportDto> ValidateAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "profilesBanks", CanProceed = true };
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
        var eligibility = BuildEligibility(loaded);
        EligibilityReportBuilder.Apply(report, eligibility);
        var rows = BuildEligibleRows(loaded, [], eligibility);

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var profileByIdentity = (await conn.QueryAsync<(int ProfileId, string IdNum, int IdType)>(
                "SELECT PROFILE_ID AS ProfileId, ID_NUM AS IdNum, ID_TYPE_ID AS IdType FROM RHODES_BANKING_SILA.PROFILES_TB WHERE ID_NUM IS NOT NULL"))
            .ToDictionary(
                x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum, x.IdType),
                x => x.ProfileId,
                StringComparer.Ordinal);

        var existingBankProfiles = (await conn.QueryAsync<int>(
                "SELECT DISTINCT PROFILE_ID FROM RHODES_BANKING_SILA.PROFILE_BANK_INFORMATION_TB"))
            .ToHashSet();

        int willInsert = 0, alreadyHasBank = 0, profileNotFound = 0, noBankData = 0;
        var notFoundSamples = new List<string>();
        var alreadySamples = new List<string>();

        foreach (var g in rows)
        {
            if (g.Bank is null)
            {
                noBankData++;
                continue;
            }

            if (!profileByIdentity.TryGetValue(
                    ClientEligibilityClassifier.BuildDbIdentityKey(g.IdNum, g.IdType), out int profileId))
            {
                profileNotFound++;
                if (notFoundSamples.Count < SampleLimit)
                    notFoundSamples.Add($"IdNum={g.IdNum}");
                continue;
            }

            if (existingBankProfiles.Contains(profileId))
            {
                alreadyHasBank++;
                if (alreadySamples.Count < SampleLimit)
                    alreadySamples.Add($"ProfileId={profileId} IdNum={g.IdNum}");
                continue;
            }

            willInsert++;
        }

        report.Stats["eligibleRows"] = rows.Count;
        report.Stats["willInsert"] = willInsert;
        report.Stats["alreadyHasBank"] = alreadyHasBank;
        report.Stats["profileNotFound"] = profileNotFound;
        report.Stats["noBankData"] = noBankData;
        report.Stats["profilesInDb"] = profileByIdentity.Count;

        ReportBuilder.AddIssue(report, "Warning", "PROFILE_NOT_FOUND",
            "Bank rows whose IdNum has no matching PROFILES_TB row — run profiles first.", profileNotFound, notFoundSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_HAS_BANK",
            "Profiles that already have bank info — will be skipped.", alreadyHasBank, alreadySamples);
        ReportBuilder.AddIssue(report, "Info", "NO_BANK_DATA",
            "Eligible rows with no BANK_ACCOUNT_NUMBER — nothing to insert.", noBankData);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var rows = BuildEligibleRows(loaded, log).Where(g => g.Bank is not null).ToList();

        Dictionary<string, int> profileByIdentity;
        HashSet<int> existingBankProfiles;

        await using (var ctx = new SilaDbContext(dbOptions))
        {
            existingBankProfiles = (await ctx.ProfileBankInformationTbs.Select(b => b.ProfileId).Distinct().ToListAsync(ct)).ToHashSet();
            profileByIdentity = (await ctx.ProfilesTbs
                    .Where(p => p.IdNum != null)
                    .Select(p => new { p.ProfileId, p.IdNum, p.IdTypeId })
                    .ToListAsync(ct))
                .ToDictionary(
                    x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum!, x.IdTypeId ?? 1),
                    x => x.ProfileId,
                    StringComparer.Ordinal);
        }

        var batch = new List<ProfileBankInformationTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, skipped = 0, profileNotFound = 0, alreadyHas = 0, insertFailed = 0;

        foreach (var g in rows)
        {
            if (!profileByIdentity.TryGetValue(
                    ClientEligibilityClassifier.BuildDbIdentityKey(g.IdNum, g.IdType), out int profileId))
            {
                profileNotFound++;
                skipped++;
                continue;
            }

            if (existingBankProfiles.Contains(profileId))
            {
                alreadyHas++;
                skipped++;
                continue;
            }

            var bank = g.Bank!;
            bank.ProfileId = profileId;
            bank.Iban = $"MIG-{profileId}";
            existingBankProfiles.Add(profileId);
            batch.Add(bank);

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
            Phase = "profilesBanks",
            Success = true,
            Message = $"Banks: {inserted} inserted, {skipped} skipped, {insertFailed} rejected by DB.",
            Stats =
            {
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["profileNotFound"] = profileNotFound,
                ["alreadyHasBank"] = alreadyHas,
                ["insertFailed"] = insertFailed,
            },
            Log = log,
        };
    }

    async Task<int> FlushAsync(List<ProfileBankInformationTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            b => $"[SKIP BANK] ProfileId={b.ProfileId} BankId={b.BankId} IBAN={b.Iban}: ",
            log, ct);
}
