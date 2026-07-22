using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Profiles;

/// <summary>PROFILE_CONTACT_DTS_TB (phones) — requires profiles already migrated (matched by ID number and type).</summary>
public sealed class ContactMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths)
{
    public async Task<MigrationReportDto> ValidateAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "profilesContacts", CanProceed = true };
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
                $"SELECT PROFILE_ID AS ProfileId, ID_NUM AS IdNum, ID_TYPE_ID AS IdType FROM {connectionFactory.QualifyTable("PROFILES_TB")} WHERE ID_NUM IS NOT NULL"))
            .ToDictionary(
                x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum, x.IdType),
                x => x.ProfileId,
                StringComparer.Ordinal);

        var existingContacts = (await conn.QueryAsync<(int ProfileId, string ContactInfo)>(
                $"""
                SELECT PROFILE_ID AS ProfileId, CONTACT_INFO AS ContactInfo
                FROM {connectionFactory.QualifyTable("PROFILE_CONTACT_DTS_TB")}
                WHERE CONTACT_TYPE_ID = 1
                """))
            .Select(x => (x.ProfileId, ContactInfo: x.ContactInfo?.Trim() ?? ""))
            .ToHashSet();

        int willInsertPhones = 0, alreadyExists = 0, profileNotFound = 0, noPhoneData = 0;
        var notFoundSamples = new List<string>();

        foreach (var g in rows)
        {
            if (g.Phones.Count == 0)
            {
                noPhoneData++;
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

            foreach (var phone in g.Phones)
            {
                if (existingContacts.Contains((profileId, phone)))
                    alreadyExists++;
                else
                    willInsertPhones++;
            }
        }

        report.Stats["eligibleRows"] = rows.Count;
        report.Stats["willInsert"] = willInsertPhones;
        report.Stats["alreadyExists"] = alreadyExists;
        report.Stats["profileNotFound"] = profileNotFound;
        report.Stats["noPhoneData"] = noPhoneData;
        report.Stats["profilesInDb"] = profileByIdentity.Count;

        ReportBuilder.AddIssue(report, "Warning", "PROFILE_NOT_FOUND",
            "Contact rows whose IdNum has no matching PROFILES_TB row — run profiles first.", profileNotFound, notFoundSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_EXISTS",
            "Phone numbers already present for the same profile — will be skipped.", alreadyExists);
        ReportBuilder.AddIssue(report, "Info", "NO_PHONE_DATA",
            "Eligible rows with no TEL_NO / TEL_NO2 / TEL_NO3 — nothing to insert.", noPhoneData);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var rows = BuildEligibleRows(loaded, log).Where(g => g.Phones.Count > 0).ToList();

        Dictionary<string, int> profileByIdentity;
        HashSet<(int ProfileId, string ContactInfo)> existingContacts;
        int nextContactId;

        await using (var ctx = new SilaDbContext(dbOptions))
        {
            nextContactId = ((await ctx.ProfileContactDtsTbs.Select(c => (int?)c.ContactId).MaxAsync(ct)) ?? -1) + 1;
            existingContacts = (await ctx.ProfileContactDtsTbs
                    .Where(c => c.ContactTypeId == 1)
                    .Select(c => new { c.ProfileId, c.ContactInfo })
                    .ToListAsync(ct))
                .Select(x => (x.ProfileId, (x.ContactInfo ?? "").Trim()))
                .ToHashSet();
            profileByIdentity = (await ctx.ProfilesTbs
                    .Where(p => p.IdNum != null)
                    .Select(p => new { p.ProfileId, p.IdNum, p.IdTypeId })
                    .ToListAsync(ct))
                .ToDictionary(
                    x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum!, x.IdTypeId ?? 1),
                    x => x.ProfileId,
                    StringComparer.Ordinal);
        }

        var batch = new List<ProfileContactDtsTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, skipped = 0, profileNotFound = 0, alreadyExists = 0, insertFailed = 0;
        var now = DateTime.Now;

        foreach (var g in rows)
        {
            if (!profileByIdentity.TryGetValue(
                    ClientEligibilityClassifier.BuildDbIdentityKey(g.IdNum, g.IdType), out int profileId))
            {
                profileNotFound++;
                skipped += g.Phones.Count;
                continue;
            }

            foreach (var phone in g.Phones)
            {
                if (existingContacts.Contains((profileId, phone)))
                {
                    alreadyExists++;
                    skipped++;
                    continue;
                }

                existingContacts.Add((profileId, phone));
                batch.Add(new ProfileContactDtsTb
                {
                    ContactId = nextContactId++,
                    ProfileId = profileId,
                    ContactTypeId = 1,
                    ContactInfo = phone,
                    CreatedOn = now,
                });

                if (batch.Count >= MigrationBatchInserter.BatchSize)
                {
                    int ok = await FlushAsync(batch, log, ct);
                    inserted += ok;
                    insertFailed += batch.Count - ok;
                    batch.Clear();
                }
            }
        }

        if (batch.Count > 0)
        {
            int ok = await FlushAsync(batch, log, ct);
            inserted += ok;
            insertFailed += batch.Count - ok;
        }

        log.Add($"Inserted={inserted} Skipped={skipped} ProfileNotFound={profileNotFound} AlreadyExists={alreadyExists} InsertFailed={insertFailed}");

        return new PhaseRunResultDto
        {
            Phase = "profilesContacts",
            Success = true,
            Message = $"Contacts: {inserted} inserted, {skipped} skipped, {insertFailed} rejected by DB.",
            Stats =
            {
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["profileNotFound"] = profileNotFound,
                ["alreadyExists"] = alreadyExists,
                ["insertFailed"] = insertFailed,
            },
            Log = log,
        };
    }

    async Task<int> FlushAsync(List<ProfileContactDtsTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            c => $"[SKIP CONTACT] ContactId={c.ContactId} ProfileId={c.ProfileId}: ",
            log, ct);
}
