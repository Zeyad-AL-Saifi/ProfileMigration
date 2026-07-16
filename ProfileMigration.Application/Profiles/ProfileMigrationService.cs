using System.Diagnostics;
using ClosedXML.Excel;
using Dapper;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Profiles;

/// <summary>
/// Validates and runs PROFILES_TB migration only (addresses applied on the profile row).
/// Works / partners / banks / contacts are separate phases.
/// </summary>
public sealed class ProfileMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths,
    ILogger<ProfileMigrationService> logger)
{
    public async Task<MigrationReportDto> ValidateAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "profiles", CanProceed = true };
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        report.Stats["clientExcelPath"] = excelPaths.ClientPath;
        report.Stats["idCardExcelPath"] = excelPaths.IdCardPath;
        report.Stats["addressExcelPath"] = excelPaths.AddressPath;

        var missingFiles = new[]
        {
            (Source: "clients", Path: excelPaths.ClientPath),
            (Source: "idCards", Path: excelPaths.IdCardPath),
            (Source: "addresses", Path: excelPaths.AddressPath),
        }.Where(x => !File.Exists(x.Path)).ToList();
        if (missingFiles.Count > 0)
        {
            ReportBuilder.AddIssue(report, "Error", "EXCEL_FILE_MISSING",
                "One or more profile Excel source files were not found.",
                missingFiles.Count,
                missingFiles.Select(x => $"{x.Source}: {x.Path}").ToList());
            return report;
        }

        using var clientProbe = new XLWorkbook(excelPaths.ClientPath);
        using var idCardProbe = new XLWorkbook(excelPaths.IdCardPath);
        using var addressProbe = new XLWorkbook(excelPaths.AddressPath);
        var sheetOrColErrors = ValidateRequiredSheetsAndColumns(clientProbe, idCardProbe, addressProbe);
        if (sheetOrColErrors.Count > 0)
        {
            var sheets = sheetOrColErrors
                .Where(e => e.Kind == "sheet")
                .Select(e => $"{e.Source}: {e.Name}")
                .ToList();
            if (sheets.Count > 0)
                ReportBuilder.AddIssue(report, "Error", "SHEET_MISSING",
                    "Required worksheet(s) not found in the profile Excel sources.", sheets.Count, sheets);
            var cols = sheetOrColErrors
                .Where(e => e.Kind == "column")
                .Select(e => $"{e.Source}: {e.Name}")
                .ToList();
            if (cols.Count > 0)
                ReportBuilder.AddIssue(report, "Error", "COLUMN_MISSING",
                    "Required column(s) not found in the profile Excel sources.", cols.Count, cols);
            return report;
        }

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        var eligibility = BuildEligibility(loaded);
        ApplyEligibilityStats(report, eligibility);

        var addressLookup = loaded.Addresses;
        var branchLookup = loaded.BranchMap;
        int unmappedCity = loaded.UnmappedCityCount;

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var existingIdNums = (await conn.QueryAsync<string>(
            "SELECT ID_NUM FROM RHODES_BANKING_SILA.PROFILES_TB WHERE ID_NUM IS NOT NULL"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var h = loaded.ClientHeaders;
        int sourceRows = 0, missingIdNum = 0, alreadyInDb = 0, unmappedBranch = 0, withAddress = 0;
        var groupCounts = new Dictionary<(string IdNum, int IdType), int>();
        var missingIdSamples = new List<string>();
        var alreadyInDbSamples = new List<string>();
        var unmappedBranchSamples = new List<string>();

        for (int r = loaded.FirstClientDataRow; r <= loaded.LastClientRow; r++)
        {
            var row = loaded.ClientSheetWs.Row(r);
            int? clientId = GetInt(row, h, "CLIENT_ID");
            if (clientId is null || clientId <= 0) continue;

            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, h, "COMPANY"));
            if (!eligibility.IsEligible(company, clientId.Value))
                continue;

            sourceRows++;

            loaded.IdCards.TryGetValue((company, clientId.Value), out IdCardData? idCard);
            string? idNum = idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID");
            int idType = idCard?.IdTypeId ?? 1;

            if (string.IsNullOrWhiteSpace(idNum))
            {
                missingIdNum++;
                if (missingIdSamples.Count < SampleLimit)
                    missingIdSamples.Add($"Row {r} CLIENT_ID={clientId}");
                continue;
            }

            idNum = idNum.Trim();
            if (existingIdNums.Contains(idNum))
            {
                alreadyInDb++;
                if (alreadyInDbSamples.Count < SampleLimit)
                    alreadyInDbSamples.Add($"Row {r} IdNum={idNum}");
                continue;
            }

            int? oldBrCode = GetInt(row, h, "COMPANY_BRANCH_CODE");
            if (oldBrCode.HasValue && !branchLookup.ContainsKey((company, oldBrCode.Value)))
            {
                unmappedBranch++;
                if (unmappedBranchSamples.Count < SampleLimit)
                    unmappedBranchSamples.Add($"Row {r} COMPANY={company} OLD_ID={oldBrCode}");
            }

            if (addressLookup.ContainsKey((company, clientId.Value)))
                withAddress++;

            var key = (idNum, idType);
            groupCounts[key] = groupCounts.TryGetValue(key, out int cnt) ? cnt + 1 : 1;
        }

        int willInsert = groupCounts.Count;
        int multiRowGroups = groupCounts.Count(g => g.Value > 1);

        report.Stats["eligibleSourceRows"] = sourceRows;
        report.Stats["missingIdNum"] = missingIdNum;
        report.Stats["alreadyInDb"] = alreadyInDb;
        report.Stats["existingIdNumsInDb"] = existingIdNums.Count;
        report.Stats["willInsert"] = willInsert;
        report.Stats["willSkip"] = missingIdNum + alreadyInDb
            + eligibility.SkippedMissingIdCard
            + eligibility.SkippedInternalDuplicates
            + eligibility.SkippedCrossCompanyMatches;
        report.Stats["willMerge"] = multiRowGroups;
        report.Stats["multiRowGroups"] = multiRowGroups;
        report.Stats["unmappedBranch"] = unmappedBranch;
        report.Stats["unmappedCity"] = unmappedCity;
        report.Stats["withAddress"] = withAddress;

        ReportBuilder.AddIssue(report, "Warning", "MISSING_ID_NUM",
            "Eligible rows without an ID number (required merge key) — will be skipped.", missingIdNum, missingIdSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_IN_DB",
            "Rows whose ID number already exists in PROFILES_TB — will be skipped.", alreadyInDb, alreadyInDbSamples);
        ReportBuilder.AddIssue(report, "Warning", "UNMAPPED_BRANCH",
            $"Rows with an unmapped COMPANY_BRANCH_CODE — will default to BranchId={DefaultBranchId}.", unmappedBranch, unmappedBranchSamples);
        ReportBuilder.AddIssue(report, "Warning", "UNMAPPED_CITY",
            "Address rows with an unmapped CITY_CODE (e.g. Asala 100+) — PermanentStateId left null.", unmappedCity);
        ReportBuilder.AddIssue(report, "Info", "MULTI_ROW_MERGE_CANDIDATES",
            "ID numbers appearing in more than one source row — will be merged into a single profile.", multiRowGroups);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var totalTimer = Stopwatch.StartNew();
        var stageTimer = Stopwatch.StartNew();
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        using var loaded = Open(excelPaths.ClientPath, excelPaths.IdCardPath, excelPaths.AddressPath);
        logger.LogInformation(
            "Profile Excel sources loaded in {ElapsedMs} ms. Client rows: {ClientRows}, ID cards: {IdCards}, addresses: {Addresses}",
            stageTimer.ElapsedMilliseconds,
            Math.Max(0, loaded.LastClientRow - loaded.FirstClientDataRow + 1),
            loaded.IdCards.Count,
            loaded.Addresses.Count);

        stageTimer.Restart();
        var eligibility = BuildEligibility(loaded);
        log.Add(
            $"[ELIGIBILITY] total={eligibility.TotalInput} eligible={eligibility.EligibleCount} " +
            $"missingCard={eligibility.SkippedMissingIdCard} internalDup={eligibility.SkippedInternalDuplicates} " +
            $"crossCompany={eligibility.SkippedCrossCompanyMatches}");

        var groups = BuildMergedGroups(loaded, log, eligibility);
        logger.LogInformation(
            "Profile classification and merge completed in {ElapsedMs} ms. Eligible: {Eligible}, groups: {Groups}",
            stageTimer.ElapsedMilliseconds,
            eligibility.EligibleCount,
            groups.Count);

        int nextProfileId;
        var seenIdNums = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        stageTimer.Restart();
        await using (var seqCtx = new SilaDbContext(dbOptions))
        {
            nextProfileId = ((await seqCtx.ProfilesTbs.Select(p => (int?)p.ProfileId).MaxAsync(ct)) ?? -1) + 1;
            var existingIdNums = await seqCtx.ProfilesTbs
                .Where(p => p.IdNum != null)
                .Select(p => p.IdNum!)
                .ToListAsync(ct);
            foreach (var n in existingIdNums) seenIdNums.Add(n);
        }
        logger.LogInformation(
            "Existing profile keys loaded in {ElapsedMs} ms. Existing IDs: {ExistingIds}",
            stageTimer.ElapsedMilliseconds,
            seenIdNums.Count);

        string conflictFile = Path.Combine(AppContext.BaseDirectory, "merge_conflicts.txt");
        await using var conflictWriter = new StreamWriter(conflictFile, append: false);
        await conflictWriter.WriteLineAsync($"Merge review — generated {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        await conflictWriter.WriteLineAsync("Rows listed here were NOT inserted and need manual handling.");
        await conflictWriter.WriteLineAsync();
        int reviewCount = 0, missingIdNumCount = 0, alreadyInDbCount = 0;

        foreach (var (key, reason) in eligibility.Skipped.OrderBy(x => x.Key.Company).ThenBy(x => x.Key.ClientId))
        {
            await conflictWriter.WriteLineAsync($"SKIPPED {reason}: {key.Company} CLIENT_ID={key.ClientId}");
            await conflictWriter.WriteLineAsync();
            reviewCount++;
        }

        var h = loaded.ClientHeaders;
        for (int r = loaded.FirstClientDataRow; r <= loaded.LastClientRow; r++)
        {
            var row = loaded.ClientSheetWs.Row(r);
            int? clientId = GetInt(row, h, "CLIENT_ID");
            if (clientId is null || clientId <= 0) continue;
            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, h, "COMPANY"));
            if (!eligibility.IsEligible(company, clientId.Value)) continue;

            loaded.IdCards.TryGetValue((company, clientId.Value), out IdCardData? idCard);
            string? idNum = idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID");
            if (string.IsNullOrWhiteSpace(idNum))
            {
                await conflictWriter.WriteLineAsync(
                    $"Row {r} (CLIENT_ID={clientId}): missing ID_NUM — required key, not inserted.");
                await conflictWriter.WriteLineAsync();
                reviewCount++;
                missingIdNumCount++;
            }
        }

        var batch = new List<ProfilesTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, highConflict = 0, insertFailed = 0;
        int processedGroups = 0;
        int skipped = missingIdNumCount
            + eligibility.SkippedMissingIdCard
            + eligibility.SkippedInternalDuplicates
            + eligibility.SkippedCrossCompanyMatches;

        foreach (var g in groups)
        {
            processedGroups++;
            if (seenIdNums.Contains(g.IdNum))
            {
                alreadyInDbCount++;
                skipped++;
                continue;
            }

            if (g.HighConflict)
            {
                await conflictWriter.WriteLineAsync(
                    $"ID_NUM={g.IdNum} ID_TYPE={g.IdType}: {g.Conflicts.Count} conflicting fields (> {MaxConflicts}) — NOT inserted.");
                foreach (var conflict in g.Conflicts)
                    await conflictWriter.WriteLineAsync($"    {conflict}");
                await conflictWriter.WriteLineAsync();
                reviewCount++;
                highConflict++;
                skipped += g.SourceRowCount;
                continue;
            }

            if (g.WasMerged)
                log.Add($"[MERGE] ID_NUM={g.IdNum} ID_TYPE={g.IdType} rows={g.SourceRowCount} conflicts={g.Conflicts.Count}");

            g.Profile.ProfileId = nextProfileId;
            nextProfileId++;
            seenIdNums.Add(g.IdNum);
            batch.Add(g.Profile);

            if (batch.Count >= MigrationBatchInserter.BatchSize)
            {
                stageTimer.Restart();
                int batchCount = batch.Count;
                int ok = await FlushProfilesAsync(batch, log, ct);
                inserted += ok;
                insertFailed += batch.Count - ok;
                batch.Clear();
                logger.LogInformation(
                    "Profile migration progress: {Processed}/{Total} groups, {Inserted} inserted, {Failed} failed. Last batch {BatchCount} rows in {ElapsedMs} ms",
                    processedGroups,
                    groups.Count,
                    inserted,
                    insertFailed,
                    batchCount,
                    stageTimer.ElapsedMilliseconds);
            }
        }

        if (batch.Count > 0)
        {
            stageTimer.Restart();
            int batchCount = batch.Count;
            int ok = await FlushProfilesAsync(batch, log, ct);
            inserted += ok;
            insertFailed += batch.Count - ok;
            logger.LogInformation(
                "Final profile batch inserted {Inserted}/{BatchCount} rows in {ElapsedMs} ms",
                ok,
                batchCount,
                stageTimer.ElapsedMilliseconds);
        }

        skipped += alreadyInDbCount;
        log.Add($"Inserted={inserted} Skipped={skipped} HighConflict={highConflict} InsertFailed={insertFailed}");
        if (reviewCount > 0)
            log.Add($"Review file: {conflictFile} ({reviewCount} entries)");
        logger.LogInformation(
            "Profile migration completed in {Elapsed}. Inserted: {Inserted}, skipped: {Skipped}, failed: {Failed}",
            totalTimer.Elapsed,
            inserted,
            skipped,
            insertFailed);

        return new PhaseRunResultDto
        {
            Phase = "profiles",
            Success = true,
            Message = $"Profiles: {inserted} inserted, {skipped} skipped ({highConflict} high-conflict), {insertFailed} rejected by DB.",
            Stats =
            {
                ["total_input"] = eligibility.TotalInput,
                ["migrated_count"] = inserted,
                ["eligible_count"] = eligibility.EligibleCount,
                ["skipped_missing_id_card"] = eligibility.SkippedMissingIdCard,
                ["skipped_internal_duplicates"] = eligibility.SkippedInternalDuplicates,
                ["skipped_cross_company_matches"] = eligibility.SkippedCrossCompanyMatches,
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["highConflict"] = highConflict,
                ["insertFailed"] = insertFailed,
                ["alreadyInDb"] = alreadyInDbCount,
                ["missingIdNum"] = missingIdNumCount,
                ["reviewCount"] = reviewCount,
                ["mergeConflictsFile"] = conflictFile,
            },
            Log = log,
        };
    }

    static void ApplyEligibilityStats(MigrationReportDto report, ClientEligibilityClassifier.Result eligibility)
    {
        report.Stats["total_input"] = eligibility.TotalInput;
        report.Stats["eligible_count"] = eligibility.EligibleCount;
        report.Stats["skipped_missing_id_card"] = eligibility.SkippedMissingIdCard;
        report.Stats["skipped_internal_duplicates"] = eligibility.SkippedInternalDuplicates;
        report.Stats["skipped_cross_company_matches"] = eligibility.SkippedCrossCompanyMatches;

        ReportBuilder.AddIssue(report, "Warning", ClientEligibilityClassifier.ReasonMissingIdCard,
            "Clients without ID_CARD_TYPE+ID_CARD_NO (Analysis join) — never migrated.",
            eligibility.SkippedMissingIdCard,
            eligibility.SampleSkipped(ClientEligibilityClassifier.ReasonMissingIdCard));

        ReportBuilder.AddIssue(report, "Warning", ClientEligibilityClassifier.ReasonInternalDuplicateAsala,
            "ASALA clients sharing the same card_key (2+) — all excluded.",
            eligibility.Skipped.Count(kv => kv.Value == ClientEligibilityClassifier.ReasonInternalDuplicateAsala),
            eligibility.SampleSkipped(ClientEligibilityClassifier.ReasonInternalDuplicateAsala));

        ReportBuilder.AddIssue(report, "Warning", ClientEligibilityClassifier.ReasonInternalDuplicateAcad,
            "ACAD clients sharing the same card_key (2+) — all excluded.",
            eligibility.Skipped.Count(kv => kv.Value == ClientEligibilityClassifier.ReasonInternalDuplicateAcad),
            eligibility.SampleSkipped(ClientEligibilityClassifier.ReasonInternalDuplicateAcad));

        ReportBuilder.AddIssue(report, "Warning", ClientEligibilityClassifier.ReasonCrossCompanyMatch,
            "card_key present in BOTH ASALA and ACAD — all CLIENT_IDs excluded at any match %.",
            eligibility.SkippedCrossCompanyMatches,
            eligibility.SampleSkipped(ClientEligibilityClassifier.ReasonCrossCompanyMatch));
    }

    async Task<int> FlushProfilesAsync(List<ProfilesTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            p => $"[SKIP] ProfileId={p.ProfileId} IdNum={p.IdNum}: ",
            log, ct);
}
