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
        EligibilityReportBuilder.Apply(report, eligibility);

        var addressLookup = loaded.Addresses;
        var branchLookup = loaded.BranchMap;
        int unmappedCity = loaded.UnmappedCityCount;

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var existingIdentities = (await conn.QueryAsync<(string IdNum, int IdType)>(
            $"""
            SELECT ID_NUM AS IdNum, ID_TYPE_ID AS IdType
            FROM {connectionFactory.QualifyTable("PROFILES_TB")}
            WHERE ID_NUM IS NOT NULL
            """))
            .Select(x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum, x.IdType))
            .ToHashSet(StringComparer.Ordinal);

        var h = loaded.ClientHeaders;
        int sourceRows = 0, missingIdNum = 0, missingCombinedNumber = 0, alreadyInDb = 0, unmappedBranch = 0, withAddress = 0;
        int willInsert = 0;
        var missingIdSamples = new List<string>();
        var missingCombinedSamples = new List<string>();
        var alreadyInDbSamples = new List<string>();
        var unmappedBranchSamples = new List<string>();

        for (int r = loaded.FirstClientDataRow; r <= loaded.LastClientRow; r++)
        {
            var row = loaded.ClientSheetWs.Row(r);
            long? clientId = GetLong(row, h, "CLIENT_ID");
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

            long? combinedNumber = GetLong(row, h, "Combined Number");
            if (combinedNumber is null || combinedNumber <= 0)
            {
                missingCombinedNumber++;
                if (missingCombinedSamples.Count < SampleLimit)
                    missingCombinedSamples.Add($"Row {r} CLIENT_ID={clientId}");
                continue;
            }

            idNum = idNum.Trim();
            if (existingIdentities.Contains(ClientEligibilityClassifier.BuildDbIdentityKey(idNum, idType)))
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

            willInsert++;
        }

        report.Stats["eligibleSourceRows"] = sourceRows;
        report.Stats["missingIdNum"] = missingIdNum;
        report.Stats["missingCombinedNumber"] = missingCombinedNumber;
        report.Stats["alreadyInDb"] = alreadyInDb;
        report.Stats["existingIdentitiesInDb"] = existingIdentities.Count;
        report.Stats["willInsert"] = willInsert;
        report.Stats["willSkip"] = missingIdNum + missingCombinedNumber + alreadyInDb
            + eligibility.SkippedMissingIdCard
            + eligibility.SkippedInternalDuplicates
            + eligibility.SkippedCrossCompanyMatches
            + eligibility.SkippedDuplicateIdNums;
        report.Stats["unmappedBranch"] = unmappedBranch;
        report.Stats["unmappedCity"] = unmappedCity;
        report.Stats["withAddress"] = withAddress;

        ReportBuilder.AddIssue(report, "Warning", "MISSING_ID_NUM",
            "Eligible rows without an ID number (required merge key) — will be skipped.", missingIdNum, missingIdSamples);
        ReportBuilder.AddIssue(report, "Warning", "MISSING_COMBINED_NUMBER",
            "Eligible rows without Combined Number (required CUST_ID) — will be skipped.", missingCombinedNumber, missingCombinedSamples);
        ReportBuilder.AddIssue(report, "Info", "ALREADY_IN_DB",
            "Rows whose ID number and ID type already exist in PROFILES_TB — will be skipped.", alreadyInDb, alreadyInDbSamples);
        ReportBuilder.AddIssue(report, "Warning", "UNMAPPED_BRANCH",
            $"Rows with an unmapped COMPANY_BRANCH_CODE — will default to BranchId={DefaultBranchId}.", unmappedBranch, unmappedBranchSamples);
        ReportBuilder.AddIssue(report, "Warning", "UNMAPPED_CITY",
            "Address rows with an unmapped CITY_CODE (e.g. Asala 100+) — PermanentStateId left null.", unmappedCity);
        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(ExcelMigrationRequest? request = null, CancellationToken ct = default)
    {
        var totalTimer = Stopwatch.StartNew();
        var stageTimer = Stopwatch.StartNew();
        var excelPaths = paths.ResolveProfileExcelPaths(request?.ExcelPath);
        var log = new List<string>();

        await PreflightDatabaseAsync(log, ct);
        logger.LogInformation(
            "Profile DB preflight passed in {ElapsedMs} ms (schema={Schema})",
            stageTimer.ElapsedMilliseconds,
            connectionFactory.DatabaseSchema);

        stageTimer.Restart();
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
            $"crossCompany={eligibility.SkippedCrossCompanyMatches} duplicateIdNum={eligibility.SkippedDuplicateIdNums}");

        var rows = BuildEligibleRows(loaded, log, eligibility);
        logger.LogInformation(
            "Profile classification completed in {ElapsedMs} ms. Eligible: {Eligible}, rows: {Rows}",
            stageTimer.ElapsedMilliseconds,
            eligibility.EligibleCount,
            rows.Count);

        stageTimer.Restart();
        var (nextProfileId, seenIdentities) = await LoadExistingProfileKeysAsync(ct);
        logger.LogInformation(
            "Existing profile keys loaded in {ElapsedMs} ms. Next ProfileId={NextProfileId}, existing keys={ExistingIds}",
            stageTimer.ElapsedMilliseconds,
            nextProfileId,
            seenIdentities.Count);

        int missingIdNumCount = Math.Max(0, eligibility.EligibleCount - rows.Count);
        int alreadyInDbCount = 0;

        var batch = new List<ProfilesTb>(MigrationBatchInserter.BatchSize);
        int inserted = 0, insertFailed = 0;
        int processedRows = 0;
        int skipped = missingIdNumCount
            + eligibility.SkippedMissingIdCard
            + eligibility.SkippedInternalDuplicates
            + eligibility.SkippedCrossCompanyMatches
            + eligibility.SkippedDuplicateIdNums;

        logger.LogInformation(
            "Starting profile inserts: {RowCount} rows to process (first DB write after up to {BatchSize} rows)",
            rows.Count,
            MigrationBatchInserter.BatchSize);
        log.Add($"[INSERT START] rows={rows.Count} schema={connectionFactory.DatabaseSchema} nextProfileId={nextProfileId}");

        foreach (var source in rows)
        {
            processedRows++;
            string identityKey = ClientEligibilityClassifier.BuildDbIdentityKey(source.IdNum, source.IdType);
            if (seenIdentities.Contains(identityKey))
            {
                alreadyInDbCount++;
                skipped++;
                continue;
            }

            source.Profile.ProfileId = nextProfileId;
            nextProfileId++;
            seenIdentities.Add(identityKey);
            batch.Add(source.Profile);

            if (batch.Count >= MigrationBatchInserter.BatchSize)
            {
                stageTimer.Restart();
                int batchCount = batch.Count;
                int ok = await FlushProfilesAsync(batch, log, ct);
                inserted += ok;
                insertFailed += batch.Count - ok;
                batch.Clear();
                logger.LogInformation(
                    "Profile migration progress: {Processed}/{Total} rows, {Inserted} inserted, {Failed} failed. Last batch {BatchCount} rows in {ElapsedMs} ms",
                    processedRows,
                    rows.Count,
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
        log.Add($"Inserted={inserted} Skipped={skipped} InsertFailed={insertFailed}");
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
            Message = $"Profiles: {inserted} inserted, {skipped} skipped, {insertFailed} rejected by DB.",
            Stats =
            {
                ["total_input"] = eligibility.TotalInput,
                ["migrated_count"] = inserted,
                ["eligible_count"] = eligibility.EligibleCount,
                ["skipped_missing_id_card"] = eligibility.SkippedMissingIdCard,
                ["skipped_cross_company_matches"] = eligibility.SkippedCrossCompanyMatches,
                ["inserted"] = inserted,
                ["skipped"] = skipped,
                ["insertFailed"] = insertFailed,
                ["alreadyInDb"] = alreadyInDbCount,
                ["missingIdNum"] = missingIdNumCount,
            },
            Log = log,
        };
    }

    async Task<int> FlushProfilesAsync(List<ProfilesTb> batch, List<string> log, CancellationToken ct) =>
        await MigrationBatchInserter.InsertAsync(
            dbOptions, batch,
            p => $"[SKIP] ProfileId={p.ProfileId} IdNum={p.IdNum}: ",
            log, ct,
            message => logger.LogError("{MigrationBatchMessage}", message));

    async Task PreflightDatabaseAsync(List<string> log, CancellationToken ct)
    {
        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var schema = connectionFactory.DatabaseSchema;
        var branchesTable = connectionFactory.QualifyTable("BRANCHES_TB");
        var profilesTable = connectionFactory.QualifyTable("PROFILES_TB");

        try
        {
            int branchCount = await conn.ExecuteScalarAsync<int>($"SELECT COUNT(1) FROM {branchesTable}");
            int profileCount = await conn.ExecuteScalarAsync<int>($"SELECT COUNT(1) FROM {profilesTable}");
            log.Add($"[PREFLIGHT] schema={schema} branches={branchCount} existingProfiles={profileCount}");

            if (branchCount == 0)
            {
                throw new InvalidOperationException(
                    $"BRANCHES_TB in schema '{schema}' is empty. Run POST /api/migration/branches/run first.");
            }
        }
        catch (Oracle.ManagedDataAccess.Client.OracleException ex) when (ex.Number == 942)
        {
            throw new InvalidOperationException(
                $"Table not found in schema '{schema}'. Check DatabaseSchema in appsettings.json.", ex);
        }
    }

    async Task<(int NextProfileId, HashSet<string> SeenIdentities)> LoadExistingProfileKeysAsync(CancellationToken ct)
    {
        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var profilesTable = connectionFactory.QualifyTable("PROFILES_TB");

        int nextProfileId = ((await conn.ExecuteScalarAsync<int?>(
            $"SELECT MAX(PROFILE_ID) FROM {profilesTable}")) ?? 0) + 1;

        var seenIdentities = (await conn.QueryAsync<(string IdNum, int IdType)>(
            $"""
            SELECT ID_NUM AS IdNum, ID_TYPE_ID AS IdType
            FROM {profilesTable}
            WHERE ID_NUM IS NOT NULL
            """))
            .Select(x => ClientEligibilityClassifier.BuildDbIdentityKey(x.IdNum, x.IdType))
            .ToHashSet(StringComparer.Ordinal);

        return (nextProfileId, seenIdentities);
    }
}
