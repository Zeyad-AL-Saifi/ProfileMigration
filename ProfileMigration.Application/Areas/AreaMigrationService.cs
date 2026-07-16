using ClosedXML.Excel;
using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;

namespace ProfileMigration.Application.Areas;

public sealed class AreaMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory,
    MigrationPathResolver paths)
{
    const int AreaMainId = 10;

    public async Task<MigrationReportDto> ValidateAsync(AreaMigrationRequest? request = null, CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "areas", CanProceed = true };
        var areaPath = paths.ResolveAreaExcelPath(request?.AreaExcelPath);

        if (!File.Exists(areaPath))
        {
            ReportBuilder.AddIssue(report, "Error", "AREA_FILE_MISSING",
                $"Area Excel not found: {areaPath}", 1, [areaPath]);
            return report;
        }

        var (valid, skipped, dupOldArea, samples) = AnalyzeAreaExcel(areaPath);

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);

        int existingConsts = await conn.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM RHODES_BANKING_SILA.C_CONSTANTS_TB WHERE CONSTANT_MAIN_ID = :mainId",
            new { mainId = AreaMainId });
        int existingLangs = await conn.ExecuteScalarAsync<int>(
            "SELECT COUNT(1) FROM RHODES_BANKING_SILA.C_CONSTANT_LANGS_TB WHERE CONSTANT_MAIN_ID = :mainId",
            new { mainId = AreaMainId });

        report.Stats["sourceValidRows"] = valid;
        report.Stats["sourceSkippedRows"] = skipped;
        report.Stats["duplicateOldAreaCodes"] = dupOldArea;
        report.Stats["existingConstantsMain10"] = existingConsts;
        report.Stats["existingLangsMain10"] = existingLangs;
        report.Stats["willDeleteThenInsertConstants"] = valid;
        report.Stats["willInsertLangRows"] = valid * 2;
        report.Stats["areaExcelPath"] = areaPath;

        ReportBuilder.AddIssue(report, "Warning", "FULL_REPLACE",
            $"Run will DELETE all CONSTANT_MAIN_ID={AreaMainId} ({existingConsts} constants, {existingLangs} langs) then insert {valid} areas.",
            existingConsts + existingLangs);

        ReportBuilder.AddIssue(report, "Warning", "DUP_OLD_AREA_CODE",
            "Duplicate OLD_AREA_CODE rows skipped (first wins).", dupOldArea, samples);

        ReportBuilder.AddIssue(report, "Info", "SKIPPED_ROWS",
            "Invalid/placeholder area rows that will be skipped.", skipped);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(AreaMigrationRequest? request = null, CancellationToken ct = default)
    {
        var areaPath = paths.ResolveAreaExcelPath(request?.AreaExcelPath);
        if (!File.Exists(areaPath))
            throw new FileNotFoundException($"Area Excel not found: {areaPath}");

        using var wb = new XLWorkbook(areaPath);
        var ws = wb.Worksheets.First();
        var h = BuildHeaderIndex(ws.Row(1));
        int lastRow = ws.LastRowUsed()!.RowNumber();

        RequireCol(h, "AREA_CODE");
        RequireCol(h, "NEW_CITY_CODE");
        RequireCol(h, "AREA_DESC_A");
        RequireCol(h, "AREA_DESC_E");
        RequireCol(h, "OLD_AREA_CODE");

        var constants = new List<CConstantsTb>();
        var langs = new List<CConstantLangsTb>();
        var seenIds = new HashSet<int>();
        int skipped = 0;
        var now = DateTime.Now;
        var log = new List<string>();

        for (int r = 2; r <= lastRow; r++)
        {
            var row = ws.Row(r);
            int? newCityId = GetInt(row, h, "NEW_CITY_CODE");
            int? oldAreaId = GetInt(row, h, "OLD_AREA_CODE");
            string? descAr = GetString(row, h, "AREA_DESC_A");
            string? descEn = GetString(row, h, "AREA_DESC_E");

            if (oldAreaId is null or <= 0 || newCityId is null or < 0 ||
                string.IsNullOrWhiteSpace(descAr) || descAr == "-")
            {
                skipped++;
                continue;
            }

            if (!seenIds.Add(oldAreaId.Value))
            {
                skipped++;
                log.Add($"[WARN] Row {r}: duplicate OLD_AREA_CODE={oldAreaId}");
                continue;
            }

            string ar = T(descAr, 200)!;
            string en = T(descEn ?? descAr, 200)!;

            constants.Add(new CConstantsTb
            {
                ConstantMainId = AreaMainId,
                ConstantId = oldAreaId.Value,
                ConstantDesc = ar,
                ParentConstantId = newCityId.Value,
                IsUpdatable = false,
                IsHidden = false,
                CreatedOn = now,
            });
            langs.Add(new CConstantLangsTb
            {
                ConstantMainId = AreaMainId,
                ConstantId = oldAreaId.Value,
                LangId = "ar",
                ConstantDesc = ar,
                CreatedOn = now,
            });
            langs.Add(new CConstantLangsTb
            {
                ConstantMainId = AreaMainId,
                ConstantId = oldAreaId.Value,
                LangId = "en",
                ConstantDesc = en,
                CreatedOn = now,
            });
        }

        await using var ctx = new SilaDbContext(dbOptions);

        int deletedLangs = await ctx.CConstantLangsTbs
            .Where(x => x.ConstantMainId == AreaMainId).ExecuteDeleteAsync(ct);
        int deletedConsts = await ctx.CConstantsTbs
            .Where(x => x.ConstantMainId == AreaMainId).ExecuteDeleteAsync(ct);

        ctx.CConstantsTbs.AddRange(constants);
        ctx.CConstantLangsTbs.AddRange(langs);
        await ctx.SaveChangesAsync(ct);

        log.Add($"Deleted MAIN_ID={AreaMainId}: {deletedConsts} constants, {deletedLangs} langs");
        log.Add($"Inserted {constants.Count} areas + {langs.Count} langs (skipped {skipped})");

        return new PhaseRunResultDto
        {
            Phase = "areas",
            Success = true,
            Message = $"Replaced MAIN_ID={AreaMainId}: {constants.Count} areas inserted, {skipped} skipped.",
            Stats =
            {
                ["deletedConstants"] = deletedConsts,
                ["deletedLangs"] = deletedLangs,
                ["insertedConstants"] = constants.Count,
                ["insertedLangs"] = langs.Count,
                ["skipped"] = skipped,
            },
            Log = log,
        };
    }

    static (int Valid, int Skipped, int DupOldArea, List<string> DupSamples) AnalyzeAreaExcel(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var h = BuildHeaderIndex(ws.Row(1));
        RequireCol(h, "OLD_AREA_CODE");
        RequireCol(h, "NEW_CITY_CODE");
        RequireCol(h, "AREA_DESC_A");

        int last = ws.LastRowUsed()!.RowNumber();
        var seen = new HashSet<int>();
        int valid = 0, skipped = 0, dups = 0;
        var samples = new List<string>();

        for (int r = 2; r <= last; r++)
        {
            var row = ws.Row(r);
            int? newCityId = GetInt(row, h, "NEW_CITY_CODE");
            int? oldAreaId = GetInt(row, h, "OLD_AREA_CODE");
            string? descAr = GetString(row, h, "AREA_DESC_A");

            if (oldAreaId is null or <= 0 || newCityId is null or < 0 ||
                string.IsNullOrWhiteSpace(descAr) || descAr == "-")
            {
                skipped++;
                continue;
            }

            if (!seen.Add(oldAreaId.Value))
            {
                dups++;
                skipped++;
                if (samples.Count < SampleLimit)
                    samples.Add($"Row {r} OLD_AREA_CODE={oldAreaId}");
                continue;
            }

            valid++;
        }

        return (valid, skipped, dups, samples);
    }
}
