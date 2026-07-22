using System.Data;
using Dapper;
using Microsoft.EntityFrameworkCore;
using ProfileMigration.Application.Data;
using ProfileMigration.Application.Dtos;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application.Constants;

/// <summary>
/// Seeds / fixes profile ID-type constants in C_CONSTANTS_TB (MAIN_ID=1):
/// - CONSTANT_ID=1 description → هوية
/// - CONSTANT_ID=5 → جواز سفر اردني (insert if missing)
/// </summary>
public sealed class ConstantMigrationService(
    DbContextOptions<SilaDbContext> dbOptions,
    IOracleConnectionFactory connectionFactory)
{
    const int MainId = 1;
    const int IdentityId = 1;
    const int JordanianPassportId = 5;
    const string IdentityDesc = "هوية";
    const string PassportDesc = "جواز سفر اردني";

    public async Task<MigrationReportDto> ValidateAsync(CancellationToken ct = default)
    {
        var report = new MigrationReportDto { Phase = "constants", CanProceed = true };

        using var conn = await connectionFactory.CreateOpenConnectionAsync(ct);
        var constantsTable = connectionFactory.QualifyTable("C_CONSTANTS_TB");

        var identity = await conn.QuerySingleOrDefaultAsync<(int? ConstantId, string? ConstantDesc)>(
            $"""
            SELECT CONSTANT_ID AS ConstantId, CONSTANT_DESC AS ConstantDesc
            FROM {constantsTable}
            WHERE CONSTANT_MAIN_ID = :mainId AND CONSTANT_ID = :constId
            """,
            new { mainId = MainId, constId = IdentityId });

        int passportCount = await conn.ExecuteScalarAsync<int>(
            $"""
            SELECT COUNT(1) FROM {constantsTable}
            WHERE CONSTANT_MAIN_ID = :mainId AND CONSTANT_ID = :constId
            """,
            new { mainId = MainId, constId = JordanianPassportId });

        bool identityExists = identity.ConstantId.HasValue;
        bool identityNeedsUpdate = identityExists
            && !string.Equals(identity.ConstantDesc?.Trim(), IdentityDesc, StringComparison.Ordinal);
        bool passportMissing = passportCount == 0;

        report.Stats["identityExists"] = identityExists;
        report.Stats["identityCurrentDesc"] = identity.ConstantDesc;
        report.Stats["identityWillUpdate"] = identityNeedsUpdate;
        report.Stats["identityWillInsert"] = !identityExists;
        report.Stats["jordanianPassportExists"] = !passportMissing;
        report.Stats["jordanianPassportWillInsert"] = passportMissing;

        if (!identityExists)
            ReportBuilder.AddIssue(report, "Warning", "IDENTITY_MISSING",
                $"Identity constant (MAIN_ID={MainId}, ID={IdentityId}) is missing and will be inserted as '{IdentityDesc}'.", 1);
        else if (identityNeedsUpdate)
            ReportBuilder.AddIssue(report, "Info", "IDENTITY_DESC_UPDATE",
                $"Identity constant desc will be updated from '{identity.ConstantDesc}' to '{IdentityDesc}'.", 1);

        if (passportMissing)
            ReportBuilder.AddIssue(report, "Info", "PASSPORT_CONST_MISSING",
                $"Jordanian passport constant (MAIN_ID={MainId}, ID={JordanianPassportId}) will be inserted.", 1);
        else
            ReportBuilder.AddIssue(report, "Info", "PASSPORT_CONST_EXISTS",
                "Jordanian passport constant already exists — will skip insert.", 1);

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(CancellationToken ct = default)
    {
        var log = new List<string>();
        var constantsTable = connectionFactory.QualifyTable("C_CONSTANTS_TB");
        await using var ctx = new SilaDbContext(dbOptions);

        var conn = ctx.Database.GetDbConnection();
        if (conn.State != ConnectionState.Open)
            await conn.OpenAsync(ct);

        bool identityUpdated = false, identityInserted = false, passportInserted = false;

        // ── Identity (1,1) ──────────────────────────────────────────────────
        string? currentDesc = null;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"""
                SELECT CONSTANT_DESC FROM {constantsTable}
                WHERE CONSTANT_MAIN_ID = 1 AND CONSTANT_ID = 1
                """;
            currentDesc = (await cmd.ExecuteScalarAsync(ct)) as string;
        }

        if (currentDesc is null)
        {
            await ctx.Database.ExecuteSqlRawAsync($"""
                INSERT INTO {constantsTable}
                    (CONSTANT_MAIN_ID, CONSTANT_ID, CONSTANT_DESC, PARENT_CONSTANT_ID,
                     PMA_CODE, IS_UPDATABLE, IS_HIDDEN, CREATED_ON)
                VALUES (1, 1, N'هوية', NULL, N'N', 0, 0, SYSDATE)
                """, ct);
            identityInserted = true;
            log.Add("[ADD] C_CONSTANTS_TB (1,1) هوية");
        }
        else if (!string.Equals(currentDesc.Trim(), IdentityDesc, StringComparison.Ordinal))
        {
            await ctx.Database.ExecuteSqlRawAsync($"""
                UPDATE {constantsTable}
                SET CONSTANT_DESC = N'هوية'
                WHERE CONSTANT_MAIN_ID = 1 AND CONSTANT_ID = 1
                """, ct);
            identityUpdated = true;
            log.Add($"[UPD] C_CONSTANTS_TB (1,1) '{currentDesc}' → '{IdentityDesc}'");
        }
        else
        {
            log.Add("[SKIP] C_CONSTANTS_TB (1,1) هوية already correct");
        }

        // ── Jordanian passport (1,5) ────────────────────────────────────────
        int passportCount;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"""
                SELECT COUNT(1) FROM {constantsTable}
                WHERE CONSTANT_MAIN_ID = 1 AND CONSTANT_ID = 5
                """;
            passportCount = Convert.ToInt32(await cmd.ExecuteScalarAsync(ct));
        }

        if (passportCount == 0)
        {
            await ctx.Database.ExecuteSqlRawAsync($"""
                INSERT INTO {constantsTable}
                    (CONSTANT_MAIN_ID, CONSTANT_ID, CONSTANT_DESC, PARENT_CONSTANT_ID,
                     PMA_CODE, IS_UPDATABLE, IS_HIDDEN, CREATED_ON)
                VALUES (1, 5, N'جواز سفر اردني', 2, N'N', 0, 0, DATE '2023-02-08')
                """, ct);
            passportInserted = true;
            log.Add("[ADD] C_CONSTANTS_TB (1,5) جواز سفر اردني");
        }
        else
        {
            log.Add("[SKIP] C_CONSTANTS_TB (1,5) Jordanian passport already exists");
        }

        return new PhaseRunResultDto
        {
            Phase = "constants",
            Success = true,
            Message = $"Constants: identity inserted={identityInserted}, updated={identityUpdated}; passport inserted={passportInserted}.",
            Stats =
            {
                ["identityInserted"] = identityInserted,
                ["identityUpdated"] = identityUpdated,
                ["jordanianPassportInserted"] = passportInserted,
            },
            Log = log,
        };
    }
}
