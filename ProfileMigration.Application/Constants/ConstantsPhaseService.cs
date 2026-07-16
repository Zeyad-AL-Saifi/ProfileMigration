using ProfileMigration.Application.Areas;
using ProfileMigration.Application.Branches;
using ProfileMigration.Application.Dtos;

namespace ProfileMigration.Application.Constants;

/// <summary>
/// Single constants phase: branches seed + ID-type constants + areas replace.
/// </summary>
public sealed class ConstantsPhaseService(
    BranchMigrationService branches,
    ConstantMigrationService idTypeConstants,
    AreaMigrationService areas)
{
    public async Task<MigrationReportDto> ValidateAsync(AreaMigrationRequest? request = null, CancellationToken ct = default)
    {
        var branchesReport = await branches.ValidateAsync(ct);
        var idTypesReport = await idTypeConstants.ValidateAsync(ct);
        var areasReport = await areas.ValidateAsync(request, ct);

        var report = new MigrationReportDto
        {
            Phase = "constants",
            CanProceed = branchesReport.CanProceed && idTypesReport.CanProceed && areasReport.CanProceed,
        };

        PrefixAndMerge(report, "branches", branchesReport);
        PrefixAndMerge(report, "idTypes", idTypesReport);
        PrefixAndMerge(report, "areas", areasReport);

        report.Breakdown["branches"] = branchesReport.Stats;
        report.Breakdown["idTypes"] = idTypesReport.Stats;
        report.Breakdown["areas"] = areasReport.Stats;

        return report;
    }

    public async Task<PhaseRunResultDto> RunAsync(AreaMigrationRequest? request = null, CancellationToken ct = default)
    {
        var log = new List<string>();

        var branchesResult = await branches.RunAsync(ct);
        log.Add($"--- branches ---");
        log.AddRange(branchesResult.Log);

        var idTypesResult = await idTypeConstants.RunAsync(ct);
        log.Add($"--- idTypes ---");
        log.AddRange(idTypesResult.Log);

        var areasResult = await areas.RunAsync(request, ct);
        log.Add($"--- areas ---");
        log.AddRange(areasResult.Log);

        bool success = branchesResult.Success && idTypesResult.Success && areasResult.Success;

        return new PhaseRunResultDto
        {
            Phase = "constants",
            Success = success,
            Message = $"Constants phase: branches + idTypes + areas. Success={success}.",
            Stats =
            {
                ["branches"] = branchesResult.Stats,
                ["idTypes"] = idTypesResult.Stats,
                ["areas"] = areasResult.Stats,
            },
            Log = log,
        };
    }

    static void PrefixAndMerge(MigrationReportDto target, string prefix, MigrationReportDto source)
    {
        foreach (var (k, v) in source.Stats)
            target.Stats[$"{prefix}.{k}"] = v;

        foreach (var issue in source.Issues)
        {
            target.Issues.Add(new MigrationIssueDto
            {
                Severity = issue.Severity,
                Code = $"{prefix.ToUpperInvariant()}_{issue.Code}",
                Message = $"[{prefix}] {issue.Message}",
                Count = issue.Count,
                Samples = issue.Samples,
            });
        }
    }
}
