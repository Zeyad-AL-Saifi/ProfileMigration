using ProfileMigration.Application.Dtos;

namespace ProfileMigration.Application;

public static class ReportBuilder
{
    public static MigrationIssueDto Issue(
        string severity, string code, string message, int count, IEnumerable<string>? samples = null)
        => new()
        {
            Severity = severity,
            Code = code,
            Message = message,
            Count = count,
            Samples = samples?.Take(Excel.ExcelHelpers.SampleLimit).ToList() ?? [],
        };

    public static void AddIssue(
        MigrationReportDto report,
        string severity, string code, string message, int count, IEnumerable<string>? samples = null)
    {
        if (count <= 0) return;
        report.Issues.Add(Issue(severity, code, message, count, samples));
        if (severity.Equals("Error", StringComparison.OrdinalIgnoreCase))
            report.CanProceed = false;
    }
}
