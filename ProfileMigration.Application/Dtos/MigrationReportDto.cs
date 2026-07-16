namespace ProfileMigration.Application.Dtos;

/// <summary>Clients Excel path override (profiles / works / banks / partners / contacts).</summary>
public sealed class ExcelMigrationRequest
{
    public string? ExcelPath { get; set; }
}

/// <summary>Areas Excel path override (areas phase only).</summary>
public sealed class AreaMigrationRequest
{
    public string? AreaExcelPath { get; set; }
}

public sealed class MigrationReportDto
{
    public string Phase { get; set; } = "";
    public bool CanProceed { get; set; } = true;
    public Dictionary<string, object?> Stats { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<MigrationIssueDto> Issues { get; set; } = [];
    public Dictionary<string, object?> Breakdown { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class MigrationIssueDto
{
    public string Severity { get; set; } = "Warning"; // Error | Warning | Info
    public string Code { get; set; } = "";
    public string Message { get; set; } = "";
    public int Count { get; set; }
    public List<string> Samples { get; set; } = [];
}

public sealed class PhaseRunResultDto
{
    public string Phase { get; set; } = "";
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public Dictionary<string, object?> Stats { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> Log { get; set; } = [];
}
