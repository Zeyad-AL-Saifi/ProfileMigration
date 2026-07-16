namespace ProfileMigration.Application.Options;

public sealed class MigrationOptions
{
    public const string SectionName = "Migration";

    public string ConnectionString { get; set; } = "";
    public string ExcelFilePath { get; set; } = "Match MF_CLIENT - Copy (3).xlsx";
    public string IdCardExcelFilePath { get; set; } = "March MF_CLINT_ID_CARD.xlsx";
    public string AddressExcelFilePath { get; set; } = "MF_CLIENT_ADDRESSES.xlsx";
    public string AreaExcelFilePath { get; set; } = "area 2.xlsx";

    /// <summary>Optional override of base directory for relative Excel paths.</summary>
    public string? ContentRoot { get; set; }
}
