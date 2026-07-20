using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using ProfileMigration.Application;
using ProfileMigration.Application.Analysis;
using ProfileMigration.Application.Options;
using ProfileMigration.Application.Profiles;
using Xunit;

namespace ProfileMigration.Tests;

public sealed class ClientAnalysisServiceTests
{
    [Fact]
    public void Analyze_InvalidClientId_ReportsRowCompanyValueAndReason()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"analysis-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            string clientsPath = Path.Combine(directory, "clients.xlsx");
            string cardsPath = Path.Combine(directory, "cards.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Match MF_CLIENT");
                SetHeaders(sheet, "COMPANY", "CLIENT_ID");
                sheet.Cell(2, 1).Value = "ASALA";
                sheet.Cell(2, 2).Value = "ABC-123";
                workbook.SaveAs(clientsPath);
            }
            SaveCards(cardsPath);

            var resolver = new MigrationPathResolver(Options.Create(new MigrationOptions
            {
                ExcelFilePath = clientsPath,
                IdCardExcelFilePath = cardsPath,
                ContentRoot = directory,
            }));

            var error = Assert.Throws<InvalidDataException>(
                () => new ClientAnalysisService(resolver).Analyze(clientsPath));

            Assert.Contains("row 2", error.Message);
            Assert.Contains("COMPANY='ASALA'", error.Message);
            Assert.Contains("value='ABC-123'", error.Message);
            Assert.Contains("not a valid integer", error.Message);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Analyze_UsesSameDuplicatePopulationAsMigrationEligibility()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"analysis-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            string clientsPath = Path.Combine(directory, "clients.xlsx");
            string cardsPath = Path.Combine(directory, "cards.xlsx");
            SaveClients(clientsPath);
            SaveCards(cardsPath);

            var resolver = new MigrationPathResolver(Options.Create(new MigrationOptions
            {
                ExcelFilePath = clientsPath,
                IdCardExcelFilePath = cardsPath,
                AddressExcelFilePath = Path.Combine(directory, "unused-addresses.xlsx"),
                ContentRoot = directory,
            }));
            var service = new ClientAnalysisService(resolver);

            var response = service.Analyze(clientsPath);
            var summary = Assert.IsType<Dictionary<string, object?>>(response["summary"]);
            var duplicates = Assert.IsType<Dictionary<string, object?>>(response["duplicates"]);
            var asalaGroups = Assert.IsType<List<Dictionary<string, object?>>>(duplicates["internal_asala"]);

            Assert.Equal(5, summary["total_clients"]);
            Assert.Equal(2, duplicates["total_internal_duplicate_records"]);
            Assert.Single(asalaGroups);
            Assert.Equal(1, summary["cross_company_matched"]);

            var eligibility = ClientEligibilityClassifier.Classify(
            [
                new("ASALA", 1, "1|INTERNAL", "INTERNAL", 1),
                new("ASALA", 2, "1|INTERNAL", "INTERNAL", 1),
                new("ASALA", 3, "1|CROSS", "CROSS", 1),
                new("ACAD", 4, "1|CROSS", "CROSS", 1),
                new("ACAD", 5, "1|UNIQUE", "UNIQUE", 1),
            ]);

            Assert.Equal(2, eligibility.SkippedInternalDuplicates);
            Assert.Equal(2, eligibility.SkippedCrossCompanyMatches);
            Assert.Equal(1, eligibility.EligibleCount);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    static void SaveClients(string path)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Match MF_CLIENT");
        SetHeaders(sheet, "COMPANY", "CLIENT_ID", "NAME1");
        AddClient(sheet, 2, "ASALA", 1, "A");
        AddClient(sheet, 3, "ASALA", 2, "B");
        AddClient(sheet, 4, "ASALA", 3, "C");
        AddClient(sheet, 5, "ACAD", 4, "D");
        AddClient(sheet, 6, "ACAD", 5, "E");
        workbook.SaveAs(path);
    }

    static void SaveCards(string path)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Match MF_CLIENT_ID_CARD");
        SetHeaders(sheet, "COMPANY", "CLIENT_ID", "ID_CARD_TYPE", "ID_CARD_NO");
        AddCard(sheet, 2, "ASALA", 1, "INTERNAL");
        AddCard(sheet, 3, "ASALA", 2, "INTERNAL");
        AddCard(sheet, 4, "ASALA", 3, "CROSS");
        AddCard(sheet, 5, "ACAD", 4, "CROSS");
        AddCard(sheet, 6, "ACAD", 5, "UNIQUE");
        workbook.SaveAs(path);
    }

    static void AddClient(IXLWorksheet sheet, int row, string company, int clientId, string name)
    {
        sheet.Cell(row, 1).Value = company;
        sheet.Cell(row, 2).Value = clientId;
        sheet.Cell(row, 3).Value = name;
    }

    static void AddCard(IXLWorksheet sheet, int row, string company, int clientId, string cardNo)
    {
        sheet.Cell(row, 1).Value = company;
        sheet.Cell(row, 2).Value = clientId;
        sheet.Cell(row, 3).Value = 1;
        sheet.Cell(row, 4).Value = cardNo;
    }

    static void SetHeaders(IXLWorksheet sheet, params string[] headers)
    {
        for (int i = 0; i < headers.Length; i++)
            sheet.Cell(1, i + 1).Value = headers[i];
    }
}
