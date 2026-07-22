using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using ProfileMigration.Application.Profiles;
using ProfileMigration.DAL.Models;
using Xunit;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Tests;

public sealed class ProfileExcelSourceTests
{
    [Fact]
    public void ValidateRequiredSheetsAndColumns_AcceptsDistributedSourcesWithRowOneHeaders()
    {
        using var clientWb = CreateClientWorkbook();
        using var idCardWb = CreateIdCardWorkbook();
        using var addressWb = CreateAddressWorkbook();

        var errors = ValidateRequiredSheetsAndColumns(clientWb, idCardWb, addressWb);

        Assert.Empty(errors);
    }

    [Fact]
    public void ValidateRequiredSheetsAndColumns_IdentifiesSourceAndMissingColumn()
    {
        using var clientWb = CreateClientWorkbook();
        using var idCardWb = CreateIdCardWorkbook(includeCardNumber: false);
        using var addressWb = CreateAddressWorkbook();

        var errors = ValidateRequiredSheetsAndColumns(clientWb, idCardWb, addressWb);

        var error = Assert.Single(errors);
        Assert.Equal("idCards", error.Source);
        Assert.Equal("column", error.Kind);
        Assert.Equal($"{IdCardSheet}.ID_CARD_NO", error.Name);
    }

    [Fact]
    public void ValidateSourceFiles_IdentifiesMissingReferenceFile()
    {
        using var files = new TemporaryProfileFiles();
        files.SaveClient(CreateClientWorkbook());
        files.SaveAddress(CreateAddressWorkbook());

        var errors = ValidateSourceFiles(files.ClientPath, files.IdCardPath, files.AddressPath);

        var error = Assert.Single(errors);
        Assert.Equal("idCards", error.Source);
        Assert.Equal("file", error.Kind);
        Assert.Equal(files.IdCardPath, error.Name);
    }

    [Fact]
    public void Open_JoinsSeparateSourcesByCompanyAndClientId()
    {
        using var files = new TemporaryProfileFiles();
        files.SaveClient(CreateClientWorkbook());
        files.SaveIdCard(CreateIdCardWorkbook());
        files.SaveAddress(CreateAddressWorkbook());

        using var loaded = Open(files.ClientPath, files.IdCardPath, files.AddressPath);
        var eligibility = BuildEligibility(loaded);
        var source = Assert.Single(BuildEligibleRows(loaded, eligibility: eligibility));

        Assert.Equal(2, loaded.FirstClientDataRow);
        Assert.True(loaded.IdCards.ContainsKey(("ASALA", 101)));
        Assert.True(loaded.Addresses.ContainsKey(("ASALA", 101)));
        Assert.True(eligibility.IsEligible("ASALA", 101));
        Assert.Equal("ABC123", source.IdNum);
        Assert.Equal(611300001906L, source.Profile.CustId);
        Assert.Equal("عينة", source.Profile.FirstNameNa);
        Assert.Equal("Sample", source.Profile.FirstNameFo);
        Assert.Equal("عينة اسم عربي عائلة", source.Profile.ProfileNameNa);
        Assert.Equal("Sample English Test Family", source.Profile.ProfileNameFo);
        Assert.Equal("عينة اسم عربي عائلة", source.Profile.ProfileNameFNa);
        Assert.Equal("Sample English Test Family", source.Profile.ProfileNameFFo);
        Assert.Equal("Street 1 Home", source.Profile.PermanentAddress);
        Assert.Equal(["0790000000"], source.Phones);
    }

    [Fact]
    public void ArrayBindings_UseEfColumnsAndValueConverters()
    {
        var options = new DbContextOptionsBuilder<SilaDbContext>()
            .UseOracle("User Id=test;Password=test;Data Source=localhost/test")
            .Options;
        using var context = new SilaDbContext(options);
        var entityType = context.Model.FindEntityType(typeof(ProfilesTb))!;
        var storeObject = StoreObjectIdentifier.Table(
            entityType.GetTableName()!,
            entityType.GetSchema());
        var profile = new ProfilesTb
        {
            ProfileId = 10,
            CustId = 611300001906L,
            BranchId = 6,
            ProfileNameNa = "اسم تجريبي",
            ProfileNameFo = "Test Name",
            CustTypeId = true,
            GenderId = true,
            EntrySourceId = false,
            CustStatusId = true,
            ResidentFlag = true,
            ReviewFlag = true,
            MailAddressFlag = false,
            PermanentAddressFlag = true,
        };

        var bindings = MigrationBatchInserter
            .BuildBindings(entityType, storeObject, new[] { profile })
            .ToDictionary(x => x.ColumnName);

        Assert.Equal("اسم تجريبي", bindings["PROFILE_NAME_NA"].Values[0]);
        Assert.Equal("Test Name", bindings["PROFILE_NAME_FO"].Values[0]);
        Assert.Equal(2, bindings["CUST_TYPE_ID"].Values[0]);
        Assert.Equal(2, bindings["GENDER_ID"].Values[0]);
        Assert.Equal(1, bindings["ENTRY_SOURCE_ID"].Values[0]);
        Assert.Equal(1, bindings["CUST_STATUS_ID"].Values[0]);
        Assert.DoesNotContain("ACTIVITY_ID", bindings.Keys);
    }

    [Fact]
    public void BranchMapping_UsesDefaultOneAndJeninTen()
    {
        var branchMap = LoadBranchIdMap();

        Assert.Equal(1, DefaultBranchId);
        Assert.Equal(1, branchMap[("ACAD", 0)]);
        Assert.Equal(10, branchMap[("ACAD", 80)]);
        Assert.Equal(1, branchMap[("ACAD", 100)]);
        Assert.Equal(1, branchMap[("ACAD", 120)]);
        Assert.Equal(1, branchMap[("ACAD", 130)]);
        Assert.Equal(1, branchMap[("ACAD", 200)]);
        Assert.Equal(1, branchMap[("ASALA", 0)]);
        Assert.Equal(10, branchMap[("ASALA", 20)]);
    }

    static XLWorkbook CreateClientWorkbook()
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(ClientSheet);
        SetHeaders(
            worksheet,
            "CLIENT_ID", "COMPANY", "Combined Number",
            "NAME1", "NAME2", "NAME3", "NAME4",
            "ANAME1", "ANAME2", "ANAME3", "ANAME4");
        worksheet.Cell(2, 1).Value = 101;
        worksheet.Cell(2, 2).Value = "ASALA";
        worksheet.Cell(2, 3).Value = 611300001906L;
        worksheet.Cell(2, 4).Value = "Sample";
        worksheet.Cell(2, 5).Value = "English";
        worksheet.Cell(2, 6).Value = "Test";
        worksheet.Cell(2, 7).Value = "Family";
        worksheet.Cell(2, 8).Value = "عينة";
        worksheet.Cell(2, 9).Value = "اسم";
        worksheet.Cell(2, 10).Value = "عربي";
        worksheet.Cell(2, 11).Value = "عائلة";
        return workbook;
    }

    static XLWorkbook CreateIdCardWorkbook(bool includeCardNumber = true)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(IdCardSheet);
        var headers = includeCardNumber
            ? new[] { "CLIENT_ID", "COMPANY", "ID_CARD_TYPE", "ID_CARD_NO" }
            : ["CLIENT_ID", "COMPANY", "ID_CARD_TYPE"];
        SetHeaders(worksheet, headers);
        worksheet.Cell(2, 1).Value = 101;
        worksheet.Cell(2, 2).Value = "ASALA";
        worksheet.Cell(2, 3).Value = 1;
        if (includeCardNumber)
            worksheet.Cell(2, 4).Value = "ABC123";
        return workbook;
    }

    static XLWorkbook CreateAddressWorkbook()
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(AddressSheet);
        SetHeaders(
            worksheet,
            "CLIENT_ID", "COMPANY", "CITY_CODE", "AREA_CODE",
            "STREET", "ADDRESS_DESC", "TEL_NO", "MAIN_FLAG");
        worksheet.Cell(2, 1).Value = 101;
        worksheet.Cell(2, 2).Value = "ASALA";
        worksheet.Cell(2, 3).Value = 1;
        worksheet.Cell(2, 4).Value = 10;
        worksheet.Cell(2, 5).Value = "Street 1";
        worksheet.Cell(2, 6).Value = "Home";
        worksheet.Cell(2, 7).Value = "0790000000";
        worksheet.Cell(2, 8).Value = 1;
        return workbook;
    }

    static void SetHeaders(IXLWorksheet worksheet, params string[] headers)
    {
        for (var column = 1; column <= headers.Length; column++)
            worksheet.Cell(1, column).Value = headers[column - 1];
    }

    sealed class TemporaryProfileFiles : IDisposable
    {
        readonly string directory = Path.Combine(
            Path.GetTempPath(), $"profile-excel-tests-{Guid.NewGuid():N}");

        public TemporaryProfileFiles() => Directory.CreateDirectory(directory);

        public string ClientPath => Path.Combine(directory, "clients.xlsx");
        public string IdCardPath => Path.Combine(directory, "id-cards.xlsx");
        public string AddressPath => Path.Combine(directory, "addresses.xlsx");

        public void SaveClient(XLWorkbook workbook) => Save(workbook, ClientPath);
        public void SaveIdCard(XLWorkbook workbook) => Save(workbook, IdCardPath);
        public void SaveAddress(XLWorkbook workbook) => Save(workbook, AddressPath);

        static void Save(XLWorkbook workbook, string path)
        {
            using (workbook)
                workbook.SaveAs(path);
        }

        public void Dispose()
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }
}
