using ClosedXML.Excel;
using ProfileMigration.DAL.Models;
using static ProfileMigration.Application.Excel.ExcelHelpers;

namespace ProfileMigration.Application.Profiles;

/// <summary>
/// Shared Excel loaders / mappers for profile-related migration phases.
/// </summary>
public static class ProfileExcelSource
{
    public const string ClientSheet = "Match MF_CLIENT";
    public const string IdCardSheet = "Match MF_CLIENT_ID_CARD";
    public const string AddressSheet = "Match MF_CLIENT_ADDRESSES";
    public const int ClientHeaderRow = 1;
    public const int IdCardHeaderRow = 1;
    public const int AddressHeaderRow = 1;

    public const int DefaultBranchId = 1;
    public const int DefaultBankId = 1;
    public const byte DefaultAccountCurrId = 1;

    public sealed record IdCardData
    {
        /// <summary>Raw Excel ID_CARD_TYPE (before DB remap) — used for analysis card_key.</summary>
        public int? RawIdCardType { get; init; }
        /// <summary>Raw Excel ID_CARD_NO — used for analysis card_key.</summary>
        public string? RawIdCardNo { get; init; }
        /// <summary>Mapped ID type for PROFILES_TB (e.g. 4 → 5 Jordanian passport).</summary>
        public int? IdTypeId { get; init; }
        public string? IdNum { get; init; }
        public DateTime? IssueDate { get; init; }
        public string? IssuePlace { get; init; }
        public DateTime? ExpiryDate { get; init; }

        public string? CardKey => ClientEligibilityClassifier.BuildCardKey(RawIdCardType, RawIdCardNo);
    }

    public sealed record AddressData(
        int? PermanentStateId,
        int? PermanentCityId,
        string? PermanentAddress,
        string? MailZip,
        string? MailPobox,
        IReadOnlyList<string> Phones);

    public sealed record EligibleSourceRow(
        int RowNumber,
        long ClientId,
        string Company,
        string IdNum,
        int IdType,
        ProfilesTb Profile,
        ProfileWorksTb? Work,
        ProfilesPartnersTb? Partner,
        ProfileBankInformationTb? Bank,
        IReadOnlyList<string> Phones);

    public sealed record ValidationError(string Source, string Kind, string Name);

    public sealed class LoadedWorkbook : IDisposable
    {
        public required XLWorkbook ClientWorkbook { get; init; }
        public required XLWorkbook IdCardWorkbook { get; init; }
        public required XLWorkbook AddressWorkbook { get; init; }
        public required Dictionary<string, int> ClientHeaders { get; init; }
        public required IXLWorksheet ClientSheetWs { get; init; }
        public required int FirstClientDataRow { get; init; }
        public required int LastClientRow { get; init; }
        public required Dictionary<(string Company, long ClientId), IdCardData> IdCards { get; init; }
        public required Dictionary<(string Company, long ClientId), AddressData> Addresses { get; init; }
        public required Dictionary<(string Company, int OldCode), int> BranchMap { get; init; }
        public int UnmappedCityCount { get; init; }

        public void Dispose()
        {
            ClientWorkbook.Dispose();
            IdCardWorkbook.Dispose();
            AddressWorkbook.Dispose();
        }
    }

    public static LoadedWorkbook Open(string clientPath, string idCardPath, string addressPath)
    {
        EnsureFileExists(clientPath, "clients");
        EnsureFileExists(idCardPath, "id cards");
        EnsureFileExists(addressPath, "addresses");

        XLWorkbook? clientWb = null;
        XLWorkbook? idCardWb = null;
        XLWorkbook? addressWb = null;
        try
        {
            clientWb = new XLWorkbook(clientPath);
            idCardWb = new XLWorkbook(idCardPath);
            addressWb = new XLWorkbook(addressPath);

            var validationErrors = ValidateRequiredSheetsAndColumns(clientWb, idCardWb, addressWb);
            if (validationErrors.Count > 0)
            {
                var details = string.Join(", ", validationErrors.Select(
                    e => $"{e.Source}.{e.Kind}:{e.Name}"));
                throw new InvalidDataException($"Invalid profile Excel sources: {details}");
            }

            var clientWs = GetRequiredSheet(clientWb, ClientSheet);
            var cityIdMap = LoadCityIdMap();
            var addresses = LoadAddresses(addressWb, cityIdMap, out int unmappedCity);

            return new LoadedWorkbook
            {
                ClientWorkbook = clientWb,
                IdCardWorkbook = idCardWb,
                AddressWorkbook = addressWb,
                ClientSheetWs = clientWs,
                ClientHeaders = BuildHeaderIndex(clientWs.Row(ClientHeaderRow)),
                FirstClientDataRow = ClientHeaderRow + 1,
                LastClientRow = clientWs.LastRowUsed()?.RowNumber() ?? ClientHeaderRow,
                IdCards = LoadIdCards(idCardWb),
                Addresses = addresses,
                BranchMap = LoadBranchIdMap(),
                UnmappedCityCount = unmappedCity,
            };
        }
        catch
        {
            clientWb?.Dispose();
            idCardWb?.Dispose();
            addressWb?.Dispose();
            throw;
        }
    }

    public static List<ValidationError> ValidateRequiredSheetsAndColumns(
        XLWorkbook clientWb,
        XLWorkbook idCardWb,
        XLWorkbook addressWb)
    {
        var errors = new List<ValidationError>();
        ValidateSource(clientWb, "clients", ClientSheet, ClientHeaderRow,
            ["CLIENT_ID", "COMPANY", "NAME1", "Combined Number"], errors);
        ValidateSource(idCardWb, "idCards", IdCardSheet, IdCardHeaderRow,
            ["CLIENT_ID", "COMPANY", "ID_CARD_TYPE", "ID_CARD_NO"], errors);
        ValidateSource(addressWb, "addresses", AddressSheet, AddressHeaderRow,
            ["CLIENT_ID", "COMPANY", "CITY_CODE", "AREA_CODE"], errors);

        return errors;
    }

    public static List<ValidationError> ValidateSourceFiles(
        string clientPath,
        string idCardPath,
        string addressPath)
    {
        var files = new[]
        {
            (Source: "clients", Path: clientPath),
            (Source: "idCards", Path: idCardPath),
            (Source: "addresses", Path: addressPath),
        };
        var errors = files
            .Where(x => !File.Exists(x.Path))
            .Select(x => new ValidationError(x.Source, "file", x.Path))
            .ToList();
        if (errors.Count > 0)
            return errors;

        using var clientWb = new XLWorkbook(clientPath);
        using var idCardWb = new XLWorkbook(idCardPath);
        using var addressWb = new XLWorkbook(addressPath);
        return ValidateRequiredSheetsAndColumns(clientWb, idCardWb, addressWb);
    }

    static void ValidateSource(
        XLWorkbook workbook,
        string source,
        string sheetName,
        int headerRow,
        string[] requiredColumns,
        List<ValidationError> errors)
    {
        var worksheet = workbook.Worksheets.FirstOrDefault(
            w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (worksheet is null)
        {
            errors.Add(new ValidationError(source, "sheet", sheetName));
            return;
        }

        var headers = BuildHeaderIndex(worksheet.Row(headerRow));
        errors.AddRange(MissingColumns(headers, requiredColumns)
            .Select(column => new ValidationError(source, "column", $"{sheetName}.{column}")));
    }

    static void EnsureFileExists(string path, string source)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Profile {source} Excel file not found: {path}", path);
    }

    /// <summary>
    /// Classify MF_CLIENT rows with the same exclusion rules as the Analysis API.
    /// </summary>
    public static ClientEligibilityClassifier.Result BuildEligibility(LoadedWorkbook loaded)
    {
        var h = loaded.ClientHeaders;
        var rows = new List<ClientEligibilityClassifier.ClientIdentityRow>();

        for (int r = loaded.FirstClientDataRow; r <= loaded.LastClientRow; r++)
        {
            var row = loaded.ClientSheetWs.Row(r);
            long? clientId = GetLong(row, h, "CLIENT_ID");
            if (clientId is null || clientId <= 0) continue;

            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, h, "COMPANY"));
            loaded.IdCards.TryGetValue((company, clientId.Value), out IdCardData? idCard);
            string? idNum = idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID");
            int idType = idCard?.IdTypeId ?? 1;
            rows.Add(new ClientEligibilityClassifier.ClientIdentityRow(
                company, clientId.Value, idCard?.CardKey, idNum, idType));
        }

        return ClientEligibilityClassifier.Classify(rows);
    }

    /// <summary>
    /// Map one output row per eligible Excel row. Duplicate identities are removed by
    /// <see cref="ClientEligibilityClassifier"/>; rows are never merged.
    /// </summary>
    public static List<EligibleSourceRow> BuildEligibleRows(
        LoadedWorkbook loaded,
        List<string>? log = null,
        ClientEligibilityClassifier.Result? eligibility = null)
    {
        log ??= [];
        eligibility ??= BuildEligibility(loaded);
        var h = loaded.ClientHeaders;
        var result = new List<EligibleSourceRow>();

        for (int r = loaded.FirstClientDataRow; r <= loaded.LastClientRow; r++)
        {
            var row = loaded.ClientSheetWs.Row(r);
            long? clientId = GetLong(row, h, "CLIENT_ID");
            if (clientId is null || clientId <= 0) continue;

            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, h, "COMPANY"));
            if (!eligibility.IsEligible(company, clientId.Value))
                continue;

            loaded.IdCards.TryGetValue((company, clientId.Value), out IdCardData? idCard);
            string? idNum = idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID");
            int idType = idCard?.IdTypeId ?? 1;
            if (string.IsNullOrWhiteSpace(idNum)) continue;

            idNum = idNum.Trim();
            int? oldBrCode = GetInt(row, h, "COMPANY_BRANCH_CODE");
            int newBranchId = DefaultBranchId;
            if (oldBrCode.HasValue && loaded.BranchMap.TryGetValue((company, oldBrCode.Value), out int foundId))
                newBranchId = foundId;

            var custId = GetLong(row, h, "Combined Number");
            if (custId is null || custId <= 0) continue;

            var profile = MapProfile(row, h, idCard, 0, custId, newBranchId);
            loaded.Addresses.TryGetValue((company, clientId.Value), out AddressData? address);
            ApplyAddress(profile, address);
            var phones = address?.Phones ?? Array.Empty<string>();
            var work = MapWork(row, h);
            var partner = MapPartner(r, row, h, log);
            var bank = MapBankInfo(row, h);
            result.Add(new EligibleSourceRow(
                r, clientId.Value, company, idNum, idType,
                profile, work, partner, bank, phones));
        }

        return result;
    }

    /// <summary>
    /// Load ID cards keyed by (COMPANY, CLIENT_ID) — same join as Analysis IdCardReader.
    /// First row wins per key.
    /// </summary>
    public static Dictionary<(string Company, long ClientId), IdCardData> LoadIdCards(XLWorkbook wb)
    {
        var result = new Dictionary<(string, long), IdCardData>();
        var ws = GetRequiredSheet(wb, IdCardSheet);
        var h = BuildHeaderIndex(ws.Row(1));
        int last = ws.LastRowUsed()!.RowNumber();

        for (int r = 2; r <= last; r++)
        {
            var row = ws.Row(r);
            long? cid = GetLong(row, h, "CLIENT_ID");
            if (cid is null || cid <= 0) continue;

            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, h, "COMPANY"));
            if (!ClientEligibilityClassifier.IsKnownCompany(company)) continue;

            var key = (company, cid.Value);
            if (result.ContainsKey(key)) continue;

            int? rawType = GetInt(row, h, "ID_CARD_TYPE");
            string? rawNo = GetString(row, h, "ID_CARD_NO");

            result[key] = new IdCardData
            {
                RawIdCardType = rawType,
                RawIdCardNo = rawNo,
                IdTypeId = MapIdTypeId(rawType),
                IdNum = rawNo,
                IssueDate = GetDateTime(row, h, "ID_CARD_ISSUE_DATE"),
                IssuePlace = GetString(row, h, "ID_CARD_ISSUE_PLACE"),
                ExpiryDate = GetDateTime(row, h, "ID_CARD_EXPIRY_DATE"),
            };
        }
        return result;
    }

    public static Dictionary<(string Company, long ClientId), AddressData> LoadAddresses(
        XLWorkbook wb, Dictionary<int, int> cityIdMap, out int unmappedCityCount)
    {
        var result = new Dictionary<(string, long), AddressData>();
        var ws = GetRequiredSheet(wb, AddressSheet);
        var h = BuildHeaderIndex(ws.Row(1));
        int last = ws.LastRowUsed()!.RowNumber();
        int unmapped = 0;
        var buckets = new Dictionary<(string, long), List<(int Pref, AddressData Data)>>();

        for (int r = 2; r <= last; r++)
        {
            var row = ws.Row(r);
            long? clientId = GetLong(row, h, "CLIENT_ID");
            if (clientId is null || clientId <= 0) continue;

            string company = GetString(row, h, "COMPANY")?.ToUpperInvariant() ?? "";
            int? cityCode = GetInt(row, h, "CITY_CODE");
            int? areaCode = GetInt(row, h, "AREA_CODE");
            int? mainFlag = GetInt(row, h, "MAIN_FLAG");
            int? addrType = GetInt(row, h, "ADDRESS_TYPE");

            int? permanentStateId = null;
            if (cityCode.HasValue)
            {
                if (cityIdMap.TryGetValue(cityCode.Value, out int mapped))
                {
                    if (mapped >= 0)
                        permanentStateId = mapped;
                }
                else
                {
                    unmapped++;
                }
            }

            int? permanentCityId = areaCode is > 0 ? areaCode : null;
            string? street = GetString(row, h, "STREET");
            string? desc = GetString(row, h, "ADDRESS_DESC");
            string? permanentAddress = JoinParts(street, desc);
            if (string.IsNullOrWhiteSpace(permanentAddress))
                permanentAddress = null;

            var phones = new List<string>();
            void AddPhone(string? p)
            {
                p = T(p, 2000);
                if (!string.IsNullOrWhiteSpace(p))
                    phones.Add(p);
            }
            AddPhone(GetString(row, h, "TEL_NO"));
            AddPhone(GetString(row, h, "TEL_NO2"));
            AddPhone(GetString(row, h, "TEL_NO3"));

            var data = new AddressData(
                permanentStateId,
                permanentCityId,
                T(permanentAddress, 300),
                T(GetString(row, h, "ZIP_CODE"), 15),
                T(GetString(row, h, "PO_BOX"), 15),
                phones);

            int pref = mainFlag == 1 ? 0 : addrType == 1 ? 1 : 2;
            var key = (company, clientId.Value);
            if (!buckets.TryGetValue(key, out var list))
                buckets[key] = list = [];
            list.Add((pref, data));
        }

        foreach (var (key, list) in buckets)
            result[key] = list.OrderBy(x => x.Pref).First().Data;

        unmappedCityCount = unmapped;
        return result;
    }

    public static void ApplyAddress(ProfilesTb profile, AddressData? address)
    {
        if (address is null) return;
        profile.PermanentStateId = address.PermanentStateId;
        profile.PermanentCityId = address.PermanentCityId;
        profile.PermanentAddress = address.PermanentAddress;
        profile.MailZip = address.MailZip;
        profile.MailPobox = address.MailPobox;
    }

    public static ProfilesTb MapProfile(
        IXLRow row, Dictionary<string, int> h,
        IdCardData? idCard, int profileId, long? custId, int branchId)
    {
        string? englishName1 = GetString(row, h, "NAME1");
        string? englishName2 = GetString(row, h, "NAME2");
        string? englishName3 = GetString(row, h, "NAME3");
        string? englishName4 = GetString(row, h, "NAME4");
        string? arabicName1 = GetString(row, h, "ANAME1");
        string? arabicName2 = GetString(row, h, "ANAME2");
        string? arabicName3 = GetString(row, h, "ANAME3");
        string? arabicName4 = GetString(row, h, "ANAME4");

        string fullNameNa = JoinParts(arabicName1, arabicName2, arabicName3, arabicName4);
        string fullNameFo = JoinParts(englishName1, englishName2, englishName3, englishName4);
        if (string.IsNullOrWhiteSpace(fullNameNa)) fullNameNa = fullNameFo;

        int? clientType = GetInt(row, h, "CLIENT_TYPE");
        bool custTypeId = clientType != 3;

        int? gender = GetInt(row, h, "GENDER");
        bool? genderId = gender.HasValue ? gender.Value == 2 : null;
        int idTypeId = idCard?.IdTypeId ?? 1;

        return new ProfilesTb
        {
            ProfileId = profileId,
            CustId = custId,
            BranchId = branchId,
            FirstNameNa = T(arabicName1, 50),
            FatherNameNa = T(arabicName2, 50),
            GrandFatherNameNa = T(arabicName3, 50),
            FamilyNameNa = T(arabicName4, 50),
            FirstNameFo = T(englishName1, 50),
            FatherNameFo = T(englishName2, 50),
            GrandFatherNameFo = T(englishName3, 50),
            FamilyNameFo = T(englishName4, 50),
            ProfileNameFo = T(fullNameFo, 200),
            ProfileNameNa = T(fullNameNa, 200),
            ProfileNameFFo = T(FilterName(fullNameFo), 200),
            ProfileNameFNa = T(FilterName(fullNameNa), 200),
            ShortName = T(GetString(row, h, "NICKNAME"), 200),
            MotherName = T(GetString(row, h, "MOTHER_NAME"), 200),
            IdTypeId = idTypeId,
            IdNum = T(idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID"), 100),
            IssueDate = idCard?.IssueDate,
            IssuePlace = T(idCard?.IssuePlace, 500),
            ExpiryDate = idCard?.ExpiryDate,
            CustTypeId = custTypeId,
            CustStatusId = true,
            IsCustomer = true,
            GenderId = genderId,
            BirthDate = GetDateTime(row, h, "BIRTH_DATE"),
            BirthPlace = T(GetString(row, h, "BIRTH_PLACE"), 200),
            MaritalStatusId = MapMaritalStatus(GetByte(row, h, "MARITAL_STATUS")),
            EducationLevelId = MapEducationLevel(GetInt(row, h, "QUALIFICATION_CODE")),
            DeathDate = GetDateTime(row, h, "DEATH_DATE"),
            ChildrenCnt = MapChildrenCnt(GetByte(row, h, "NO_DEPENDENTS_CHILDREN")),
            DependentsCnt = MapDependentsCnt(GetByte(row, h, "NO_OF_DEPENDENTS")),
            FamilyMembersCnt = GetByte(row, h, "NO_FAMILY"),
            ParentsNoWorkCnt = GetByte(row, h, "NO_FAMILY_WORKERS"),
            ChronicFlag = ExcelFlag(GetInt(row, h, "CHRONIC")),
            SpecialNeedsFlag = ExcelFlag(GetInt(row, h, "SPECIAL_NEED_FLAG")),
            Weight = GetByte(row, h, "WEIGHT"),
            Hight = GetByte(row, h, "LENGTH"),
            IncomeAmount = GetDecimal(row, h, "MONTHLY_INCOME"),
            GroupId = null,
            EntrySourceId = false,
            ResidentFlag = true,
            ReviewFlag = true,
            MailAddressFlag = false,
            PermanentAddressFlag = true,
            IsCompleated = false,
            IsPortalCompleated = false,
            BlackListCount = 0,
            CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
            UpdatedOn = GetDateTime(row, h, "MODIFY_DATE"),
        };
    }

    public static ProfileWorksTb? MapWork(IXLRow row, Dictionary<string, int> h)
    {
        string? workPlace = GetString(row, h, "WORK_PLACE");
        int? profession = GetInt(row, h, "PROFESSION_CODE");
        decimal? salary = GetDecimal(row, h, "MONTHLY_INCOME");

        bool hasPlace = !string.IsNullOrWhiteSpace(workPlace);
        bool hasJob = profession.HasValue;
        bool hasPay = salary.HasValue;
        if (!hasPlace && !hasJob && !hasPay) return null;

        string place = hasPlace ? T(workPlace, 200)! : "غير محدد";
        return new ProfileWorksTb
        {
            WorkPlace = place,
            WorkNatureId = null,
            Salary = salary,
            SalaryPeriodId = 2,
            CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
        };
    }

    public static ProfileBankInformationTb? MapBankInfo(IXLRow row, Dictionary<string, int> h)
    {
        string? accountNo = GetString(row, h, "BANK_ACCOUNT_NUMBER");

        return new ProfileBankInformationTb
        {
            BankId = DefaultBankId,
            BankAccountNo = T(accountNo, 50),
            AccountCurrId = DefaultAccountCurrId,
            Iban = "",
            CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
        };
    }

    public static ProfilesPartnersTb? MapPartner(int rowNumber, IXLRow row, Dictionary<string, int> h, List<string> log)
    {
        string? pname1 = GetString(row, h, "PARTNER_NAME");
        if (string.IsNullOrWhiteSpace(pname1)) return null;

        string? idNum = GetString(row, h, "PARTNER_NATIONAL_ID");
        if (string.IsNullOrWhiteSpace(idNum))
        {
            log.Add($"[SKIP PARTNER] Row {rowNumber}: PARTNER_NAME='{pname1}' but PARTNER_NATIONAL_ID is missing — ID_NUM is required, not inserted.");
            return null;
        }

        return new ProfilesPartnersTb
        {
            FirstNameNa = T(pname1, 50),
            FatherNameNa = T(GetString(row, h, "PARTNER_NAME_2"), 50),
            GrandFatherNameNa = T(GetString(row, h, "PARTNER_NAME_5"), 50),
            FamilyNameNa = T(GetString(row, h, "PARTNER_NAME_4"), 50),
            IdTypeId = 2,
            IdNum = T(idNum, 100)!,
            EducationLevelId = MapEducationLevel(GetInt(row, h, "PARTNER_EDUCATION")),
            PhoneNumber = T(GetString(row, h, "PARTNER_MOBILE"), 30),
            CurrentExperienceNotes = GetString(row, h, "PARTNER_WORK_DESC"),
            SharesCnt = 0,
            ContributionPercent = 0,
            IsBankBorrower = false,
            IsHidden = false,
            CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
        };
    }

    public static IXLWorksheet GetRequiredSheet(XLWorkbook wb, string name) =>
        wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new InvalidOperationException($"Required worksheet not found: {name}");

    public static IEnumerable<string> MissingColumns(Dictionary<string, int> headers, params string[] required) =>
        required.Where(c => !headers.ContainsKey(c));
}
