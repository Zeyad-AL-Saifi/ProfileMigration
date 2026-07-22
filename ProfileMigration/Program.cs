using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using ProfileMigration.DAL.Models;

// ─── Configuration ────────────────────────────────────────────────────────────
var config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json")
    .Build();

string connStr = config.GetConnectionString("OracleDb")
    ?? throw new Exception("Missing 'OracleDb' connection string in appsettings.json.");

string databaseSchema = config["DatabaseSchema"] ?? "RHODES_BANKING";
SilaDbContext.DefaultSchema = databaseSchema;

string rawPath = config["ExcelFilePath"] ?? "Clients Details Combined.xlsx";
string excelPath = Path.IsPathRooted(rawPath)
    ? rawPath
    : Path.Combine(AppContext.BaseDirectory, rawPath);

string rawAreaPath = config["AreaExcelFilePath"] ?? "area 2.xlsx";
string areaExcelPath = Path.IsPathRooted(rawAreaPath)
    ? rawAreaPath
    : Path.Combine(AppContext.BaseDirectory, rawAreaPath);

Console.WriteLine($"Source : {excelPath}");
Console.WriteLine($"Areas  : {areaExcelPath}");
Console.WriteLine($"Schema : {databaseSchema}");
Console.WriteLine($"Target : {connStr}\n");

// ─── EF Core options ──────────────────────────────────────────────────────────
var dbOptions = new DbContextOptionsBuilder<SilaDbContext>()
    .UseOracle(connStr)
    .Options;

// ─── Open the workbook ONCE — shared across all phases ────────────────────────
// ClosedXML loads the full file into memory on open. Opening it once avoids
// tripling memory/parse time for the 120K+ row Excel.
Console.WriteLine("Opening workbook...");
using var wb = new XLWorkbook(excelPath);
Console.WriteLine("  Done.\n");

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 1 — Build ID card lookup from the already-open workbook
//  Key: CLIENT_ID → IdCardData
//  Memory: ~120K lightweight records (~24 MB) — acceptable for a migration tool
// ════════════════════════════════════════════════════════════════════════════════
Console.WriteLine("Phase 1 — Indexing ID card data...");
var idCardLookup = LoadIdCards(wb);
Console.WriteLine($"  {idCardLookup.Count} ID card records indexed.\n");

Console.WriteLine("Phase 1b — Indexing client addresses...");
var cityIdMap = LoadCityIdMap();
var addressLookup = LoadAddresses(wb, cityIdMap);
Console.WriteLine($"  {addressLookup.Count} address records indexed.\n");

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2a — Ensure reference branches exist (idempotent seed)
//
//  Mirrors the stakeholder BRANCHES_TB / BRANCH_LANGS_TB script:
//  insert when missing, skip when already present.
// ════════════════════════════════════════════════════════════════════════════════
Console.WriteLine("Phase 2a — Ensuring reference branches...");
await EnsureBranchesSeededAsync(dbOptions);
Console.WriteLine();

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2c — Replace area constants (CONSTANT_MAIN_ID = 10) from area 2.xlsx
// ════════════════════════════════════════════════════════════════════════════════
Console.WriteLine("Phase 2c — Replacing area constants (MAIN_ID=10) from area Excel...");
await ReplaceAreaConstantsAsync(areaExcelPath, dbOptions);
Console.WriteLine();

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2b — Branch ID remap (no inserts)
//
//  Branches already exist in the target DB. Map each source company's
//  COMPANY_BRANCH_CODE (OLD_ID) → existing BRANCHES_TB.BranchId (New ID).
//  Composite key (COMPANY, OLD_ID) because the same code means different
//  branches across ACAD vs Asala.
// ════════════════════════════════════════════════════════════════════════════════
Console.WriteLine("Phase 2b — Loading branch ID map (OLD_ID → New ID)...");
var branchLookup = LoadBranchIdMap();
Console.WriteLine($"  {branchLookup.Count} mappings loaded.\n");

Console.WriteLine("  Branch ID mapping:");
Console.WriteLine($"  {"Company",-10} {"OLD_ID",8}  {"New ID",8}");
Console.WriteLine($"  {new string('-', 30)}");
foreach (var ((company, oldCode), newId) in branchLookup.OrderBy(x => x.Key.Company).ThenBy(x => x.Key.OldCode))
    Console.WriteLine($"  {company,-10} {oldCode,8}  {newId,8}");
Console.WriteLine();

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 3 — Migrate profiles (Match MF_CLIENT → PROFILES_TB)
// ════════════════════════════════════════════════════════════════════════════════
Console.WriteLine("Phase 3 — Migrating profiles...");

const int BatchSize = 500;
int inserted = 0, skipped = 0;
// ─── Reporting counters ───────────────────────────────────────────────────────
int insertedDirect = 0;   // single-row groups, inserted as-is
int insertedMerged = 0;   // multi-row groups folded with <= MaxConflicts conflicts
int highConflict   = 0;   // groups exceeding MaxConflicts — sent to review, not inserted
int insertFailed   = 0;   // rows rejected by the DB during insert (FK/check/etc.)
int batchNo        = 0;
// ID numbers of every minor-fix merge, collected for a one-shot list at the end.
var mergedIdNums   = new List<string>();
var batchSw = new System.Diagnostics.Stopwatch();

// Seed sequences and seen-IdNums from DB so re-runs don't collide.
int nextProfileId;
int nextWorkId;
int nextPartnerId;
int nextContactId;
var seenIdNums = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
await using (var seqCtx = new SilaDbContext(dbOptions))
{
    nextProfileId = ((await seqCtx.ProfilesTbs
        .Select(p => (int?)p.ProfileId)
        .MaxAsync()) ?? -1) + 1;

    nextWorkId = ((await seqCtx.ProfileWorksTbs
        .Select(w => (int?)w.WorkId)
        .MaxAsync()) ?? -1) + 1;

    nextPartnerId = ((await seqCtx.ProfilesPartnersTbs
        .Select(p => (int?)p.PartnerId)
        .MaxAsync()) ?? -1) + 1;

    nextContactId = ((await seqCtx.ProfileContactDtsTbs
        .Select(c => (int?)c.ContactId)
        .MaxAsync()) ?? -1) + 1;

    var existingIdNums = await seqCtx.ProfilesTbs
        .Where(p => p.IdNum != null)
        .Select(p => p.IdNum)
        .ToListAsync();
    foreach (var n in existingIdNums) seenIdNums.Add(n);
}

var ws = wb.Worksheet("Match MF_CLIENT");
var h  = BuildHeaderIndex(ws.Row(4));   // headers in row 4, data from row 5
int lastRow = ws.LastRowUsed()!.RowNumber();

// ─── Merge configuration ──────────────────────────────────────────────────────
// A client registered in BOTH companies (ASALA + ACAD) appears as two rows that
// share the same (ID_NUM, ID_TYPE_ID).  Such rows are folded into one profile:
//   - one side blank      -> take the non-empty value
//   - both equal          -> keep
//   - both filled, differ -> take the NEWER row's value (+1 conflict)
// If a pair disagrees on more than MaxConflicts fields the merge is considered
// unsafe: nothing is inserted and the pair is written to the review file instead.
const int MaxConflicts = 3;
string conflictFile = Path.Combine(AppContext.BaseDirectory, "merge_conflicts.txt");
File.WriteAllText(conflictFile,
    $"Merge review — generated {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n" +
    $"Rows listed here were NOT inserted and need manual handling.\r\n\r\n");
int reviewCount = 0;

// ─── Pass A — bucket rows by merge key (ID_NUM, ID_TYPE_ID) ────────────────────
// ID_NUM is a required key: any row missing it is invalid and goes to review.
var groups = new Dictionary<(string IdNum, int IdType), List<RowData>>();

for (int r = 5; r <= lastRow; r++)
{
    var row       = ws.Row(r);
    int? clientId = GetInt(row, h, "CLIENT_ID");
    if (clientId is null || clientId <= 0) continue;

    idCardLookup.TryGetValue(clientId.Value, out IdCardData? idCard);

    string? idNum  = idCard?.IdNum ?? GetString(row, h, "NATIONAL_ID");
    int     idType = idCard?.IdTypeId ?? 1;   // already remapped in LoadIdCards

    if (string.IsNullOrWhiteSpace(idNum))
    {
        WriteConflict(conflictFile, $"Row {r} (CLIENT_ID={clientId}): missing ID_NUM — required key, not inserted.");
        reviewCount++;
        skipped++;
        continue;
    }

    // Skip rows whose IdNum is already in the DB from a previous run.
    if (seenIdNums.Contains(idNum))
    {
        skipped++;
        Console.Error.WriteLine($"  [SKIP] Row {r} (CLIENT_ID={clientId}): IdNum '{idNum}' already in DB");
        continue;
    }

    // Branch: (COMPANY, COMPANY_BRANCH_CODE/OLD_ID) → existing New ID.
    // Missing / unmapped OLD_ID defaults to BranchId 6.
    const int DefaultBranchId = 6;
    string company   = GetString(row, h, "COMPANY")?.ToUpperInvariant() ?? "";
    int? oldBrCode   = GetInt(row, h, "COMPANY_BRANCH_CODE");
    int newBranchId  = DefaultBranchId;
    if (oldBrCode.HasValue)
    {
        if (branchLookup.TryGetValue((company, oldBrCode.Value), out int foundId))
            newBranchId = foundId;
        else
            Console.Error.WriteLine(
                $"  [WARN] Row {r} (CLIENT_ID={clientId}): unmapped branch OLD_ID={oldBrCode} COMPANY={company} → default BranchId={DefaultBranchId}");
    }

    var profile = MapProfile(row, h, idCard, 0, newBranchId);   // ProfileId assigned later
    addressLookup.TryGetValue((company, clientId.Value), out AddressData? address);
    ApplyAddress(profile, address);
    var phones  = address?.Phones ?? Array.Empty<string>();
    var work    = MapWork(row, h);                              // null when no work/salary/profession data
    var partner = MapPartner(r, row, h);                        // null when no PARTNER_NAME (or missing PARTNER_NATIONAL_ID)
    var bank    = MapBankInfo(row, h);                          // null when no BANK_ACCOUNT_NUMBER
    DateTime sortDate = GetDateTime(row, h, "MODIFY_DATE")
                        ?? GetDateTime(row, h, "CREATE_DATE")
                        ?? DateTime.MinValue;

    var key = (idNum.Trim(), idType);
    if (!groups.TryGetValue(key, out var list))
        groups[key] = list = new List<RowData>();
    list.Add(new RowData(r, clientId.Value, company, sortDate, profile, work, partner, bank, phones));
}

// ─── Pass B — merge each group, then batch-insert ─────────────────────────────
// Each queued profile carries a flag: was it inserted directly (single row) or
// after a merge (multi-row group folded within the conflict budget)?  This lets
// the per-batch report split "direct" vs "minor fix".
// Each queued entry also carries the work row (if the winning source row had
// WORK_PLACE) so it can be inserted right after its profile commits.
var batch = new List<(ProfilesTb Profile, ProfileWorksTb? Work, ProfilesPartnersTb? Partner, ProfileBankInformationTb? Bank, List<ProfileContactDtsTb> Contacts, bool Merged, int Conflicts)>(BatchSize);
int worksQueued = 0, worksInserted = 0, worksFailed = 0;
int partnersQueued = 0, partnersInserted = 0, partnersFailed = 0;
int banksQueued = 0, banksInserted = 0, banksFailed = 0;
int contactsQueued = 0, contactsInserted = 0, contactsFailed = 0;

foreach (var ((idNum, idType), rows) in groups)
{
    ProfilesTb merged;
    ProfileWorksTb? work;          // work from the winning (newest) row
    ProfilesPartnersTb? partner;   // partner from the winning (newest) row
    ProfileBankInformationTb? bank;
    IReadOnlyList<string> phones;
    bool wasMerged;
    int conflictCount = 0;

    if (rows.Count == 1)
    {
        merged  = rows[0].Profile;
        work    = rows[0].Work;
        partner = rows[0].Partner;
        bank    = rows[0].Bank;
        phones  = rows[0].Phones;
        wasMerged = false;
    }
    else
    {
        // Fold newest-first so the newer row wins on differing values.
        var ordered = rows.OrderByDescending(x => x.SortDate).ToList();
        var conflicts = new List<string>();
        merged = ordered[0].Profile;
        // Work/Partner/Bank/Phones: take from the newest row that has the data.
        work    = ordered.Select(x => x.Work).FirstOrDefault(w => w is not null);
        partner = ordered.Select(x => x.Partner).FirstOrDefault(p => p is not null);
        bank    = ordered.Select(x => x.Bank).FirstOrDefault(b => b is not null);
        phones  = ordered.Select(x => x.Phones).FirstOrDefault(p => p.Count > 0) ?? Array.Empty<string>();
        for (int i = 1; i < ordered.Count; i++)
            MergeInto(merged, ordered[i].Profile, conflicts);

        conflictCount = conflicts.Count;

        if (conflictCount > MaxConflicts)
        {
            var sources = string.Join(", ", rows.Select(x => $"{x.Company}#row{x.RowNumber}(CLIENT_ID={x.ClientId})"));
            WriteConflict(conflictFile,
                $"ID_NUM={idNum} ID_TYPE={idType}: {conflictCount} conflicting fields (> {MaxConflicts}) — NOT inserted.\r\n" +
                $"    Sources: {sources}\r\n" +
                "    " + string.Join("\r\n    ", conflicts));
            reviewCount++;
            highConflict++;
            skipped += rows.Count;
            continue;
        }

        // Minor-fix merge (1..MaxConflicts conflicts): log the ID_NUM and source
        // rows so it can be validated against the sheet.  Merges are rare, so the
        // extra console lines stay manageable.
        var mergeSources = string.Join(", ", rows.Select(x => $"{x.Company}#row{x.RowNumber}(CLIENT_ID={x.ClientId})"));
        Console.WriteLine(
            $"  [MERGE] ID_NUM={idNum} ID_TYPE={idType} rows={rows.Count} conflicts={conflictCount}  sources: {mergeSources}");
        mergedIdNums.Add(idNum);

        wasMerged = true;
    }

    merged.ProfileId = nextProfileId;
    merged.CustId    = nextProfileId;

    if (work is not null)
    {
        work.ProfileId = nextProfileId;
        work.WorkId    = nextWorkId++;
        worksQueued++;
    }

    if (partner is not null)
    {
        partner.ProfileId  = nextProfileId;
        partner.PartnerId  = nextPartnerId++;
        partnersQueued++;
    }

    if (bank is not null)
    {
        bank.ProfileId = nextProfileId;
        // IBAN is NOT NULL + unique index; no source → unique placeholder per profile
        bank.Iban = $"MIG-{nextProfileId}";
        banksQueued++;
    }

    var contactRows = new List<ProfileContactDtsTb>(phones.Count);
    foreach (var phone in phones)
    {
        contactRows.Add(new ProfileContactDtsTb
        {
            ContactId     = nextContactId++,
            ProfileId     = nextProfileId,
            ContactTypeId = 1,          // phone
            ContactInfo   = phone,
            CreatedOn     = DateTime.Now,
        });
    }
    contactsQueued += contactRows.Count;

    nextProfileId++;
    seenIdNums.Add(idNum);
    batch.Add((merged, work, partner, bank, contactRows, wasMerged, conflictCount));

    if (batch.Count >= BatchSize)
        await FlushBatchAsync();
}

if (batch.Count > 0)
    await FlushBatchAsync();

Console.WriteLine($"\nDone.");
Console.WriteLine($"  Inserted total           : {inserted}");
Console.WriteLine($"    - directly (1 row)     : {insertedDirect}");
Console.WriteLine($"    - with a minor fix     : {insertedMerged}   (merged within {MaxConflicts}-conflict budget)");
if (mergedIdNums.Count > 0)
    Console.WriteLine($"      merged ID_NUMs       : {string.Join(", ", mergedIdNums)}");
Console.WriteLine($"  High-conflict (> {MaxConflicts})     : {highConflict}   (sent to review, not inserted)");
Console.WriteLine($"  Rejected by DB           : {insertFailed}");
Console.WriteLine($"  Skipped (other)          : {skipped}");
Console.WriteLine($"  Sent to review (total)   : {reviewCount}");
Console.WriteLine($"  Work records inserted    : {worksInserted} of {worksQueued} queued" +
                  (worksFailed > 0 ? $"  ({worksFailed} rejected)" : ""));
Console.WriteLine($"  Partner records inserted : {partnersInserted} of {partnersQueued} queued" +
                  (partnersFailed > 0 ? $"  ({partnersFailed} rejected)" : ""));
Console.WriteLine($"  Contact records inserted : {contactsInserted} of {contactsQueued} queued" +
                  (contactsFailed > 0 ? $"  ({contactsFailed} rejected)" : ""));
Console.WriteLine($"  Bank records inserted    : {banksInserted} of {banksQueued} queued" +
                  (banksFailed > 0 ? $"  ({banksFailed} rejected)" : ""));
if (reviewCount > 0)
    Console.WriteLine($"  Review file: {conflictFile}");

// Insert the queued batch, report the direct/merged/failed split for it, and
// update the running totals.  Declared as a local function so it shares the
// counters above without passing them around.
async Task FlushBatchAsync()
{
    batchNo++;
    int directInBatch = batch.Count(b => !b.Merged);
    int mergedInBatch = batch.Count(b =>  b.Merged);

    batchSw.Restart();
    int ok = await InsertProfileBatchAsync(dbOptions, batch.Select(b => b.Profile).ToList());
    batchSw.Stop();

    // Insert work rows for this batch.  Their ProfileId FK requires the owning
    // profile to have committed; the row-by-row fallback inside the insert helper
    // logs (and skips) any work whose profile was rejected, so no orphan is forced.
    var works = batch.Where(b => b.Work is not null).Select(b => b.Work!).ToList();
    int worksOk = works.Count > 0 ? await InsertWorkBatchAsync(dbOptions, works) : 0;
    worksInserted += worksOk;
    worksFailed   += works.Count - worksOk;

    // Same pattern as works: PARTNER_ID FK requires the owning profile to have
    // committed first; row-by-row fallback inside the helper logs/skips any
    // partner whose profile was rejected.
    var partners = batch.Where(b => b.Partner is not null).Select(b => b.Partner!).ToList();
    int partnersOk = partners.Count > 0 ? await InsertPartnerBatchAsync(dbOptions, partners) : 0;
    partnersInserted += partnersOk;
    partnersFailed   += partners.Count - partnersOk;

    var contacts = batch.SelectMany(b => b.Contacts).ToList();
    int contactsOk = contacts.Count > 0 ? await InsertContactBatchAsync(dbOptions, contacts) : 0;
    contactsInserted += contactsOk;
    contactsFailed   += contacts.Count - contactsOk;

    var banks = batch.Where(b => b.Bank is not null).Select(b => b.Bank!).ToList();
    int banksOk = banks.Count > 0 ? await InsertBankBatchAsync(dbOptions, banks) : 0;
    banksInserted += banksOk;
    banksFailed   += banks.Count - banksOk;

    int failedInBatch = batch.Count - ok;
    inserted     += ok;
    insertFailed += failedInBatch;

    // When every queued row committed (the normal case) the direct/merged split
    // is exact.  If the DB rejected some rows we can't know which category they
    // fell in without per-row tracking, so we credit the categories by what was
    // queued and rely on `insertFailed` to flag that the split is approximate.
    insertedDirect += directInBatch;
    insertedMerged += mergedInBatch;

    string note = failedInBatch > 0 ? "  (split approximate — see [SKIP] rows)" : "";
    Console.WriteLine(
        $"  [batch {batchNo}] queued={batch.Count} " +
        $"direct={directInBatch} merged={mergedInBatch} " +
        $"inserted={ok} failed={failedInBatch} " +
        $"works={worksOk}/{works.Count} " +
        $"partners={partnersOk}/{partners.Count} " +
        $"contacts={contactsOk}/{contacts.Count} " +
        $"banks={banksOk}/{banks.Count} " +
        $"(took {batchSw.Elapsed.TotalSeconds:F2}s)  running total={inserted}{note}");

    batch.Clear();
}

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2a impl — Idempotent branch seed (matches stakeholder SQL script)
//
//  Country/State/City IDs: 1..9 for Jenin..Gaza (CountryId = 1).
//  Skip when PK already exists; insert only missing BRANCHES / BRANCH_LANGS rows.
// ════════════════════════════════════════════════════════════════════════════════

static async Task EnsureBranchesSeededAsync(DbContextOptions<SilaDbContext> opts)
{
    // country=1; state/city/governorate = branch id (same as stakeholder PL/SQL vars)
    const int CountryId = 1;
    var today = DateTime.Today;
    var now   = DateTime.Now;

    // (BranchId, EnglishName, IsCentral)
    var branches = new (int Id, string NameEn, bool IsCentral)[]
    {
        (1, "Jenin",     false),
        (2, "Tulkarm",   false),
        (3, "Tubas",     false),
        (4, "Nablus",    false),
        (5, "Ramallah",  true),    // IS_CENTRAL_FLAG = 1
        (6, "Jericho",   false),
        (7, "Bethlehem", false),
        (8, "Hebron",    false),
        (9, "Gaza",      false),
    };

    // LANG_ID values match the script exactly: lowercase 'en' / 'ar'
    var langs = new (int BranchId, string LangId, string Name)[]
    {
        (1, "en", "Jenin"),     (1, "ar", "جنين"),
        (2, "en", "Tulkarm"),   (2, "ar", "طولكرم"),
        (3, "en", "Tubas"),     (3, "ar", "طوباس"),
        (4, "en", "Nablus"),    (4, "ar", "نابلس"),
        (5, "en", "Ramallah"),  (5, "ar", "رام الله"),
        (6, "en", "Jericho"),   (6, "ar", "أريحا"),
        (7, "en", "Bethlehem"), (7, "ar", "بيت لحم"),
        (8, "en", "Hebron"),    (8, "ar", "الخليل"),
        (9, "en", "Gaza"),      (9, "ar", "غزة"),
    };

    await using var ctx = new SilaDbContext(opts);

    var existingIds = (await ctx.BranchesTbs
        .Select(b => b.BranchId)
        .ToListAsync())
        .ToHashSet();

    var existingLangKeys = (await ctx.BranchLangsTbs
        .Select(l => new { l.BranchId, l.LangId })
        .ToListAsync())
        .Select(x => (x.BranchId, LangId: x.LangId.ToLowerInvariant()))
        .ToHashSet();

    int branchesInserted = 0, branchesSkipped = 0;
    int langsInserted = 0, langsSkipped = 0;

    foreach (var (id, nameEn, isCentral) in branches)
    {
        if (existingIds.Contains(id))
        {
            branchesSkipped++;
            Console.WriteLine($"  [SKIP] BRANCHES_TB BranchId={id} ({nameEn}) already exists");
            continue;
        }

        ctx.BranchesTbs.Add(new BranchesTb
        {
            BranchId            = id,
            BranchName          = nameEn,
            CountryId           = CountryId,
            StateId             = id,
            CityId              = id,
            Address             = null,
            PmaBranchCode       = null,
            IsCentralFlag       = isCentral,
            CurrDate            = today,
            PrevDate            = null,
            DailyCloseFlag      = false,
            DailyCloseStartTime = null,
            CloseFlag           = false,
            CreatedOn           = now,
            CreatedBy           = null,
            UpdatedOn           = null,
            UpdatedBy           = null,
            IsHidden            = false,
            GovernorateId       = id,
        });
        branchesInserted++;
        Console.WriteLine($"  [ADD]  BRANCHES_TB BranchId={id} ({nameEn})");
    }

    foreach (var (branchId, langId, name) in langs)
    {
        if (existingLangKeys.Contains((branchId, langId.ToLowerInvariant())))
        {
            langsSkipped++;
            Console.WriteLine($"  [SKIP] BRANCH_LANGS_TB ({branchId},{langId}) already exists");
            continue;
        }

        ctx.BranchLangsTbs.Add(new BranchLangsTb
        {
            BranchId   = branchId,
            LangId     = langId,
            BranchName = name,
            Address    = null,
            CreatedOn  = now,
        });
        langsInserted++;
        Console.WriteLine($"  [ADD]  BRANCH_LANGS_TB ({branchId},{langId}) '{name}'");
    }

    if (branchesInserted + langsInserted > 0)
        await ctx.SaveChangesAsync();

    Console.WriteLine(
        $"  Branches: {branchesInserted} inserted, {branchesSkipped} skipped | " +
        $"Langs: {langsInserted} inserted, {langsSkipped} skipped");

    // Jordanian passport constant (C_CONSTANTS_TB) — insert if missing
    await EnsureJordanianPassportConstantAsync(ctx);
}

/// <summary>
/// Ensures CONSTANT_MAIN_ID=1 / CONSTANT_ID=5 (جواز سفر اردني) exists.
/// Skips the ROWID-based UPDATE from the stakeholder script (not portable).
/// </summary>
static async Task EnsureJordanianPassportConstantAsync(SilaDbContext ctx)
{
    var constantsTable = $"{SilaDbContext.DefaultSchema}.C_CONSTANTS_TB";

    var checkSql = $"""
        SELECT COUNT(1)
        FROM {constantsTable}
        WHERE CONSTANT_MAIN_ID = 1
          AND CONSTANT_ID = 5
        """;

    var conn = ctx.Database.GetDbConnection();
    if (conn.State != System.Data.ConnectionState.Open)
        await conn.OpenAsync();

    await using (var cmd = conn.CreateCommand())
    {
        cmd.CommandText = checkSql;
        var count = Convert.ToInt32(await cmd.ExecuteScalarAsync());
        if (count > 0)
        {
            Console.WriteLine("  [SKIP] C_CONSTANTS_TB (1,5) Jordanian passport already exists");
            return;
        }
    }

    var insertSql = $"""
        INSERT INTO {constantsTable}
            (CONSTANT_MAIN_ID, CONSTANT_ID, CONSTANT_DESC, PARENT_CONSTANT_ID,
             PMA_CODE, IS_UPDATABLE, IS_HIDDEN, CREATED_ON)
        VALUES
            (1, 5, N'جواز سفر اردني', 2, N'N', 0, 0, DATE '2023-02-08')
        """;

    await ctx.Database.ExecuteSqlRawAsync(insertSql);
    Console.WriteLine("  [ADD]  C_CONSTANTS_TB (1,5) جواز سفر اردني");
}

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2c impl — Replace areas into C_CONSTANTS_TB / C_CONSTANT_LANGS_TB
//
//  Mapping (CONSTANT_MAIN_ID = 10):
//    CONSTANT_ID        ← OLD_AREA_CODE
//      (AREA_CODE alone is NOT unique across cities — 697 rows / 158 codes —
//       so OLD_AREA_CODE is used as the unique CONSTANT_ID.)
//    PARENT_CONSTANT_ID ← NEW_CITY_CODE
//    CONSTANT_DESC      ← AREA_DESC_A
//    langs ar/en        ← AREA_DESC_A / AREA_DESC_E
//    other columns      ← null / defaults (IS_UPDATABLE=0, IS_HIDDEN=0)
//
//  Existing rows with MAIN_ID=10 are deleted then re-inserted (full replace).
// ════════════════════════════════════════════════════════════════════════════════

static async Task ReplaceAreaConstantsAsync(string areaExcelPath, DbContextOptions<SilaDbContext> opts)
{
    const int AreaMainId = 10;

    if (!File.Exists(areaExcelPath))
        throw new FileNotFoundException($"Area Excel not found: {areaExcelPath}");

    using var wb = new XLWorkbook(areaExcelPath);
    var ws = wb.Worksheets.First();
    var headerRow = ws.Row(1);
    var h = BuildHeaderIndex(headerRow);
    int lastRow = ws.LastRowUsed()!.RowNumber();

    RequireCol(h, "AREA_CODE");
    RequireCol(h, "NEW_CITY_CODE");
    RequireCol(h, "AREA_DESC_A");
    RequireCol(h, "AREA_DESC_E");
    RequireCol(h, "OLD_AREA_CODE");

    var constants = new List<CConstantsTb>();
    var langs     = new List<CConstantLangsTb>();
    var seenIds   = new HashSet<int>();
    int skipped   = 0;
    var now       = DateTime.Now;

    for (int r = 2; r <= lastRow; r++)
    {
        var row = ws.Row(r);

        int? areaCode  = GetInt(row, h, "AREA_CODE");
        int? newCityId = GetInt(row, h, "NEW_CITY_CODE");
        int? oldAreaId = GetInt(row, h, "OLD_AREA_CODE");
        string? descAr = GetString(row, h, "AREA_DESC_A");
        string? descEn = GetString(row, h, "AREA_DESC_E");

        // Skip the blank/placeholder row (AREA=0, NEW_CITY=-1, DESC='-')
        if (oldAreaId is null or <= 0 || newCityId is null or < 0 ||
            string.IsNullOrWhiteSpace(descAr) || descAr == "-")
        {
            skipped++;
            continue;
        }

        if (!seenIds.Add(oldAreaId.Value))
        {
            Console.Error.WriteLine(
                $"  [WARN] Row {r}: duplicate OLD_AREA_CODE={oldAreaId} (AREA_CODE={areaCode}) — skipped");
            skipped++;
            continue;
        }

        string ar = T(descAr, 200)!;
        string en = T(descEn ?? descAr, 200)!;

        constants.Add(new CConstantsTb
        {
            ConstantMainId   = AreaMainId,
            ConstantId       = oldAreaId.Value,
            ConstantDesc     = ar,
            ParentConstantId = newCityId.Value,
            PmaCode          = null,
            ConstantCondition = null,
            ConstantCode1    = null,
            ConstantCode2    = null,
            ConstantCode3    = null,
            IsUpdatable      = false,
            IsHidden         = false,
            CreatedOn        = now,
        });

        langs.Add(new CConstantLangsTb
        {
            ConstantMainId = AreaMainId,
            ConstantId     = oldAreaId.Value,
            LangId         = "ar",
            ConstantDesc   = ar,
            CreatedOn      = now,
        });
        langs.Add(new CConstantLangsTb
        {
            ConstantMainId = AreaMainId,
            ConstantId     = oldAreaId.Value,
            LangId         = "en",
            ConstantDesc   = en,
            CreatedOn      = now,
        });
    }

    await using var ctx = new SilaDbContext(opts);

    int deletedLangs = await ctx.CConstantLangsTbs
        .Where(x => x.ConstantMainId == AreaMainId)
        .ExecuteDeleteAsync();
    int deletedConsts = await ctx.CConstantsTbs
        .Where(x => x.ConstantMainId == AreaMainId)
        .ExecuteDeleteAsync();

    Console.WriteLine($"  Deleted existing MAIN_ID={AreaMainId}: {deletedConsts} constants, {deletedLangs} langs");

    ctx.CConstantsTbs.AddRange(constants);
    ctx.CConstantLangsTbs.AddRange(langs);
    await ctx.SaveChangesAsync();

    Console.WriteLine(
        $"  Inserted {constants.Count} areas + {langs.Count} lang rows " +
        $"(skipped {skipped}). CONSTANT_ID=OLD_AREA_CODE, PARENT=NEW_CITY_CODE.");
}

static void RequireCol(Dictionary<string, int> h, string col)
{
    if (!h.ContainsKey(col))
        throw new Exception($"Area Excel missing required column '{col}'.");
}

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 2b impl — Stakeholder branch remap
//
//  OLD_ID = MF_CLIENT.COMPANY_BRANCH_CODE (per company).
//  New ID = existing BRANCHES_TB.BranchId in the target system.
// ════════════════════════════════════════════════════════════════════════════════

static Dictionary<(string Company, int OldCode), int> LoadBranchIdMap()
{
    // (COMPANY, OLD_ID) → New ID
    return new Dictionary<(string, int), int>
    {
        // ── ACAD ──────────────────────────────────────────────────────────────
        { ("ACAD",   0), 5 },   // (empty / default)
        { ("ACAD",  10), 4 },   // نابلس
        { ("ACAD",  20), 5 },   // رام الله
        { ("ACAD",  30), 6 },   // أريحا
        { ("ACAD",  40), 8 },   // الخليل
        { ("ACAD",  50), 9 },   // فرع غزة
        { ("ACAD",  60), 7 },   // بيت لحم
        { ("ACAD",  70), 5 },   // هوب
        { ("ACAD",  80), 1 },   // جنين
        { ("ACAD",  90), 2 },   // طولكرم
        { ("ACAD", 100), 5 },   // الإدارة العامة
        { ("ACAD", 110), 3 },   // طوباس
        { ("ACAD", 120), 2 },   // قلقيلية
        { ("ACAD", 130), 4 },   // مكتب سلفيت
        { ("ACAD", 200), 1 },   // فرع الشمال
        { ("ACAD", 400), 8 },   // فرع الجنوب
        { ("ACAD", 999), 5 },   // قروض الموظفين

        // ── Asala ─────────────────────────────────────────────────────────────
        { ("ASALA",   0), 5 },  // Default / Transfer
        { ("ASALA",  10), 5 },  // رام الله
        { ("ASALA",  20), 1 },  // جنين
        { ("ASALA",  30), 9 },  // غزة
        { ("ASALA",  40), 5 },  // مكتب الإدارة
        { ("ASALA",  50), 7 },  // بيت لحم
        { ("ASALA",  60), 4 },  // نابلس
        { ("ASALA",  70), 8 },  // الخليل
        { ("ASALA",  80), 9 },  // غزة - الرمال
        { ("ASALA",  90), 9 },  // غزة - دير البلح
        { ("ASALA", 100), 9 },  // غزة - بيت حانون
        { ("ASALA", 110), 9 },  // غزة - النصيرات
        { ("ASALA", 120), 9 },  // غزة - جباليا
        { ("ASALA", 130), 2 },  // طولكرم
        // ASALA 990 (NA) — not in stakeholder map → falls back to default BranchId 6
    };
}

// ════════════════════════════════════════════════════════════════════════════════
//  Phase 1 impl — ID card loader
// ════════════════════════════════════════════════════════════════════════════════

static Dictionary<int, IdCardData> LoadIdCards(XLWorkbook wb)
{
    var result = new Dictionary<int, IdCardData>();
    var ws     = wb.Worksheet("Match MF_CLIENT_ID_CARD");
    var h      = BuildHeaderIndex(ws.Row(1));   // headers in row 1
    int last   = ws.LastRowUsed()!.RowNumber();

    for (int r = 2; r <= last; r++)
    {
        var row = ws.Row(r);
        int? cid = GetInt(row, h, "CLIENT_ID");
        if (cid is null || result.ContainsKey(cid.Value)) continue;

        result[cid.Value] = new IdCardData
        {
            IdTypeId   = MapIdTypeId(GetInt(row, h, "ID_CARD_TYPE")),
            IdNum      = GetString(row, h, "ID_CARD_NO"),
            IssueDate  = GetDateTime(row, h, "ID_CARD_ISSUE_DATE"),
            IssuePlace = GetString(row, h, "ID_CARD_ISSUE_PLACE"),
            ExpiryDate = GetDateTime(row, h, "ID_CARD_EXPIRY_DATE"),
        };
    }
    return result;
}

// ════════════════════════════════════════════════════════════════════════════════
//  Address mapping (Match MF_CLIENT_ADDRESSES → PROFILES_TB + PROFILE_CONTACT_DTS_TB)
//
//  CITY_CODE → PERMANENT_STATE_ID via stakeholder NEW_IDs map (-1 / unmapped → null)
//  AREA_CODE → PERMANENT_CITY_ID  (literal; 0 → null)
//  ZIP_CODE  → MAIL_ZIP
//  PO_BOX    → MAIL_POBOX
//  STREET + ADDRESS_DESC → PERMANENT_ADDRESS
//  TEL_NO / TEL_NO2 / TEL_NO3 → PROFILE_CONTACT_DTS_TB (CONTACT_TYPE_ID = 1)
// ════════════════════════════════════════════════════════════════════════════════

static Dictionary<int, int> LoadCityIdMap() => new()
{
    // OLD_CITY_CODE → NEW_ID  (stakeholder CityB remap)
    { 0,  -1 },  // placeholder → null
    { 1,   7 },  // سلفيت
    { 2,   3 },  // القدس
    { 3,   6 },  // الادارة العامة
    { 4,  16 },  // طوباس
    { 5,   6 },  // رام الله
    { 6,   1 },  // أريحا
    { 7,  10 },  // نابلس
    { 8,   9 },  // قلقيلية
    { 9,  11 },  // غزة
    { 10,  4 },  // بيت لحم
    { 11,  5 },  // جنين
    { 12,  8 },  // طولكرم
    { 13,  2 },  // الخليل
};

static Dictionary<(string Company, int ClientId), AddressData> LoadAddresses(
    XLWorkbook wb, Dictionary<int, int> cityIdMap)
{
    var result = new Dictionary<(string, int), AddressData>();
    var ws     = wb.Worksheet("Match MF_CLIENT_ADDRESSES");
    var h      = BuildHeaderIndex(ws.Row(1));
    int last   = ws.LastRowUsed()!.RowNumber();

    // Collect candidates per client; prefer MAIN_FLAG=1, then ADDRESS_TYPE=1.
    var buckets = new Dictionary<(string, int), List<(int Pref, AddressData Data)>>();

    for (int r = 2; r <= last; r++)
    {
        var row = ws.Row(r);
        int? clientId = GetInt(row, h, "CLIENT_ID");
        if (clientId is null || clientId <= 0) continue;

        string company = GetString(row, h, "COMPANY")?.ToUpperInvariant() ?? "";
        int? cityCode  = GetInt(row, h, "CITY_CODE");
        int? areaCode  = GetInt(row, h, "AREA_CODE");
        int? mainFlag  = GetInt(row, h, "MAIN_FLAG");
        int? addrType  = GetInt(row, h, "ADDRESS_TYPE");

        int? permanentStateId = null;
        if (cityCode.HasValue && cityIdMap.TryGetValue(cityCode.Value, out int mapped))
        {
            if (mapped >= 0)
                permanentStateId = mapped;
            // mapped == -1 → leave null
        }
        // Asala city codes (100+) have no stakeholder map yet → leave null

        int? permanentCityId = areaCode is > 0 ? areaCode : null;

        string? street = GetString(row, h, "STREET");
        string? desc   = GetString(row, h, "ADDRESS_DESC");
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

        // Preference: MAIN_FLAG=1 best, then ADDRESS_TYPE=1, else anything
        int pref = mainFlag == 1 ? 0 : addrType == 1 ? 1 : 2;
        var key = (company, clientId.Value);
        if (!buckets.TryGetValue(key, out var list))
            buckets[key] = list = new List<(int, AddressData)>();
        list.Add((pref, data));
    }

    foreach (var (key, list) in buckets)
        result[key] = list.OrderBy(x => x.Pref).First().Data;

    return result;
}

static void ApplyAddress(ProfilesTb profile, AddressData? address)
{
    if (address is null) return;

    profile.PermanentStateId = address.PermanentStateId;
    profile.PermanentCityId  = address.PermanentCityId;
    profile.PermanentAddress = address.PermanentAddress;
    profile.MailZip          = address.MailZip;
    profile.MailPobox        = address.MailPobox;
}

// ════════════════════════════════════════════════════════════════════════════════
//  Profile mapping
// ════════════════════════════════════════════════════════════════════════════════

static ProfilesTb MapProfile(
    IXLRow row, Dictionary<string, int> h,
    IdCardData? idCard, int profileId, int branchId)
{
    string? name1  = GetString(row, h, "NAME1");
    string? name2  = GetString(row, h, "NAME2");
    string? name3  = GetString(row, h, "NAME3");
    string? name4  = GetString(row, h, "NAME4");
    string? aname1 = GetString(row, h, "ANAME1");
    string? aname2 = GetString(row, h, "ANAME2");
    string? aname3 = GetString(row, h, "ANAME3");
    string? aname4 = GetString(row, h, "ANAME4");

    // NAME1-4 are the local (NA) names; ANAME1-4 are the foreign (FO) names.
    string fullNameNa = JoinParts(name1, name2, name3, name4);
    string fullNameFo = JoinParts(aname1, aname2, aname3, aname4);
    if (string.IsNullOrWhiteSpace(fullNameNa)) fullNameNa = fullNameFo;

    // Excel CLIENT_TYPE: 1=Individual, 2=Group, 3=Company/Institution
    // DB   CUST_TYPE_ID: 1=Firm, 2=Individual  (stored via converter: false->1, true->2)
    // Map by meaning: Institution/Company → Firm(1); Individual/Group → Individual(2).
    int? clientType = GetInt(row, h, "CLIENT_TYPE");
    bool custTypeId = clientType != 3;   // true = Individual(2), false = Firm(1)

    // Gender: old/new both 1=Male, 2=Female (converter: false->1, true->2).
    int? gender = GetInt(row, h, "GENDER");
    bool? genderId = gender.HasValue ? gender.Value == 2 : null;

    // ID type already remapped in LoadIdCards; fallback NATIONAL_ID → هوية (1).
    int idTypeId = idCard?.IdTypeId ?? 1;

    return new ProfilesTb
    {
        ProfileId = profileId,
        CustId    = profileId,
        BranchId  = branchId,           // resolved via (COMPANY, OLD_ID) → New ID map (default 6)

        FirstNameNa       = T(name1,  50),
        FatherNameNa      = T(name2,  50),
        GrandFatherNameNa = T(name3,  50),
        FamilyNameNa      = T(name4,  50),
        FirstNameFo       = T(aname1, 50),
        FatherNameFo      = T(aname2, 50),
        GrandFatherNameFo = T(aname3, 50),
        FamilyNameFo      = T(aname4, 50),
        ProfileNameFo     = T(fullNameFo, 200),
        ProfileNameNa     = T(fullNameNa, 200),    // REQUIRED — never null
        ProfileNameFFo    = T(FilterName(fullNameFo), 200),
        ProfileNameFNa    = T(FilterName(fullNameNa), 200),
        ShortName         = T(GetString(row, h, "NICKNAME"),    200),
        MotherName        = T(GetString(row, h, "MOTHER_NAME"), 200),

        IdTypeId   = idTypeId,
        IdNum      = T(idCard?.IdNum    ?? GetString(row, h, "NATIONAL_ID"), 100),
        IssueDate  = idCard?.IssueDate,
        IssuePlace = T(idCard?.IssuePlace, 500),
        ExpiryDate = idCard?.ExpiryDate,

        CustTypeId   = custTypeId,
        CustStatusId = true,     // default bool converter: true -> 1 (active)
        IsCustomer   = true,     // true -> 1 (is a customer)

        GenderId         = genderId,
        BirthDate        = GetDateTime(row, h, "BIRTH_DATE"),
        BirthPlace       = T(GetString(row, h, "BIRTH_PLACE"), 200),
        MaritalStatusId  = MapMaritalStatus(GetByte(row, h, "MARITAL_STATUS")),
        EducationLevelId = MapEducationLevel(GetInt(row, h, "QUALIFICATION_CODE")),
        DeathDate        = GetDateTime(row, h, "DEATH_DATE"),

        ChildrenCnt      = MapChildrenCnt(GetByte(row, h, "NO_DEPENDENTS_CHILDREN")),
        DependentsCnt    = MapDependentsCnt(GetByte(row, h, "NO_OF_DEPENDENTS")),
        FamilyMembersCnt = GetByte(row, h, "NO_FAMILY"),
        ParentsNoWorkCnt = GetByte(row, h, "NO_FAMILY_WORKERS"),

        ChronicFlag      = ExcelFlag(GetInt(row, h, "CHRONIC")),
        SpecialNeedsFlag = ExcelFlag(GetInt(row, h, "SPECIAL_NEED_FLAG")),

        Weight = GetByte(row, h, "WEIGHT"),
        Hight  = GetByte(row, h, "LENGTH"),

        IncomeAmount = GetDecimal(row, h, "MONTHLY_INCOME"),
        GroupId      = null,

        EntrySourceId        = false,   // code converter: false -> 1 (Company)
        ResidentFlag         = true,    // 1 = resident
        ReviewFlag           = true,    // 1 = reviewed by the company (per stakeholder default)
        MailAddressFlag      = false,   // 0 = mail address NOT same as current address (per stakeholder default)
        PermanentAddressFlag = true,    // 1
        IsCompleated         = false,   // 0
        IsPortalCompleated   = false,   // 0
        BlackListCount       = 0,

        CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
        UpdatedOn = GetDateTime(row, h, "MODIFY_DATE"),
    };
}

// ════════════════════════════════════════════════════════════════════════════════
//  Work mapping — PROFILE_WORKS_TB  (FK: ProfileId → PROFILES_TB)
//
//  PROFESSION_CODE → WORK_NATURE_ID : left null (no agreed target codes yet;
//                                     stakeholder sheet default = null).
//  MONTHLY_INCOME  → SALARY + SALARY_PERIOD_ID = 2 (شهري).
//  WORK_PLACE      → WORK_PLACE (NOT NULL). Source often blank; use "غير محدد"
//                    when we still insert because salary/profession exist.
//
//  Insert when any of: WORK_PLACE, PROFESSION_CODE, or MONTHLY_INCOME is present.
//  WorkId / ProfileId assigned later once the owning profile's ProfileId is known.
// ════════════════════════════════════════════════════════════════════════════════
static ProfileWorksTb? MapWork(IXLRow row, Dictionary<string, int> h)
{
    string? workPlace = GetString(row, h, "WORK_PLACE");
    int? profession   = GetInt(row, h, "PROFESSION_CODE");
    decimal? salary   = GetDecimal(row, h, "MONTHLY_INCOME");

    bool hasPlace = !string.IsNullOrWhiteSpace(workPlace);
    bool hasJob   = profession.HasValue;
    bool hasPay   = salary.HasValue;

    if (!hasPlace && !hasJob && !hasPay)
        return null;

    // WORK_PLACE is required in DB; placeholder when source is blank.
    string place = hasPlace ? T(workPlace, 200)! : "غير محدد";

    return new ProfileWorksTb
    {
        // WorkId / ProfileId assigned later
        WorkPlace      = place,
        WorkNatureId   = null,              // keep null per stakeholder
        Salary         = salary,
        SalaryPeriodId = 2,                 // شهري
        CreatedOn      = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
    };
}

// ════════════════════════════════════════════════════════════════════════════════
//  Bank information — PROFILE_BANK_INFORMATION_TB  (FK: ProfileId)
//
//  BANK_ACCOUNT_NUMBER → BANK_ACCOUNT_NO
//  Missing required fields (stakeholder defaults for migration):
//    BANK_ID         = 1
//    ACCOUNT_CURR_ID = 1
//    IBAN            = "MIG-{ProfileId}"  (unique; set after ProfileId is assigned)
//  Insert only when BANK_ACCOUNT_NUMBER is present.
// ════════════════════════════════════════════════════════════════════════════════
const int DefaultBankId      = 1;
const byte DefaultAccountCurrId = 1;

static ProfileBankInformationTb? MapBankInfo(IXLRow row, Dictionary<string, int> h)
{
    string? accountNo = GetString(row, h, "BANK_ACCOUNT_NUMBER");
    if (string.IsNullOrWhiteSpace(accountNo)) return null;

    return new ProfileBankInformationTb
    {
        // ProfileId / Iban assigned later
        BankId        = DefaultBankId,
        BankAccountNo = T(accountNo, 50),
        AccountCurrId = DefaultAccountCurrId,
        Iban          = "",                 // placeholder until ProfileId is known
        CreatedOn     = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
    };
}

// ════════════════════════════════════════════════════════════════════════════════
//  Partner mapping — PROFILES_PARTNERS_TB
//
//  Created only when PARTNER_NAME is present. ID_NUM is NOT NULL, so a row with
//  a name but no PARTNER_NATIONAL_ID is skipped and logged.
//
//  ID_TYPE_ID (partner constants differ from profile ones):
//    new 1=رقم تسجيل, 2=هوية, 3=مشتغل مرخص, 4=جواز سفر,
//        5=الرقم الوطني, 6=جواز سفر اردني
//    Source has no partner ID-type column; PARTNER_NATIONAL_ID ⇒ هوية ⇒ 2.
//
//  Defaults (no source): IsBankBorrower=لا(1), SharesCnt=0, ContributionPercent=0
//  PARTNER_WORK is intentionally unmapped (stakeholder status = r).
// ════════════════════════════════════════════════════════════════════════════════
static ProfilesPartnersTb? MapPartner(int rowNumber, IXLRow row, Dictionary<string, int> h)
{
    string? pname1 = GetString(row, h, "PARTNER_NAME");
    if (string.IsNullOrWhiteSpace(pname1)) return null;

    string? idNum = GetString(row, h, "PARTNER_NATIONAL_ID");
    if (string.IsNullOrWhiteSpace(idNum))
    {
        Console.Error.WriteLine(
            $"  [SKIP PARTNER] Row {rowNumber}: PARTNER_NAME='{pname1}' but PARTNER_NATIONAL_ID is missing — ID_NUM is required, not inserted.");
        return null;
    }

    string? pname2 = GetString(row, h, "PARTNER_NAME_2");
    string? pname4 = GetString(row, h, "PARTNER_NAME_4");
    string? pname5 = GetString(row, h, "PARTNER_NAME_5");

    return new ProfilesPartnersTb
    {
        // PartnerId / ProfileId assigned later
        FirstNameNa       = T(pname1, 50),
        FatherNameNa      = T(pname2, 50),
        GrandFatherNameNa = T(pname5, 50),
        FamilyNameNa      = T(pname4, 50),

        IdTypeId = 2,                          // PARTNER_NATIONAL_ID → هوية (new partner ID_TYPE = 2)
        IdNum    = T(idNum, 100)!,

        EducationLevelId = MapEducationLevel(GetInt(row, h, "PARTNER_EDUCATION")),

        PhoneNumber            = T(GetString(row, h, "PARTNER_MOBILE"), 30),
        CurrentExperienceNotes = GetString(row, h, "PARTNER_WORK_DESC"),

        // Stakeholder defaults — no source columns
        SharesCnt           = 0,
        ContributionPercent = 0,
        IsBankBorrower      = false,   // code converter: false → 1 (لا)

        IsHidden  = false,
        CreatedOn = GetDateTime(row, h, "CREATE_DATE") ?? DateTime.Now,
    };
}

// ════════════════════════════════════════════════════════════════════════════════
//  Merge helpers
// ════════════════════════════════════════════════════════════════════════════════

// Fold `older` into `target` (which holds the newer row's values).
//   - target empty, older has value -> copy it in (fills a gap, no conflict)
//   - both non-empty and unequal     -> keep target (newer wins) and record a conflict
//   - equal, or older empty          -> no change
// ProfileId/CustId are identity/sequence fields assigned after merging, so they
// are excluded from the field-by-field merge.
static void MergeInto(ProfilesTb target, ProfilesTb older, List<string> conflicts)
{
    foreach (var prop in typeof(ProfilesTb).GetProperties())
    {
        if (!prop.CanRead || !prop.CanWrite) continue;
        if (prop.Name is nameof(ProfilesTb.ProfileId) or nameof(ProfilesTb.CustId)) continue;

        object? tVal = prop.GetValue(target);
        object? oVal = prop.GetValue(older);

        bool tEmpty = IsEmpty(tVal);
        bool oEmpty = IsEmpty(oVal);

        if (tEmpty && !oEmpty)
        {
            prop.SetValue(target, oVal);   // fill gap from older row
        }
        else if (!tEmpty && !oEmpty && !Equals(tVal, oVal))
        {
            // Both filled and differ — newer (target) already wins; just log it.
            conflicts.Add($"{prop.Name}: newer='{tVal}' older='{oVal}'");
        }
    }
}

static bool IsEmpty(object? v) =>
    v is null || (v is string s && string.IsNullOrWhiteSpace(s));

static void WriteConflict(string path, string text) =>
    File.AppendAllText(path, text + "\r\n\r\n");

// ════════════════════════════════════════════════════════════════════════════════
//  DB helpers
// ════════════════════════════════════════════════════════════════════════════════

// Inserts the batch and returns the number of rows that committed successfully.
// On a batch failure it retries row-by-row so a single bad row doesn't lose the
// whole batch; each rejected row is logged and excluded from the returned count.
static async Task<int> InsertProfileBatchAsync(
    DbContextOptions<SilaDbContext> opts,
    List<ProfilesTb> batch)
{
    try
    {
        using var ctx = new SilaDbContext(opts);
        ctx.ProfilesTbs.AddRange(batch);
        await ctx.SaveChangesAsync();
        return batch.Count;
    }
    catch
    {
        // Batch failed — retry one by one so only the bad row is lost.
        int ok = 0;
        foreach (var profile in batch)
        {
            try
            {
                using var ctx = new SilaDbContext(opts);
                ctx.ProfilesTbs.Add(profile);
                await ctx.SaveChangesAsync();
                ok++;
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                Console.Error.WriteLine(
                    $"  [SKIP] ProfileId={profile.ProfileId}: {inner.Message}");
                Console.Error.WriteLine(
                    $"         BranchId={profile.BranchId}" +
                    $" MaritalStatusId={profile.MaritalStatusId}" +
                    $" GenderId={profile.GenderId}" +
                    $" CustTypeId={profile.CustTypeId}" +
                    $" Weight={profile.Weight}" +
                    $" Hight={profile.Hight}" +
                    $" ChildrenCnt={profile.ChildrenCnt}" +
                    $" DependentsCnt={profile.DependentsCnt}" +
                    $" FamilyMembersCnt={profile.FamilyMembersCnt}" +
                    $" IdTypeId={profile.IdTypeId}" +
                    $" GroupId={profile.GroupId}");
            }
        }
        return ok;
    }
}

// Inserts the work batch and returns how many committed.  Mirrors the profile
// inserter's row-by-row fallback: a work row whose profile was rejected fails its
// ProfileId FK and is logged/skipped rather than aborting the batch.
static async Task<int> InsertWorkBatchAsync(
    DbContextOptions<SilaDbContext> opts,
    List<ProfileWorksTb> batch)
{
    try
    {
        using var ctx = new SilaDbContext(opts);
        ctx.ProfileWorksTbs.AddRange(batch);
        await ctx.SaveChangesAsync();
        return batch.Count;
    }
    catch
    {
        int ok = 0;
        foreach (var work in batch)
        {
            try
            {
                using var ctx = new SilaDbContext(opts);
                ctx.ProfileWorksTbs.Add(work);
                await ctx.SaveChangesAsync();
                ok++;
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                Console.Error.WriteLine(
                    $"  [SKIP WORK] WorkId={work.WorkId} ProfileId={work.ProfileId}: {inner.Message}");
            }
        }
        return ok;
    }
}

// Inserts the partner batch and returns how many committed.  Same row-by-row
// fallback as works: a partner whose profile was rejected fails its PROFILE_ID
// FK and is logged/skipped rather than aborting the batch.
static async Task<int> InsertPartnerBatchAsync(
    DbContextOptions<SilaDbContext> opts,
    List<ProfilesPartnersTb> batch)
{
    try
    {
        using var ctx = new SilaDbContext(opts);
        ctx.ProfilesPartnersTbs.AddRange(batch);
        await ctx.SaveChangesAsync();
        return batch.Count;
    }
    catch
    {
        int ok = 0;
        foreach (var partner in batch)
        {
            try
            {
                using var ctx = new SilaDbContext(opts);
                ctx.ProfilesPartnersTbs.Add(partner);
                await ctx.SaveChangesAsync();
                ok++;
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                Console.Error.WriteLine(
                    $"  [SKIP PARTNER] PartnerId={partner.PartnerId} ProfileId={partner.ProfileId}: {inner.Message}");
            }
        }
        return ok;
    }
}

static async Task<int> InsertContactBatchAsync(
    DbContextOptions<SilaDbContext> opts,
    List<ProfileContactDtsTb> batch)
{
    try
    {
        using var ctx = new SilaDbContext(opts);
        ctx.ProfileContactDtsTbs.AddRange(batch);
        await ctx.SaveChangesAsync();
        return batch.Count;
    }
    catch
    {
        int ok = 0;
        foreach (var contact in batch)
        {
            try
            {
                using var ctx = new SilaDbContext(opts);
                ctx.ProfileContactDtsTbs.Add(contact);
                await ctx.SaveChangesAsync();
                ok++;
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                Console.Error.WriteLine(
                    $"  [SKIP CONTACT] ContactId={contact.ContactId} ProfileId={contact.ProfileId}: {inner.Message}");
            }
        }
        return ok;
    }
}

static async Task<int> InsertBankBatchAsync(
    DbContextOptions<SilaDbContext> opts,
    List<ProfileBankInformationTb> batch)
{
    try
    {
        using var ctx = new SilaDbContext(opts);
        ctx.ProfileBankInformationTbs.AddRange(batch);
        await ctx.SaveChangesAsync();
        return batch.Count;
    }
    catch
    {
        int ok = 0;
        foreach (var bank in batch)
        {
            try
            {
                using var ctx = new SilaDbContext(opts);
                ctx.ProfileBankInformationTbs.Add(bank);
                await ctx.SaveChangesAsync();
                ok++;
            }
            catch (Exception ex)
            {
                var inner = ex.InnerException ?? ex;
                Console.Error.WriteLine(
                    $"  [SKIP BANK] ProfileId={bank.ProfileId} BankId={bank.BankId} IBAN={bank.Iban}: {inner.Message}");
            }
        }
        return ok;
    }
}

// ════════════════════════════════════════════════════════════════════════════════
//  Cell accessors
// ════════════════════════════════════════════════════════════════════════════════

static Dictionary<string, int> BuildHeaderIndex(IXLRow headerRow) =>
    headerRow.CellsUsed()
             .ToDictionary(
                 c => c.Value.ToString().Trim(),
                 c => c.Address.ColumnNumber,
                 StringComparer.OrdinalIgnoreCase);

static int? GetInt(IXLRow row, Dictionary<string, int> h, string col)
{
    if (!h.TryGetValue(col, out int c)) return null;
    return GetCellInt(row.Cell(c));
}

static int? GetCellInt(IXLCell cell)
{
    if (cell.IsEmpty()) return null;
    if (cell.TryGetValue(out long l))    return (int)l;
    if (cell.TryGetValue(out double d))  return Convert.ToInt32(d);
    return null;
}

static string? GetString(IXLRow row, Dictionary<string, int> h, string col)
{
    if (!h.TryGetValue(col, out int c)) return null;
    var s = row.Cell(c).GetString().Trim();
    return string.IsNullOrWhiteSpace(s) ? null : s;
}

static DateTime? GetDateTime(IXLRow row, Dictionary<string, int> h, string col)
{
    if (!h.TryGetValue(col, out int c)) return null;
    var cell = row.Cell(c);
    return cell.TryGetValue(out DateTime dt) ? dt : null;
}

static byte? GetByte(IXLRow row, Dictionary<string, int> h, string col)
{
    int? v = GetInt(row, h, col);
    return v is >= 0 and <= 255 ? (byte)v.Value : null;
}

static decimal? GetDecimal(IXLRow row, Dictionary<string, int> h, string col)
{
    if (!h.TryGetValue(col, out int c)) return null;
    var cell = row.Cell(c);
    return cell.TryGetValue(out double d) ? (decimal)d : null;
}

// Source 0/1 flag -> bool? for the model's bool? flag columns (default EF
// converter stores false->0, true->1).  Null source stays null.
static bool? ExcelFlag(int? v) => v.HasValue ? v == 1 : null;

// Old ID_CARD_TYPE → new ID_TYPE_ID (by meaning):
//   old 1 هوية              → new 1 هوية
//   old 2 مشتغل مرخص        → new 2 مشتغل مرخص
//   old 3 جواز سفر فلسطيني  → new 3 جواز سفر
//   old 4 جواز سفر اردني    → new 5 جواز سفر اردني
//   new 4 الرقم الوطني has no old equivalent
static int? MapIdTypeId(int? oldCode) => oldCode switch
{
    1 => 1,
    2 => 2,
    3 => 3,
    4 => 5,
    _ => oldCode,
};

// MARITAL_STATUS 1-4 match new codes. Old 5=Engaged vs new 5=Separated —
// stakeholder mapping sheet passes the numeric value through as-is.
static byte? MapMaritalStatus(byte? v) => v;

// NO_OF_DEPENDENTS occasionally has implausible values (e.g. 31, 42) that are
// still valid bytes (0-255), so GetByte's range check alone wouldn't catch a
// bad entry. Per stakeholder, anything under 100 is plausible; 100+ is bad
// data and dropped to null rather than inserted as-is.
static byte? MapDependentsCnt(byte? v) => v.HasValue && v.Value < 100 ? v : null;

// NO_DEPENDENTS_CHILDREN had 2 rows at value 42 (CLIENT_ID 20931, 29031),
// flagged by the source data note itself as likely bad data - unlike
// NO_OF_DEPENDENTS (capped at 100), a family having 20+ children is already
// implausible, so this uses a tighter cap.
static byte? MapChildrenCnt(byte? v) => v.HasValue && v.Value < 20 ? v : null;

// QUALIFICATION_CODE / PARTNER_EDUCATION (old 1-7) → EDUCATION_LEVEL_ID (new 1-9).
// New system inserts "إعدادي" between Basic and Secondary, so codes shift:
//   old 1 اساسي     → new 2 أساسي
//   old 2 ثانوي     → new 4 ثانوي
//   old 3 دبلوم     → new 5 دبلوم
//   old 4 بكالوريوس → new 6 بكالوريوس
//   old 5 ماجستير   → new 7 ماجستير
//   old 6 دكتوراه   → new 8 دكتوراه
//   old 7 اي/لا يوجد → new 1 أمي
static int? MapEducationLevel(int? oldCode) => oldCode switch
{
    1 => 2, 2 => 4, 3 => 5, 4 => 6, 5 => 7, 6 => 8, 7 => 1,
    _ => null,
};

static string JoinParts(params string?[] parts) =>
    string.Join(" ", parts.Where(p => !string.IsNullOrWhiteSpace(p))).Trim();

static string FilterName(string? name)
{
    if (string.IsNullOrWhiteSpace(name)) return string.Empty;
    return System.Text.RegularExpressions.Regex
        .Replace(name, @"[^\p{L}\p{N}\s]", "").Trim();
}

static string? T(string? s, int maxLen) =>
    s is null ? null : s.Length <= maxLen ? s : s[..maxLen];

// ════════════════════════════════════════════════════════════════════════════════
//  DTOs
// ════════════════════════════════════════════════════════════════════════════════

// One source row, pre-mapped, awaiting merge within its (ID_NUM, ID_TYPE_ID) group.
// Work/Partner are null when the row has no WORK_PLACE/PARTNER_NAME (both optional).
record RowData(int RowNumber, int ClientId, string Company, DateTime SortDate,
               ProfilesTb Profile, ProfileWorksTb? Work, ProfilesPartnersTb? Partner,
               ProfileBankInformationTb? Bank, IReadOnlyList<string> Phones);

record AddressData(
    int? PermanentStateId,
    int? PermanentCityId,
    string? PermanentAddress,
    string? MailZip,
    string? MailPobox,
    IReadOnlyList<string> Phones);

record IdCardData
{
    public int?      IdTypeId   { get; init; }
    public string?   IdNum      { get; init; }
    public DateTime? IssueDate  { get; init; }
    public string?   IssuePlace { get; init; }
    public DateTime? ExpiryDate { get; init; }
}
