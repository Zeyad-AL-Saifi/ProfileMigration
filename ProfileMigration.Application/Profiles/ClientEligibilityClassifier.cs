namespace ProfileMigration.Application.Profiles;

/// <summary>
/// Pre-migration exclusion gate — same rules as Client Excel Analysis API
/// (<c>duplicate_detector.py</c>): missing card, internal company duplicates, cross-company card_key.
/// </summary>
public static class ClientEligibilityClassifier
{
    public const string ReasonMissingIdCard = "MISSING_ID_CARD";
    public const string ReasonInternalDuplicateAsala = "INTERNAL_DUPLICATE_ASALA";
    public const string ReasonInternalDuplicateAcad = "INTERNAL_DUPLICATE_ACAD";
    public const string ReasonCrossCompanyMatch = "CROSS_COMPANY_MATCH";
    public const string ReasonDuplicateIdNum = "DUPLICATE_ID_NUM";

    public sealed record ClientIdentityRow(
        string Company,
        long ClientId,
        string? CardKey,
        string? IdNum = null,
        int? IdType = null);

    public sealed class Result
    {
        public int TotalInput { get; init; }
        public int EligibleCount => Eligible.Count;
        public HashSet<(string Company, long ClientId)> Eligible { get; init; } = new();
        public Dictionary<(string Company, long ClientId), string> Skipped { get; init; } = new();

        public int SkippedMissingIdCard { get; init; }
        public int SkippedInternalDuplicates { get; init; }
        public int SkippedCrossCompanyMatches { get; init; }
        public int SkippedDuplicateIdNums { get; init; }

        /// <summary>Analysis-style count: ASALA rows × ACAD rows for every shared card key.</summary>
        public int CrossCompanyMatchedPairs { get; init; }

        /// <summary>Analysis-style count of records in internal duplicate card-key groups.</summary>
        public int TotalInternalDuplicateRecords { get; init; }

        public bool IsEligible(string company, long clientId) =>
            Eligible.Contains((NormalizeCompany(company), clientId));

        public IReadOnlyList<string> SampleSkipped(string reason, int limit = 20) =>
            Skipped
                .Where(kv => kv.Value == reason)
                .Take(limit)
                .Select(kv => $"{kv.Key.Company}#CLIENT_ID={kv.Key.ClientId}")
                .ToList();
    }

    /// <summary>
    /// Build card_key exactly like analysis: normalize(ID_CARD_TYPE) + "|" + trim(ID_CARD_NO).
    /// Null when either part is missing.
    /// </summary>
    public static string? BuildCardKey(int? rawIdCardType, string? rawIdCardNo)
    {
        if (rawIdCardType is null) return null;
        if (string.IsNullOrWhiteSpace(rawIdCardNo)) return null;
        return $"{rawIdCardType.Value}|{rawIdCardNo.Trim()}";
    }

    public static string BuildDbIdentityKey(string idNum, int idType) =>
        $"{idType}|{idNum.Trim().ToUpperInvariant()}";

    public static string NormalizeCompany(string? company) =>
        (company ?? "").Trim().ToUpperInvariant();

    public static bool IsKnownCompany(string company) =>
        company is "ASALA" or "ACAD";

    public static Result Classify(IEnumerable<ClientIdentityRow> rows)
    {
        var input = new List<(string Company, long ClientId, string? CardKey, string? IdNum, int? IdType)>();
        foreach (var r in rows)
        {
            var company = NormalizeCompany(r.Company);
            if (!IsKnownCompany(company) || r.ClientId <= 0) continue;
            input.Add((company, r.ClientId, r.CardKey, NormalizeIdNum(r.IdNum), r.IdType));
        }

        int crossCompanyMatchedPairs = input
            .Where(x => !string.IsNullOrWhiteSpace(x.CardKey))
            .GroupBy(x => x.CardKey!, StringComparer.Ordinal)
            .Sum(group =>
                group.Count(x => x.Company == "ASALA") *
                group.Count(x => x.Company == "ACAD"));

        int totalInternalDuplicateRecords = input
            .Where(x => !string.IsNullOrWhiteSpace(x.CardKey))
            .GroupBy(x => (x.Company, x.CardKey))
            .Where(group => group.Count() >= 2)
            .Sum(group => group.Count());

        // Dedupe by (company, clientId) — first wins (same as id-card join keep=first)
        var unique = new Dictionary<(string, long), (string? CardKey, string? IdNum, int? IdType)>();
        foreach (var row in input)
        {
            var key = (row.Company, row.ClientId);
            if (!unique.ContainsKey(key))
                unique[key] = (row.CardKey, row.IdNum, row.IdType);
        }

        var skipped = new Dictionary<(string Company, long ClientId), string>();
        var missing = 0;
        int internalDupCount = 0;

        // Repeated MF_CLIENT rows for the same client are duplicates too. Mark the
        // client key once, but count every source record just like the analysis API.
        foreach (var duplicate in input.GroupBy(x => (x.Company, x.ClientId)).Where(g => g.Count() >= 2))
        {
            var reason = duplicate.Key.Company == "ASALA"
                ? ReasonInternalDuplicateAsala
                : ReasonInternalDuplicateAcad;
            skipped[duplicate.Key] = reason;
            internalDupCount += duplicate.Count();
        }

        foreach (var (key, identity) in unique)
        {
            if (skipped.ContainsKey(key)) continue;
            if (string.IsNullOrWhiteSpace(identity.CardKey))
            {
                skipped[key] = ReasonMissingIdCard;
                missing++;
            }
        }

        // Remaining with valid card_key
        var withCard = unique
            .Where(kv => !skipped.ContainsKey(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value.CardKey))
            .Select(kv => (kv.Key.Item1, kv.Key.Item2, CardKey: kv.Value.CardKey!))
            .ToList();

        // Cross-company: card_key present in BOTH ASALA and ACAD → exclude ALL sides
        var byCard = withCard.GroupBy(x => x.CardKey, StringComparer.Ordinal);
        var crossKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var g in byCard)
        {
            bool hasAsala = g.Any(x => x.Item1 == "ASALA");
            bool hasAcad = g.Any(x => x.Item1 == "ACAD");
            if (hasAsala && hasAcad)
                crossKeys.Add(g.Key);
        }

        int cross = 0;
        foreach (var row in withCard)
        {
            if (!crossKeys.Contains(row.CardKey)) continue;
            var key = (row.Item1, row.Item2);
            if (skipped.ContainsKey(key)) continue;
            skipped[key] = ReasonCrossCompanyMatch;
            cross++;
        }

        // Internal duplicates within one company (card_key count >= 2)
        foreach (var company in new[] { "ASALA", "ACAD" })
        {
            var reason = company == "ASALA" ? ReasonInternalDuplicateAsala : ReasonInternalDuplicateAcad;
            var companyRows = withCard.Where(x => x.Item1 == company).ToList();
            foreach (var g in companyRows.GroupBy(x => x.CardKey, StringComparer.Ordinal))
            {
                if (g.Count() < 2) continue;
                // Skip groups already fully covered by cross-company (still count only newly tagged)
                foreach (var row in g)
                {
                    var key = (row.Item1, row.Item2);
                    if (skipped.ContainsKey(key)) continue;
                    skipped[key] = reason;
                    internalDupCount++;
                }
            }
        }

        // A second identity guard prevents the old merge behavior. If two otherwise
        // eligible clients resolve to the same DB identity, exclude every source row.
        // This also catches NATIONAL_ID fallback collisions that card_key cannot see.
        int duplicateIdNums = 0;
        var duplicateDbKeys = unique
            .Where(kv => !string.IsNullOrWhiteSpace(kv.Value.IdNum))
            .GroupBy(
                kv => (IdNum: kv.Value.IdNum!, IdType: kv.Value.IdType ?? 1),
                new DbIdentityComparer())
            .Where(g => g.Count() >= 2);

        foreach (var group in duplicateDbKeys)
        {
            foreach (var row in group)
            {
                if (skipped.ContainsKey(row.Key)) continue;
                skipped[row.Key] = ReasonDuplicateIdNum;
                duplicateIdNums++;
            }
        }

        var eligible = new HashSet<(string Company, long ClientId)>();
        foreach (var key in unique.Keys)
        {
            if (!skipped.ContainsKey(key))
                eligible.Add(key);
        }

        return new Result
        {
            TotalInput = input.Count,
            Eligible = eligible,
            Skipped = skipped,
            SkippedMissingIdCard = missing,
            SkippedInternalDuplicates = internalDupCount,
            SkippedCrossCompanyMatches = cross,
            SkippedDuplicateIdNums = duplicateIdNums,
            CrossCompanyMatchedPairs = crossCompanyMatchedPairs,
            TotalInternalDuplicateRecords = totalInternalDuplicateRecords,
        };
    }

    static string? NormalizeIdNum(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    sealed class DbIdentityComparer : IEqualityComparer<(string IdNum, int IdType)>
    {
        public bool Equals((string IdNum, int IdType) x, (string IdNum, int IdType) y) =>
            x.IdType == y.IdType &&
            StringComparer.OrdinalIgnoreCase.Equals(x.IdNum, y.IdNum);

        public int GetHashCode((string IdNum, int IdType) obj) =>
            HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(obj.IdNum), obj.IdType);
    }
}
