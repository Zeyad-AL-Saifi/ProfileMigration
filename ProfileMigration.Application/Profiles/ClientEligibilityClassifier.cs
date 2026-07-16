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

    public sealed record ClientIdentityRow(string Company, int ClientId, string? CardKey);

    public sealed class Result
    {
        public int TotalInput { get; init; }
        public int EligibleCount => Eligible.Count;
        public HashSet<(string Company, int ClientId)> Eligible { get; init; } = new();
        public Dictionary<(string Company, int ClientId), string> Skipped { get; init; } = new();

        public int SkippedMissingIdCard { get; init; }
        public int SkippedInternalDuplicates { get; init; }
        public int SkippedCrossCompanyMatches { get; init; }

        public bool IsEligible(string company, int clientId) =>
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

    public static string NormalizeCompany(string? company) =>
        (company ?? "").Trim().ToUpperInvariant();

    public static bool IsKnownCompany(string company) =>
        company is "ASALA" or "ACAD";

    public static Result Classify(IEnumerable<ClientIdentityRow> rows)
    {
        var input = new List<(string Company, int ClientId, string? CardKey)>();
        foreach (var r in rows)
        {
            var company = NormalizeCompany(r.Company);
            if (!IsKnownCompany(company) || r.ClientId <= 0) continue;
            input.Add((company, r.ClientId, r.CardKey));
        }

        // Dedupe by (company, clientId) — first wins (same as id-card join keep=first)
        var unique = new Dictionary<(string, int), string?>();
        foreach (var row in input)
        {
            var key = (row.Company, row.ClientId);
            if (!unique.ContainsKey(key))
                unique[key] = row.CardKey;
        }

        var skipped = new Dictionary<(string Company, int ClientId), string>();
        var missing = 0;

        foreach (var (key, cardKey) in unique)
        {
            if (string.IsNullOrWhiteSpace(cardKey))
            {
                skipped[key] = ReasonMissingIdCard;
                missing++;
            }
        }

        // Remaining with valid card_key
        var withCard = unique
            .Where(kv => !skipped.ContainsKey(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
            .Select(kv => (kv.Key.Item1, kv.Key.Item2, CardKey: kv.Value!))
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
        int internalDupCount = 0;
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

        var eligible = new HashSet<(string Company, int ClientId)>();
        foreach (var key in unique.Keys)
        {
            if (!skipped.ContainsKey(key))
                eligible.Add(key);
        }

        return new Result
        {
            TotalInput = unique.Count,
            Eligible = eligible,
            Skipped = skipped,
            SkippedMissingIdCard = missing,
            SkippedInternalDuplicates = internalDupCount,
            SkippedCrossCompanyMatches = cross,
        };
    }
}
