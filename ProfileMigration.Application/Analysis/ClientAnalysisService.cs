using ClosedXML.Excel;
using ProfileMigration.Application.Profiles;
using static ProfileMigration.Application.Excel.ExcelHelpers;
using static ProfileMigration.Application.Profiles.ProfileExcelSource;

namespace ProfileMigration.Application.Analysis;

/// <summary>
/// .NET port of the Code Python analysis API. Analysis and migration share
/// ClientEligibilityClassifier, so duplicate counts and exclusions cannot drift.
/// </summary>
public sealed class ClientAnalysisService(MigrationPathResolver paths)
{
    const string ClientSheetName = "Match MF_CLIENT";
    static readonly HashSet<string> NonComparableColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "COMPANY", "CLIENT_ID", "ID_CARD_TYPE", "ID_CARD_NO", "card_key",
        "CREATE_DATE", "MODIFY_DATE", "Combined Number", "CREATE_USER", "MODIFY_USER",
        "CREATE_TERMINAL", "MODIFY_TERMINAL", "CLIENT_TYPE_DESC_E", "CLIENT_TYPE_DESC_A",
        "COMPANY_BRANCH_DESC_A", "COMPANY_BRANCH_DESC_E", "NAME5", "ANAME5",
        "MARRIED_FLAG", "ORACLE_CITIZEN_ID", "PARTNER_MOTHER_NAME", "PARTNER_WORK",
        "PARTNER_WORK_NAME", "HA_CITY_CODE", "HA_AREA_CODE", "HA_SUBURB_CODE",
        "HA_ADDRESS_1", "HA_ADDRESS_2", "HA_ADDRESS_3", "HA_STREET", "HA_FAMOUS_PLACE",
        "HA_TEL1", "HA_TEL2", "HA_MOBILE", "HA_FAX", "HA_ZIP_CODE", "HA_POBOX",
        "HA_NO", "HA_RENT_FLAG", "HA_EMAIL", "HA_WEBSITE", "NO_YEARS_HOUSE",
        "NO_MONTHS_REMAIN", "REGISTRATION_ID_1", "REGISTRATION_ID_2", "REG_ISSUE_PLACE",
        "RENT_AMT", "NO_WORKERS", "FAMILY_PROVIDER_RMRK", "FAMILY_PROVIDER_FLAG",
        "CHECK_FLAG", "OWNER_NAME", "SPOUSE_CLIENT_ID", "FAMILY_ADDRESS", "FAMILY_TEL",
        "RETIREMENT_NO", "MILITARY_NO", "WEIGHT", "LENGTH", "CHRONIC",
        "ENTRY_EXCEPTION_FLAG", "FLEX_FIELD4", "FLEX_FIELD5", "FLEX_FIELD6",
        "ORACLE_GROUP_ID", "CITY_BIRTH_CODE", "CITIZENSHIP_CODE", "MONTHLY_EXPENSES",
        "NET_INCOME", "CITIZEN_ACC_NO", "NO_DEPENDENTS_WIVES", "MAIN_GROUP_ID",
        "GROUP_TYPE", "CBOS_ID", "EMPLOYER", "SAVING_FLAG", "RELIGION", "DEATH_REASON",
        "ORACLE_COMPANY_ID", "IRTEQAA_FLAG", "TRAINING_CYCLE_NAME", "TRAINING_CYCLE_DATE",
        "PC_RECEVING_DATE", "SOCIAL_SECUR_FLAG", "SOCIAL_SITE_FLAG", "SOCIAL_SITE",
        "SOCIAL_SECUR", "PARENTS_NAME1", "PARENTS_NAME2", "PARENTS_NAME3", "PARENTS_NAME4",
        "PARENTS_NAT_ID", "PARENTS_TEL_NO", "PARENTS_RELATION_CODE",
        "PARENTS_MONTHLY_INCOME", "HIJRI_BIRTH_DATE", "SOCIAL_SEC_FLAG",
        "SOCIAL_INSURANCE_AMT", "PARENTS_ID", "PARENTS_FLAG", "RATION_CARD",
        "SUPPLIER_FLAG", "IDP_FLAG", "HOUSEHOLD_FLA", "COMP_RELATIONSHIP",
        "CLIENT_WORK_DESC",
    };

    sealed record AnalysisRow(
        string Company,
        long ClientId,
        int? CardType,
        string? CardNo,
        string? CardKey,
        Dictionary<string, object?> Values);

    public bool IsReady(out string idCardPath)
    {
        idCardPath = paths.ResolveProfileExcelPaths().IdCardPath;
        if (!File.Exists(idCardPath)) return false;
        try
        {
            using var workbook = new XLWorkbook(idCardPath);
            _ = LoadIdCards(workbook);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public Dictionary<string, object?> Analyze(string clientPath)
    {
        var data = Load(clientPath);
        var internalAsala = BuildInternalGroups(data.Rows, "ASALA");
        var internalAcad = BuildInternalGroups(data.Rows, "ACAD");
        var groups = internalAsala.Concat(internalAcad).ToList();
        int duplicateRecords = groups.Sum(g => Convert.ToInt32(g["count"]));
        var pairs = BuildCrossPairs(data.Rows);

        return new()
        {
            ["summary"] = new Dictionary<string, object?>
            {
                ["total_clients"] = data.Rows.Count,
                ["asala_count"] = data.Rows.Count(r => r.Company == "ASALA"),
                ["acad_count"] = data.Rows.Count(r => r.Company == "ACAD"),
                ["clients_with_id_card"] = data.Rows.Count(r => r.CardKey is not null),
                ["clients_without_id_card"] = data.Rows.Count(r => r.CardKey is null),
                ["cross_company_matched"] = pairs.Count,
            },
            ["duplicates"] = new Dictionary<string, object?>
            {
                ["internal_asala"] = internalAsala,
                ["internal_acad"] = internalAcad,
                ["problems_count"] = duplicateRecords,
                ["total_internal_duplicate_records"] = duplicateRecords,
                ["same_person_count"] = groups.Count(g => (bool)g["is_same_person"]!),
                ["different_person_count"] = groups.Count(g => !(bool)g["is_same_person"]!),
            },
        };
    }

    public Dictionary<string, object?> AnalyzeMatches(string clientPath, int skip, int? take)
    {
        var data = Load(clientPath);
        var comparable = data.Headers.Where(h => !NonComparableColumns.Contains(h)).ToList();
        var scored = BuildCrossPairs(data.Rows)
            .Select(pair =>
            {
                var left = (AnalysisRow)pair.Left;
                var right = (AnalysisRow)pair.Right;
                var score = Score(left, right, comparable);
                return (left, right, score);
            })
            .ToList();

        var selected = (take.HasValue ? scored.Skip(skip).Take(take.Value) : scored.Skip(skip)).ToList();
        var clientPairs = selected.Select(x => new Dictionary<string, object?>
        {
            ["card_type"] = x.left.CardType,
            ["card_no"] = x.left.CardNo,
            ["asala_client_id"] = x.left.ClientId,
            ["asala_create_date"] = Serialize(GetValue(x.left, "CREATE_DATE")),
            ["asala_modify_date"] = Serialize(GetValue(x.left, "MODIFY_DATE")),
            ["acad_client_id"] = x.right.ClientId,
            ["acad_create_date"] = Serialize(GetValue(x.right, "CREATE_DATE")),
            ["acad_modify_date"] = Serialize(GetValue(x.right, "MODIFY_DATE")),
            ["newer_client_id"] = ResolveNewer(x.left, x.right)?.ClientId,
            ["newer_company"] = ResolveNewer(x.left, x.right)?.Company,
            ["match_percentage"] = x.score.Percentage,
            ["is_same_person"] = x.score.Percentage >= 100,
            ["matched_fields"] = x.score.Matched,
            ["total_fields"] = x.score.Total,
            ["different_fields"] = x.score.Differences,
        }).ToList();

        return new()
        {
            ["pagination"] = new Dictionary<string, object?>
            {
                ["total"] = scored.Count,
                ["skip"] = skip,
                ["take"] = take,
                ["returned"] = clientPairs.Count,
                ["has_more"] = skip + clientPairs.Count < scored.Count,
            },
            ["bucket_summary"] = new Dictionary<string, object?>
            {
                ["perfect_100"] = scored.Count(x => x.score.Percentage == 100),
                ["high_75_99"] = scored.Count(x => x.score.Percentage >= 75 && x.score.Percentage < 100),
                ["medium_50_74"] = scored.Count(x => x.score.Percentage >= 50 && x.score.Percentage < 75),
                ["low_below_50"] = scored.Count(x => x.score.Percentage < 50),
            },
            ["client_pairs"] = clientPairs,
        };
    }

    (List<AnalysisRow> Rows, List<string> Headers, ClientEligibilityClassifier.Result Eligibility) Load(string clientPath)
    {
        var idCardPath = paths.ResolveProfileExcelPaths(clientPath).IdCardPath;
        if (!File.Exists(clientPath)) throw new FileNotFoundException("Client Excel file not found.", clientPath);
        if (!File.Exists(idCardPath)) throw new FileNotFoundException("ID-card Excel file not found.", idCardPath);

        using var clientWb = new XLWorkbook(clientPath);
        using var idCardWb = new XLWorkbook(idCardPath);
        var ws = clientWb.Worksheets.FirstOrDefault(w =>
            string.Equals(w.Name, ClientSheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException($"Sheet '{ClientSheetName}' was not found.");
        var headers = BuildHeaderIndex(ws.Row(1));
        var missing = MissingColumns(headers, "COMPANY", "CLIENT_ID").ToList();
        if (missing.Count > 0)
            throw new InvalidDataException($"Missing required columns: {string.Join(", ", missing)}");

        var idCards = LoadIdCards(idCardWb);
        var rows = new List<AnalysisRow>();
        var identities = new List<ClientEligibilityClassifier.ClientIdentityRow>();
        int last = ws.LastRowUsed()?.RowNumber() ?? 1;

        for (int rowNumber = 2; rowNumber <= last; rowNumber++)
        {
            var row = ws.Row(rowNumber);
            string company = ClientEligibilityClassifier.NormalizeCompany(GetString(row, headers, "COMPANY"));
            var clientIdCell = row.Cell(headers["CLIENT_ID"]);
            long? clientId = GetLong(row, headers, "CLIENT_ID");
            if (!ClientEligibilityClassifier.IsKnownCompany(company))
                throw new InvalidDataException($"Unsupported COMPANY '{company}' at row {rowNumber}.");
            if (clientId is null || clientId <= 0)
                throw new InvalidDataException(
                    $"Invalid CLIENT_ID at row {rowNumber}. " +
                    $"COMPANY='{company}', value='{DisplayCellValue(clientIdCell)}', " +
                    $"reason='{DescribeInvalidClientId(clientIdCell)}'.");

            idCards.TryGetValue((company, clientId.Value), out var card);
            var values = headers.ToDictionary(
                h => h.Key,
                h => CellValue(row.Cell(h.Value)),
                StringComparer.OrdinalIgnoreCase);
            var analysisRow = new AnalysisRow(
                company, clientId.Value, card?.RawIdCardType, card?.RawIdCardNo,
                card?.CardKey, values);
            rows.Add(analysisRow);
            identities.Add(new(
                company, clientId.Value, card?.CardKey, card?.IdNum, card?.IdTypeId ?? 1));
        }

        return (rows, headers.Keys.ToList(), ClientEligibilityClassifier.Classify(identities));
    }

    static string DisplayCellValue(IXLCell cell)
    {
        if (cell.IsEmpty()) return "<empty>";
        var value = cell.GetFormattedString().Trim();
        return string.IsNullOrEmpty(value) ? "<blank>" : value;
    }

    static string DescribeInvalidClientId(IXLCell cell)
    {
        if (cell.IsEmpty() || string.IsNullOrWhiteSpace(cell.GetFormattedString()))
            return "CLIENT_ID is empty";

        if (cell.TryGetValue<double>(out var number))
        {
            if (double.IsNaN(number) || double.IsInfinity(number))
                return "CLIENT_ID is not a finite number";
            if (number != Math.Truncate(number))
                return "CLIENT_ID must be a whole number without decimals";
            if (number < long.MinValue || number > long.MaxValue)
                return "CLIENT_ID is outside the supported Int64 range";
            if (number <= 0)
                return "CLIENT_ID must be greater than zero";
        }

        return "CLIENT_ID is not a valid integer";
    }

    static List<Dictionary<string, object?>> BuildInternalGroups(
        List<AnalysisRow> rows,
        string company)
    {
        var comparable = rows.SelectMany(r => r.Values.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(h => !NonComparableColumns.Contains(h)).ToList();

        return rows
            .Where(r => r.Company == company && r.CardKey is not null)
            .GroupBy(r => r.CardKey!, StringComparer.Ordinal)
            .Where(group => group.Count() >= 2)
            .Select(group =>
            {
                var comparisons = new List<Dictionary<string, object?>>();
                var items = group.ToList();
                for (int i = 0; i < items.Count; i++)
                for (int j = i + 1; j < items.Count; j++)
                    comparisons.Add(BuildPairComparison(items[i], items[j], comparable));
                double percentage = comparisons.Count == 0
                    ? 0
                    : Math.Round(comparisons.Average(c => Convert.ToDouble(c["match_percentage"])), 2);

                return new Dictionary<string, object?>
                {
                    ["card_type"] = items[0].CardType,
                    ["card_no"] = items[0].CardNo,
                    ["count"] = items.Count,
                    ["client_ids"] = items.Select(r => r.ClientId).ToList(),
                    ["companies"] = new[] { company },
                    ["match_percentage"] = percentage,
                    ["is_same_person"] = comparisons.Count > 0 &&
                                         comparisons.All(c => (bool)c["is_same_person"]!),
                    ["pair_comparisons"] = comparisons,
                };
            }).ToList();
    }

    static Dictionary<string, object?> BuildPairComparison(
        AnalysisRow left, AnalysisRow right, List<string> comparable)
    {
        var score = Score(left, right, comparable);
        return new()
        {
            ["client_id_1"] = left.ClientId,
            ["create_date_1"] = Serialize(GetValue(left, "CREATE_DATE")),
            ["modify_date_1"] = Serialize(GetValue(left, "MODIFY_DATE")),
            ["client_id_2"] = right.ClientId,
            ["create_date_2"] = Serialize(GetValue(right, "CREATE_DATE")),
            ["modify_date_2"] = Serialize(GetValue(right, "MODIFY_DATE")),
            ["newer_client_id"] = ResolveNewer(left, right)?.ClientId,
            ["match_percentage"] = score.Percentage,
            ["is_same_person"] = score.Percentage >= 100,
            ["different_fields"] = score.Differences.Select(d => new Dictionary<string, object?>
            {
                ["field"] = d["field"],
                ["client_id_1"] = left.ClientId,
                ["value_1"] = d["asala_value"],
                ["client_id_2"] = right.ClientId,
                ["value_2"] = d["acad_value"],
            }).ToList(),
        };
    }

    static List<(object Left, object Right)> BuildCrossPairs(List<AnalysisRow> rows)
    {
        var pairs = new List<(object, object)>();
        foreach (var group in rows.Where(r => r.CardKey is not null)
                     .GroupBy(r => r.CardKey!, StringComparer.Ordinal))
        {
            var asala = group.Where(r => r.Company == "ASALA").ToList();
            var acad = group.Where(r => r.Company == "ACAD").ToList();
            foreach (var left in asala)
            foreach (var right in acad)
                pairs.Add((left, right));
        }
        return pairs;
    }

    static (int Matched, int Total, double Percentage, List<Dictionary<string, object?>> Differences)
        Score(AnalysisRow left, AnalysisRow right, List<string> fields)
    {
        int matched = 0;
        var differences = new List<Dictionary<string, object?>>();
        foreach (var field in fields)
        {
            object? leftValue = GetValue(left, field);
            object? rightValue = GetValue(right, field);
            if (Normalize(leftValue) == Normalize(rightValue))
            {
                matched++;
                continue;
            }
            differences.Add(new()
            {
                ["field"] = field,
                ["asala_value"] = Serialize(leftValue),
                ["acad_value"] = Serialize(rightValue),
            });
        }
        double percentage = fields.Count == 0 ? 0 : Math.Round(matched * 100d / fields.Count, 2);
        return (matched, fields.Count, percentage, differences);
    }

    static AnalysisRow? ResolveNewer(AnalysisRow left, AnalysisRow right)
    {
        DateTime? leftDate = EffectiveDate(left);
        DateTime? rightDate = EffectiveDate(right);
        if (leftDate == rightDate) return null;
        if (leftDate is null) return right;
        if (rightDate is null) return left;
        return leftDate > rightDate ? left : right;
    }

    static DateTime? EffectiveDate(AnalysisRow row) =>
        AsDate(GetValue(row, "MODIFY_DATE")) ?? AsDate(GetValue(row, "CREATE_DATE"));

    static DateTime? AsDate(object? value) =>
        value switch
        {
            DateTime dt => dt,
            DateTimeOffset dto => dto.DateTime,
            _ when DateTime.TryParse(Convert.ToString(value), out var parsed) => parsed,
            _ => null,
        };

    static object? GetValue(AnalysisRow row, string column) =>
        row.Values.TryGetValue(column, out var value) ? value : null;

    static object? CellValue(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;
        if (cell.TryGetValue<DateTime>(out var date)) return date;
        if (cell.TryGetValue<long>(out var integer)) return integer;
        if (cell.TryGetValue<double>(out var number)) return number;
        if (cell.TryGetValue<bool>(out var boolean)) return boolean;
        return cell.GetFormattedString();
    }

    static string Normalize(object? value) =>
        value switch
        {
            null => "",
            DateTime date => date.ToString("yyyy-MM-dd"),
            double number when number == Math.Truncate(number) => ((long)number).ToString(),
            _ => Convert.ToString(value)?.Trim().ToLowerInvariant() ?? "",
        };

    static object? Serialize(object? value) =>
        value switch
        {
            DateTime date => date.ToString("O"),
            double number when number == Math.Truncate(number) => (long)number,
            _ => value,
        };
}
