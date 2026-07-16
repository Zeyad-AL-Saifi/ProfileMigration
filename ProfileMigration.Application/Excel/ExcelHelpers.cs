using ClosedXML.Excel;

namespace ProfileMigration.Application.Excel;

public static class ExcelHelpers
{
    public const int SampleLimit = 20;

    public static Dictionary<string, int> BuildHeaderIndex(IXLRow headerRow) =>
        headerRow.CellsUsed()
                 .ToDictionary(
                     c => c.Value.ToString()!.Trim(),
                     c => c.Address.ColumnNumber,
                     StringComparer.OrdinalIgnoreCase);

    public static void RequireCol(Dictionary<string, int> h, string col)
    {
        if (!h.ContainsKey(col))
            throw new InvalidOperationException($"Excel missing required column '{col}'.");
    }

    public static int? GetInt(IXLRow row, Dictionary<string, int> h, string col)
    {
        if (!h.TryGetValue(col, out int c)) return null;
        return GetCellInt(row.Cell(c));
    }

    public static int? GetCellInt(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;
        if (cell.TryGetValue(out long l)) return (int)l;
        if (cell.TryGetValue(out double d)) return Convert.ToInt32(d);
        return null;
    }

    public static string? GetString(IXLRow row, Dictionary<string, int> h, string col)
    {
        if (!h.TryGetValue(col, out int c)) return null;
        var s = row.Cell(c).GetString().Trim();
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    public static DateTime? GetDateTime(IXLRow row, Dictionary<string, int> h, string col)
    {
        if (!h.TryGetValue(col, out int c)) return null;
        var cell = row.Cell(c);
        return cell.TryGetValue(out DateTime dt) ? dt : null;
    }

    public static byte? GetByte(IXLRow row, Dictionary<string, int> h, string col)
    {
        int? v = GetInt(row, h, col);
        return v is >= 0 and <= 255 ? (byte)v.Value : null;
    }

    public static decimal? GetDecimal(IXLRow row, Dictionary<string, int> h, string col)
    {
        if (!h.TryGetValue(col, out int c)) return null;
        var cell = row.Cell(c);
        return cell.TryGetValue(out double d) ? (decimal)d : null;
    }

    public static string? T(string? s, int maxLen) =>
        s is null ? null : s.Length <= maxLen ? s : s[..maxLen];

    public static string JoinParts(params string?[] parts) =>
        string.Join(" ", parts.Where(p => !string.IsNullOrWhiteSpace(p))).Trim();

    public static string FilterName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        return System.Text.RegularExpressions.Regex
            .Replace(name, @"[^\p{L}\p{N}\s]", "").Trim();
    }

    public static bool? ExcelFlag(int? v) => v.HasValue ? v == 1 : null;

    public static int? MapIdTypeId(int? oldCode) => oldCode switch
    {
        1 => 1,
        2 => 2,
        3 => 3,
        4 => 5,
        _ => oldCode,
    };

    public static byte? MapMaritalStatus(byte? v) => v;

    public static byte? MapDependentsCnt(byte? v) => v.HasValue && v.Value < 100 ? v : null;

    public static byte? MapChildrenCnt(byte? v) => v.HasValue && v.Value < 20 ? v : null;

    public static int? MapEducationLevel(int? oldCode) => oldCode switch
    {
        1 => 2, 2 => 4, 3 => 5, 4 => 6, 5 => 7, 6 => 8, 7 => 1,
        _ => null,
    };

    public static Dictionary<(string Company, int OldCode), int> LoadBranchIdMap() => new()
    {
        { ("ACAD",   0), 5 },
        { ("ACAD",  10), 4 },
        { ("ACAD",  20), 5 },
        { ("ACAD",  30), 6 },
        { ("ACAD",  40), 8 },
        { ("ACAD",  50), 9 },
        { ("ACAD",  60), 7 },
        { ("ACAD",  70), 5 },
        { ("ACAD",  80), 1 },
        { ("ACAD",  90), 2 },
        { ("ACAD", 100), 5 },
        { ("ACAD", 110), 3 },
        { ("ACAD", 120), 2 },
        { ("ACAD", 130), 4 },
        { ("ACAD", 200), 1 },
        { ("ACAD", 400), 8 },
        { ("ACAD", 999), 5 },
        { ("ASALA",   0), 5 },
        { ("ASALA",  10), 5 },
        { ("ASALA",  20), 1 },
        { ("ASALA",  30), 9 },
        { ("ASALA",  40), 5 },
        { ("ASALA",  50), 7 },
        { ("ASALA",  60), 4 },
        { ("ASALA",  70), 8 },
        { ("ASALA",  80), 9 },
        { ("ASALA",  90), 9 },
        { ("ASALA", 100), 9 },
        { ("ASALA", 110), 9 },
        { ("ASALA", 120), 9 },
        { ("ASALA", 130), 2 },
    };

    public static Dictionary<int, int> LoadCityIdMap() => new()
    {
        { 0,  -1 },
        { 1,   7 },
        { 2,   3 },
        { 3,   6 },
        { 4,  16 },
        { 5,   6 },
        { 6,   1 },
        { 7,  10 },
        { 8,   9 },
        { 9,  11 },
        { 10,  4 },
        { 11,  5 },
        { 12,  8 },
        { 13,  2 },
    };
}
