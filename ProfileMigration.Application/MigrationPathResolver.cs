using Microsoft.Extensions.Options;
using ProfileMigration.Application.Options;

namespace ProfileMigration.Application;

public sealed class MigrationPathResolver(IOptions<MigrationOptions> options)
{
    public sealed record ProfileExcelPaths(
        string ClientPath,
        string IdCardPath,
        string AddressPath);

    public string ResolveExcelPath(string? overridePath = null)
    {
        var o = options.Value;
        var path = IsUsableOverride(overridePath) ? overridePath! : o.ExcelFilePath;
        return Resolve(path, o.ContentRoot);
    }

    public ProfileExcelPaths ResolveProfileExcelPaths(string? clientOverridePath = null)
    {
        var o = options.Value;
        return new ProfileExcelPaths(
            ResolveExcelPath(clientOverridePath),
            Resolve(o.IdCardExcelFilePath, o.ContentRoot),
            Resolve(o.AddressExcelFilePath, o.ContentRoot));
    }

    public string ResolveAreaExcelPath(string? overridePath = null)
    {
        var o = options.Value;
        var path = IsUsableOverride(overridePath) ? overridePath! : o.AreaExcelFilePath;
        return Resolve(path, o.ContentRoot);
    }

    /// <summary>
    /// Ignores blank values and Swagger placeholders like "string".
    /// </summary>
    static bool IsUsableOverride(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return false;
        var t = path.Trim();
        if (t.Equals("string", StringComparison.OrdinalIgnoreCase)) return false;
        if (t.Equals("null", StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }

    static string Resolve(string path, string? contentRoot)
    {
        if (Path.IsPathRooted(path)) return path;
        var root = contentRoot ?? AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(root, path));
    }
}
