using ProfileMigration.Application.Dtos;

namespace ProfileMigration.Api.Forms;

/// <summary>Form-data for profile phases. The upload contains the MF_CLIENT worksheet only.</summary>
public sealed class ExcelMigrationForm
{
    /// <summary>
    /// Optional MF_CLIENT Excel upload. ID-card and address sources come from configured reference files.
    /// If omitted, uses ExcelFilePath from appsettings.
    /// </summary>
    public IFormFile? ExcelFile { get; set; }
}

/// <summary>Form-data for constants phase (areas Excel only; branches/id-types need no file).</summary>
public sealed class ConstantsMigrationForm
{
    /// <summary>Optional area Excel upload. If omitted, uses AreaExcelFilePath from appsettings.</summary>
    public IFormFile? AreaExcelFile { get; set; }
}

public static class MigrationFormFiles
{
    public static async Task<(ExcelMigrationRequest Request, IDisposable? Cleanup)> ToExcelRequestAsync(
        ExcelMigrationForm? form, CancellationToken ct)
    {
        form ??= new ExcelMigrationForm();
        if (form.ExcelFile is { Length: > 0 } file && IsRealUpload(file))
        {
            var (path, cleanup) = await SaveTempAsync(file, ct);
            return (new ExcelMigrationRequest { ExcelPath = path }, cleanup);
        }

        // No file → use appsettings (ExcelPath stays null)
        return (new ExcelMigrationRequest(), null);
    }

    public static async Task<(AreaMigrationRequest Request, IDisposable? Cleanup)> ToAreaRequestAsync(
        ConstantsMigrationForm? form, CancellationToken ct)
    {
        form ??= new ConstantsMigrationForm();
        if (form.AreaExcelFile is { Length: > 0 } file && IsRealUpload(file))
        {
            var (path, cleanup) = await SaveTempAsync(file, ct);
            return (new AreaMigrationRequest { AreaExcelPath = path }, cleanup);
        }

        return (new AreaMigrationRequest(), null);
    }

    /// <summary>
    /// Swagger often posts an empty file part or a placeholder; treat those as "not sent".
    /// </summary>
    static bool IsRealUpload(IFormFile file)
    {
        if (file.Length <= 0) return false;
        var name = file.FileName?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(name)) return false;
        if (name.Equals("string", StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }

    static async Task<(string Path, IDisposable Cleanup)> SaveTempAsync(IFormFile file, CancellationToken ct)
    {
        var ext = Path.GetExtension(file.FileName);
        if (string.IsNullOrWhiteSpace(ext)) ext = ".xlsx";

        var path = Path.Combine(Path.GetTempPath(), $"mig-{Guid.NewGuid():N}{ext}");
        await using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await file.CopyToAsync(fs, ct);
        }

        return (path, new TempFileCleanup(path));
    }

    sealed class TempFileCleanup(string path) : IDisposable
    {
        public void Dispose()
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // best-effort cleanup
            }
        }
    }
}
