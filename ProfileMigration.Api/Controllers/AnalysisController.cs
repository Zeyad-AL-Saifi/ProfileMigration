using Microsoft.AspNetCore.Mvc;
using ProfileMigration.Api.Forms;
using ProfileMigration.Application.Analysis;

namespace ProfileMigration.Api.Controllers;

[ApiController]
public sealed class AnalysisController(ClientAnalysisService service) : ControllerBase
{
    [HttpGet("/health")]
    public IActionResult Health()
    {
        bool ready = service.IsReady(out string idCardPath);
        return Ok(new
        {
            status = ready ? "healthy" : "unhealthy",
            id_card_loaded = ready,
            id_card_path = idCardPath,
        });
    }

    [HttpPost("/api/v1/analyze")]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(200_000_000)]
    public async Task<IActionResult> Analyze([FromForm] ExcelMigrationForm? form, CancellationToken ct)
    {
        if (ValidateUpload(form) is { } error) return error;
        var (request, cleanup) = await MigrationFormFiles.ToExcelRequestAsync(form, ct);
        using (cleanup)
        {
            try
            {
                return Ok(service.Analyze(request.ExcelPath
                    ?? throw new InvalidDataException("An Excel file is required.")));
            }
            catch (InvalidDataException ex)
            {
                return BadRequest(new { detail = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { detail = $"Analysis failed: {ex.Message}" });
            }
        }
    }

    [HttpPost("/api/v1/match-analysis")]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(200_000_000)]
    public async Task<IActionResult> MatchAnalysis(
        [FromForm] ExcelMigrationForm? form,
        [FromQuery] int? skip,
        [FromQuery] int? take,
        CancellationToken ct)
    {
        if (skip is < 0) return BadRequest(new { detail = "skip must be zero or greater." });
        if (take is < 1) return BadRequest(new { detail = "take must be one or greater." });
        if (ValidateUpload(form) is { } error) return error;

        var (request, cleanup) = await MigrationFormFiles.ToExcelRequestAsync(form, ct);
        using (cleanup)
        {
            try
            {
                return Ok(service.AnalyzeMatches(
                    request.ExcelPath ?? throw new InvalidDataException("An Excel file is required."),
                    skip ?? 0,
                    take));
            }
            catch (InvalidDataException ex)
            {
                return BadRequest(new { detail = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { detail = $"Analysis failed: {ex.Message}" });
            }
        }
    }

    BadRequestObjectResult? ValidateUpload(ExcelMigrationForm? form)
    {
        var file = form?.File is { Length: > 0 } ? form.File : form?.ExcelFile;
        if (file is null || file.Length == 0)
            return BadRequest(new { detail = "An Excel file is required." });

        string extension = Path.GetExtension(file.FileName);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { detail = "Uploaded file must be an Excel file (.xlsx or .xls)." });

        return null;
    }
}
