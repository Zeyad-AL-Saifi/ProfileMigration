using Microsoft.AspNetCore.Http.Timeouts;
using Microsoft.AspNetCore.Mvc;
using ProfileMigration.Api.Forms;
using ProfileMigration.Application.Constants;
using ProfileMigration.Application.Dtos;

namespace ProfileMigration.Api.Controllers;

/// <summary>
/// Unified constants phase: branches + ID-type constants (هوية / جواز أردني) + areas.
/// </summary>
[ApiController]
[Route("api/migration/constants")]
[RequestSizeLimit(200_000_000)]
public sealed class ConstantsMigrationController(ConstantsPhaseService service) : ControllerBase
{
    /// <summary>Validate branches + ID constants + areas in one report.</summary>
    [HttpPost("validate")]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(typeof(MigrationReportDto), StatusCodes.Status200OK)]
    public async Task<ActionResult<MigrationReportDto>> Validate(
        [FromForm] ConstantsMigrationForm? form, CancellationToken ct)
    {
        var (request, cleanup) = await MigrationFormFiles.ToAreaRequestAsync(form, ct);
        using (cleanup)
            return Ok(await service.ValidateAsync(request, ct));
    }

    /// <summary>
    /// Run in order: branches → ID constants → areas replace (MAIN_ID=10).
    /// Form fields: AreaExcelFile (optional) or AreaExcelPath (optional).
    /// </summary>
    [HttpPost("run")]
    [Consumes("multipart/form-data")]
    [RequestTimeout(milliseconds: 7_200_000)]
    [ProducesResponseType(typeof(PhaseRunResultDto), StatusCodes.Status200OK)]
    public async Task<ActionResult<PhaseRunResultDto>> Run(
        [FromForm] ConstantsMigrationForm? form, CancellationToken ct)
    {
        var (request, cleanup) = await MigrationFormFiles.ToAreaRequestAsync(form, ct);
        using (cleanup)
            return Ok(await service.RunAsync(request, ct));
    }
}
