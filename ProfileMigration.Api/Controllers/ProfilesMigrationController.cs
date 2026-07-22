using Microsoft.AspNetCore.Http.Timeouts;
using Microsoft.AspNetCore.Mvc;
using ProfileMigration.Api.Forms;
using ProfileMigration.Application.Dtos;
using ProfileMigration.Application.Profiles;
using ProfileMigration.Application.Profiles.Commands;

namespace ProfileMigration.Api.Controllers;

[ApiController]
[Route("api/migration/profiles")]
[RequestSizeLimit(200_000_000)]
public sealed class ProfilesMigrationController(
    ProfileMigrationService service,
    ResetProfileMigrationDataCommandHandler resetHandler) : ControllerBase
{
    [HttpPost("validate")]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(typeof(MigrationReportDto), StatusCodes.Status200OK)]
    public async Task<ActionResult<MigrationReportDto>> Validate(
        [FromForm] ExcelMigrationForm? form, CancellationToken ct)
    {
        var (request, cleanup) = await MigrationFormFiles.ToExcelRequestAsync(form, ct);
        using (cleanup)
            return Ok(await service.ValidateAsync(request, ct));
    }

    [HttpPost("run")]
    [Consumes("multipart/form-data")]
    [RequestTimeout(milliseconds: 7_200_000)]
    [ProducesResponseType(typeof(PhaseRunResultDto), StatusCodes.Status200OK)]
    public async Task<ActionResult<PhaseRunResultDto>> Run(
        [FromForm] ExcelMigrationForm? form, CancellationToken ct)
    {
        var (request, cleanup) = await MigrationFormFiles.ToExcelRequestAsync(form, ct);
        using (cleanup)
            return Ok(await service.RunAsync(request, ct));
    }

    [HttpDelete("reset")]
    [ProducesResponseType(typeof(ResetProfileMigrationDataResult), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status400BadRequest)]
    public async Task<ActionResult<ResetProfileMigrationDataResult>> Reset(
        [FromQuery] bool confirm,
        CancellationToken ct)
    {
        if (!confirm)
        {
            return BadRequest(new ProblemDetails
            {
                Title = "Confirmation required",
                Detail = "Set confirm=true to delete all profile migration data.",
                Status = StatusCodes.Status400BadRequest,
            });
        }

        return Ok(await resetHandler.HandleAsync(
            new ResetProfileMigrationDataCommand(confirm),
            ct));
    }
}
