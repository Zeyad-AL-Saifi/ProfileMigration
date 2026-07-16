using Microsoft.AspNetCore.Http.Timeouts;
using Microsoft.AspNetCore.Mvc;
using ProfileMigration.Api.Forms;
using ProfileMigration.Application.Dtos;
using ProfileMigration.Application.Profiles;

namespace ProfileMigration.Api.Controllers;

[ApiController]
[Route("api/migration/profilesContacts")]
[RequestSizeLimit(200_000_000)]
public sealed class ContactsMigrationController(ContactMigrationService service) : ControllerBase
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
}
