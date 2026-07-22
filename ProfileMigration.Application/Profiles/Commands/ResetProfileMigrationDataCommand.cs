using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application.Profiles.Commands;

public sealed record ResetProfileMigrationDataCommand(bool Confirm);

public sealed record ResetProfileMigrationDataResult(
    int ContactsDeleted,
    int BankInformationDeleted,
    int WorksDeleted,
    int PartnersDeleted,
    int ProfilesDeleted)
{
    public int TotalDeleted =>
        ContactsDeleted +
        BankInformationDeleted +
        WorksDeleted +
        PartnersDeleted +
        ProfilesDeleted;
}

public sealed class ResetProfileMigrationDataCommandHandler(
    SilaDbContext dbContext,
    ILogger<ResetProfileMigrationDataCommandHandler> logger)
{
    public async Task<ResetProfileMigrationDataResult> HandleAsync(
        ResetProfileMigrationDataCommand command,
        CancellationToken ct = default)
    {
        if (!command.Confirm)
            throw new ArgumentException(
                "Explicit confirmation is required to delete profile migration data.",
                nameof(command));

        await using var transaction = await dbContext.Database.BeginTransactionAsync(ct);

        try
        {
            int contactsDeleted = await dbContext.ProfileContactDtsTbs.ExecuteDeleteAsync(ct);
            int bankInformationDeleted = await dbContext.ProfileBankInformationTbs.ExecuteDeleteAsync(ct);
            int worksDeleted = await dbContext.ProfileWorksTbs.ExecuteDeleteAsync(ct);
            int partnersDeleted = await dbContext.ProfilesPartnersTbs.ExecuteDeleteAsync(ct);
            int profilesDeleted = await dbContext.ProfilesTbs.ExecuteDeleteAsync(ct);

            await transaction.CommitAsync(ct);

            var result = new ResetProfileMigrationDataResult(
                contactsDeleted,
                bankInformationDeleted,
                worksDeleted,
                partnersDeleted,
                profilesDeleted);

            logger.LogWarning(
                "Profile migration data reset completed. Contacts={Contacts}, BankInformation={BankInformation}, Works={Works}, Partners={Partners}, Profiles={Profiles}, Total={Total}",
                result.ContactsDeleted,
                result.BankInformationDeleted,
                result.WorksDeleted,
                result.PartnersDeleted,
                result.ProfilesDeleted,
                result.TotalDeleted);

            return result;
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync(CancellationToken.None);
            logger.LogError(ex, "Profile migration data reset failed; transaction rolled back");
            throw;
        }
    }
}
