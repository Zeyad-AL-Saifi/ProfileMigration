using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;
using ProfileMigration.Application.Profiles.Commands;
using ProfileMigration.DAL.Models;
using Xunit;

namespace ProfileMigration.Tests;

public sealed class ResetProfileMigrationDataCommandHandlerTests
{
    [Fact]
    public async Task HandleAsync_WithoutConfirmation_RejectsCommandBeforeDatabaseAccess()
    {
        var options = new DbContextOptionsBuilder<SilaDbContext>()
            .UseOracle("User Id=test;Password=test;Data Source=localhost:1521/test")
            .Options;
        await using var dbContext = new SilaDbContext(options);
        var handler = new ResetProfileMigrationDataCommandHandler(
            dbContext,
            NullLogger<ResetProfileMigrationDataCommandHandler>.Instance);

        await Assert.ThrowsAsync<ArgumentException>(() =>
            handler.HandleAsync(new ResetProfileMigrationDataCommand(false)));
    }

    [Fact]
    public void Result_TotalDeleted_SumsAllDeletedRows()
    {
        var result = new ResetProfileMigrationDataResult(1, 2, 3, 4, 5);

        Assert.Equal(15, result.TotalDeleted);
    }
}
