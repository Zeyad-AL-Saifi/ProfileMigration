using System.Data;
using Microsoft.Extensions.Options;
using Oracle.ManagedDataAccess.Client;
using ProfileMigration.Application.Options;

namespace ProfileMigration.Application.Data;

public interface IOracleConnectionFactory
{
    Task<IDbConnection> CreateOpenConnectionAsync(CancellationToken ct = default);
    string ConnectionString { get; }
}

public sealed class OracleConnectionFactory(IOptions<MigrationOptions> options) : IOracleConnectionFactory
{
    public string ConnectionString =>
        string.IsNullOrWhiteSpace(options.Value.ConnectionString)
            ? throw new InvalidOperationException("Missing Oracle connection string (ConnectionStrings:OracleDb).")
            : options.Value.ConnectionString;

    public async Task<IDbConnection> CreateOpenConnectionAsync(CancellationToken ct = default)
    {
        var conn = new OracleConnection(ConnectionString)
        {
            // Required so Dapper named params (e.g. :mainId) bind by name, not position.
            BindByName = true,
        };
        await conn.OpenAsync(ct);
        return conn;
    }
}
