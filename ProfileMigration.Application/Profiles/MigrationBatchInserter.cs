using System.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using Oracle.ManagedDataAccess.Client;
using ProfileMigration.DAL.Models;

namespace ProfileMigration.Application.Profiles;

/// <summary>
/// Oracle array-binding insert helper with divide-and-conquer fallback.
/// EF Core metadata remains the source of table, column, and value-converter mappings.
/// </summary>
public static class MigrationBatchInserter
{
    public const int BatchSize = 5_000;
    const int MaxTransientAttempts = 3;
    const int CommandTimeoutSeconds = 120;
    const int MaxRejectedRowsPerBatch = 100;

    public static async Task<int> InsertAsync<TEntity>(
        DbContextOptions<SilaDbContext> dbOptions,
        IReadOnlyList<TEntity> items,
        Func<TEntity, string> skipMessage,
        List<string> log,
        CancellationToken ct,
        Action<string>? progress = null) where TEntity : class
    {
        if (items.Count == 0) return 0;
        var fallbackState = new FallbackState();
        return await InsertWithFallbackAsync(
            dbOptions, items, skipMessage, log, fallbackState, progress, depth: 0, ct);
    }

    static async Task<int> InsertWithFallbackAsync<TEntity>(
        DbContextOptions<SilaDbContext> dbOptions,
        IReadOnlyList<TEntity> items,
        Func<TEntity, string> skipMessage,
        List<string> log,
        FallbackState fallbackState,
        Action<string>? progress,
        int depth,
        CancellationToken ct) where TEntity : class
    {
        try
        {
            await InsertBatchAsync(dbOptions, items, ct);
            return items.Count;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (BatchAbortException)
        {
            throw;
        }
        catch (Exception ex) when (IsTransientConnectionFailure(ex))
        {
            string message = $"[ABORT BATCH] Transient Oracle connection failure after {MaxTransientAttempts} attempts: {RootMessage(ex)}";
            log.Add(message);
            progress?.Invoke(message);
            throw;
        }
        catch (Exception ex) when (IsSystematicBatchFailure(ex))
        {
            string message = $"[ABORT BATCH] Systematic DB error — fix before retry: {RootMessage(ex)}";
            log.Add(message);
            progress?.Invoke(message);
            throw;
        }
        catch (Exception ex)
        {
            if (items.Count == 1)
            {
                log.Add(skipMessage(items[0]) + RootMessage(ex));
                fallbackState.RejectedRows++;
                if (fallbackState.RejectedRows >= MaxRejectedRowsPerBatch)
                {
                    string message =
                        $"[ABORT BATCH] Rejected {fallbackState.RejectedRows} individual rows. " +
                        $"This indicates a widespread data/schema problem. Last error: {RootMessage(ex)}";
                    log.Add(message);
                    progress?.Invoke(message);
                    throw new BatchAbortException(message, ex);
                }
                return 0;
            }

            if (depth == 0)
            {
                string message = $"[BATCH SPLIT] {items.Count} rows failed as a batch: {RootMessage(ex)}";
                log.Add(message);
                progress?.Invoke(message);
            }

            int midpoint = items.Count / 2;
            var left = items.Take(midpoint).ToArray();
            var right = items.Skip(midpoint).ToArray();

            int leftInserted = await InsertWithFallbackAsync(
                dbOptions, left, skipMessage, log, fallbackState, progress, depth + 1, ct);
            int rightInserted = await InsertWithFallbackAsync(
                dbOptions, right, skipMessage, log, fallbackState, progress, depth + 1, ct);
            return leftInserted + rightInserted;
        }
    }

    static async Task InsertBatchAsync<TEntity>(
        DbContextOptions<SilaDbContext> dbOptions,
        IReadOnlyList<TEntity> items,
        CancellationToken ct) where TEntity : class
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                await ExecuteArrayBoundInsertAsync(dbOptions, items, ct);
                return;
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex) when (
                IsTransientConnectionFailure(ex) &&
                attempt < MaxTransientAttempts)
            {
                await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), ct);
            }
        }
    }

    static async Task ExecuteArrayBoundInsertAsync<TEntity>(
        DbContextOptions<SilaDbContext> dbOptions,
        IReadOnlyList<TEntity> items,
        CancellationToken ct) where TEntity : class
    {
        string connectionString;
        string qualifiedTableName;
        IReadOnlyList<ColumnBinding> bindings;

        await using (var metadataContext = new SilaDbContext(dbOptions))
        {
            var entityType = metadataContext.Model.FindEntityType(typeof(TEntity))
                ?? throw new InvalidOperationException(
                    $"EF metadata not found for {typeof(TEntity).Name}.");
            string tableName = entityType.GetTableName()
                ?? throw new InvalidOperationException(
                    $"Table mapping not found for {typeof(TEntity).Name}.");
            string? schema = entityType.GetSchema();
            var storeObject = StoreObjectIdentifier.Table(tableName, schema);

            qualifiedTableName = string.IsNullOrWhiteSpace(schema)
                ? QuoteIdentifier(tableName)
                : $"{QuoteIdentifier(schema)}.{QuoteIdentifier(tableName)}";
            connectionString = metadataContext.Database.GetDbConnection().ConnectionString;
            bindings = BuildBindings(entityType, storeObject, items);
        }

        string columns = string.Join(", ", bindings.Select(x => QuoteIdentifier(x.ColumnName)));
        string parameters = string.Join(", ", bindings.Select((_, index) => $":p{index}"));

        await using var connection = new OracleConnection(connectionString)
        {
            BindByName = true,
        };
        await connection.OpenAsync(ct);
        await using var transaction = await connection.BeginTransactionAsync(ct);
        await using var command = connection.CreateCommand();
        command.BindByName = true;
        command.ArrayBindCount = items.Count;
        command.CommandTimeout = CommandTimeoutSeconds;
        command.Transaction = (OracleTransaction)transaction;
        command.CommandText =
            $"INSERT INTO {qualifiedTableName} ({columns}) VALUES ({parameters})";

        for (int index = 0; index < bindings.Count; index++)
        {
            var binding = bindings[index];
            var parameter = new OracleParameter($"p{index}", binding.OracleDbType)
            {
                Direction = ParameterDirection.Input,
                Value = binding.Values,
            };
            if (binding.Size is int size)
            {
                parameter.Size = size;
                parameter.ArrayBindSize = Enumerable.Repeat(size, items.Count).ToArray();
            }
            command.Parameters.Add(parameter);
        }

        await command.ExecuteNonQueryAsync(ct);
        await transaction.CommitAsync(ct);
    }

    internal static IReadOnlyList<ColumnBinding> BuildBindings<TEntity>(
        IEntityType entityType,
        StoreObjectIdentifier storeObject,
        IReadOnlyList<TEntity> items) where TEntity : class
    {
        var result = new List<ColumnBinding>();

        foreach (var property in entityType.GetProperties())
        {
            if (property.PropertyInfo is null)
                continue;

            string? columnName = property.GetColumnName(storeObject);
            if (string.IsNullOrWhiteSpace(columnName))
                continue;

            var converter = property.GetTypeMapping().Converter;
            var values = new object?[items.Count];
            bool hasValue = false;
            for (int index = 0; index < items.Count; index++)
            {
                object? value = property.PropertyInfo.GetValue(items[index]);
                value = converter?.ConvertToProvider(value) ?? value;
                value = NormalizeProviderValue(value);
                values[index] = value;
                hasValue |= value is not null;
            }

            // Omitting an all-null column preserves its Oracle default.
            if (!hasValue && !property.IsPrimaryKey())
                continue;

            Type providerType = converter?.ProviderClrType ?? property.ClrType;
            providerType = Nullable.GetUnderlyingType(providerType) ?? providerType;
            var oracleType = ResolveOracleType(providerType);
            int? size = oracleType is OracleDbType.Varchar2 or OracleDbType.NVarchar2
                ? property.GetMaxLength() ?? MaxStringLength(values)
                : null;

            result.Add(new ColumnBinding(
                columnName,
                oracleType,
                values.Select(value => value ?? DBNull.Value).ToArray(),
                size));
        }

        if (result.Count == 0)
            throw new InvalidOperationException(
                $"No insertable columns found for {entityType.ClrType.Name}.");

        return result;
    }

    static object? NormalizeProviderValue(object? value) => value switch
    {
        null => null,
        bool flag => flag ? 1 : 0,
        Enum enumValue => Convert.ToInt32(enumValue),
        Guid guid => guid.ToString(),
        _ => value,
    };

    static OracleDbType ResolveOracleType(Type type)
    {
        if (type == typeof(string) || type == typeof(Guid))
            return OracleDbType.Varchar2;
        if (type == typeof(DateTime))
            return OracleDbType.Date;
        if (type == typeof(DateTimeOffset))
            return OracleDbType.TimeStampTZ;
        if (type == typeof(decimal))
            return OracleDbType.Decimal;
        if (type == typeof(double))
            return OracleDbType.Double;
        if (type == typeof(float))
            return OracleDbType.Single;
        if (type == typeof(long) || type == typeof(ulong))
            return OracleDbType.Int64;
        if (type == typeof(int) || type == typeof(uint) ||
            type == typeof(short) || type == typeof(ushort) ||
            type == typeof(byte) || type == typeof(sbyte) ||
            type == typeof(bool))
            return OracleDbType.Int32;
        if (type == typeof(byte[]))
            return OracleDbType.Blob;

        throw new NotSupportedException(
            $"Oracle array binding does not support CLR type {type.FullName}.");
    }

    static int MaxStringLength(IEnumerable<object?> values) =>
        Math.Max(1, values.OfType<string>().Select(x => x.Length).DefaultIfEmpty(1).Max());

    static string QuoteIdentifier(string identifier) =>
        $"\"{identifier.Replace("\"", "\"\"")}\"";

    internal sealed record ColumnBinding(
        string ColumnName,
        OracleDbType OracleDbType,
        object[] Values,
        int? Size);

    static bool IsTransientConnectionFailure(Exception exception)
    {
        for (Exception? current = exception; current is not null; current = current.InnerException)
        {
            if (current is not OracleException oracleException)
                continue;

            if (oracleException.Number is 50000 or 12170 or 12535 or 12537 or 12541 or 12545)
                return true;

            if (oracleException.Message.Contains(
                    "Connection request timed out",
                    StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Errors that affect every row in a batch — splitting would only waste hours.
    /// </summary>
    static bool IsSystematicBatchFailure(Exception exception)
    {
        for (Exception? current = exception; current is not null; current = current.InnerException)
        {
            if (current is not OracleException oracleException)
                continue;

            if (oracleException.Number is
                54 or   // resource busy / NOWAIT lock
                60 or   // deadlock
                1013 or // command cancelled or timed out
                942 or   // table or view does not exist
                904 or   // invalid identifier
                1031 or  // insufficient privileges
                2289 or  // FK — parent key missing
                2291 or  // FK — child record violates parent
                30006 or // DML lock timeout
                4043 or  // object invalid
                4044)   // object does not exist
                return true;
        }

        return false;
    }

    sealed class FallbackState
    {
        public int RejectedRows { get; set; }
    }

    sealed class BatchAbortException(string message, Exception innerException)
        : Exception(message, innerException);

    static string RootMessage(Exception exception)
    {
        while (exception.InnerException is not null)
            exception = exception.InnerException;
        return exception.Message;
    }
}
