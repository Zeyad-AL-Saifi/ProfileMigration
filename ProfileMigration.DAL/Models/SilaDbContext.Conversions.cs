using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace ProfileMigration.DAL.Models;

// ════════════════════════════════════════════════════════════════════════════════
//  Hand-written partial — survives EF Power Tools re-scaffolding.
//
//  The scaffolder maps every Oracle NUMBER(1) column to a C# bool, but a few of
//  these columns are 2-VALUE DOMAIN CODES (1/2), not true/false flags:
//
//      GENDER_ID        1 = Male      2 = Female
//      CUST_TYPE_ID     1 = Firm      2 = Individual
//      ENTRY_SOURCE_ID  1 = Company   2 = Self
//
//  Each has exactly two valid values in the source data, so a bool can carry them
//  faithfully as long as we control how the bool is stored: false -> 1, true -> 2.
//  These converters do exactly that, keeping the generated model classes and the
//  generated SilaDbContext.cs untouched.  The mapping is centralised here so it is
//  easy to reconfigure or extend.
// ════════════════════════════════════════════════════════════════════════════════
public partial class SilaDbContext
{
    // false -> low code (1), true -> high code (2).  Edit here to reconfigure.
    private const int CodeForFalse = 1;
    private const int CodeForTrue  = 2;

    // bool? <-> NUMBER(1) where 1 and 2 are the two domain codes.
    private static readonly ValueConverter<bool, int> CodeBoolConverter =
        new(b => b ? CodeForTrue : CodeForFalse,
            n => n == CodeForTrue);

    private static readonly ValueConverter<bool?, int?> NullableCodeBoolConverter =
        new(b => b.HasValue ? (b.Value ? CodeForTrue : CodeForFalse) : (int?)null,
            n => n.HasValue ? n.Value == CodeForTrue : (bool?)null);

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<ProfilesTb>(entity =>
        {
            // CUST_TYPE_ID is non-nullable bool in the model.
            entity.Property(e => e.CustTypeId).HasConversion(CodeBoolConverter);
            // GENDER_ID and ENTRY_SOURCE_ID are nullable bool in the model.
            entity.Property(e => e.GenderId).HasConversion(NullableCodeBoolConverter);
            entity.Property(e => e.EntrySourceId).HasConversion(NullableCodeBoolConverter);
        });

        modelBuilder.Entity<ProfilesPartnersTb>(entity =>
        {
            // IS_BANK_BORROWER is documented as نعم=2/لا=1 (a 1/2 domain code, not a
            // 0/1 flag) and is non-nullable bool in the model.
            entity.Property(e => e.IsBankBorrower).HasConversion(CodeBoolConverter);
        });
    }
}
