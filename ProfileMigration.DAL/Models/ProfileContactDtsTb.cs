#nullable disable
using System;

namespace ProfileMigration.DAL.Models;

public partial class ProfileContactDtsTb
{
    public int ContactId { get; set; }

    public int ProfileId { get; set; }

    public int ContactTypeId { get; set; }

    public int? ContactSubTypeId { get; set; }

    public string ContactInfo { get; set; }

    public int? ContactCountryId { get; set; }

    public DateTime? CreatedOn { get; set; }

    public int? CreatedBy { get; set; }

    public DateTime? UpdatedOn { get; set; }

    public int? UpdatedBy { get; set; }
}
