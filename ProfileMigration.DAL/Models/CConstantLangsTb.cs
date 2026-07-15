#nullable disable
using System;

namespace ProfileMigration.DAL.Models;

public partial class CConstantLangsTb
{
    public int ConstantMainId { get; set; }

    public int ConstantId { get; set; }

    public string LangId { get; set; }

    public string ConstantDesc { get; set; }

    public DateTime? CreatedOn { get; set; }

    public int? CreatedBy { get; set; }

    public DateTime? UpdatedOn { get; set; }

    public int? UpdatedBy { get; set; }
}
