#nullable disable
using System;

namespace ProfileMigration.DAL.Models;

public partial class CConstantsTb
{
    public int ConstantMainId { get; set; }

    public int ConstantId { get; set; }

    public string ConstantDesc { get; set; }

    public int? ParentConstantId { get; set; }

    public string PmaCode { get; set; }

    public string ConstantCondition { get; set; }

    public string ConstantCode1 { get; set; }

    public string ConstantCode2 { get; set; }

    public string ConstantCode3 { get; set; }

    public bool IsUpdatable { get; set; }

    public bool IsHidden { get; set; }

    public DateTime? CreatedOn { get; set; }

    public int? CreatedBy { get; set; }

    public DateTime? UpdatedOn { get; set; }

    public int? UpdatedBy { get; set; }
}
