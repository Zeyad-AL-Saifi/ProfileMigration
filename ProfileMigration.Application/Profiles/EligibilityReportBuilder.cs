using ProfileMigration.Application.Dtos;

namespace ProfileMigration.Application.Profiles;

public static class EligibilityReportBuilder
{
    public static void Apply(
        MigrationReportDto report,
        ClientEligibilityClassifier.Result eligibility)
    {
        report.Stats["total_input"] = eligibility.TotalInput;
        report.Stats["eligible_count"] = eligibility.EligibleCount;
        report.Stats["skipped_missing_id_card"] = eligibility.SkippedMissingIdCard;
        report.Stats["skipped_internal_duplicates"] = eligibility.SkippedInternalDuplicates;
        report.Stats["skipped_cross_company_matches"] = eligibility.SkippedCrossCompanyMatches;
        report.Stats["skipped_duplicate_id_nums"] = eligibility.SkippedDuplicateIdNums;

        Add(report, ClientEligibilityClassifier.ReasonMissingIdCard,
            "Clients without ID_CARD_TYPE and ID_CARD_NO are excluded.",
            eligibility.SkippedMissingIdCard, eligibility);
        Add(report, ClientEligibilityClassifier.ReasonInternalDuplicateAsala,
            "ASALA clients sharing the same card_key are all excluded.",
            eligibility.Skipped.Count(x => x.Value == ClientEligibilityClassifier.ReasonInternalDuplicateAsala),
            eligibility);
        Add(report, ClientEligibilityClassifier.ReasonInternalDuplicateAcad,
            "ACAD clients sharing the same card_key are all excluded.",
            eligibility.Skipped.Count(x => x.Value == ClientEligibilityClassifier.ReasonInternalDuplicateAcad),
            eligibility);
        Add(report, ClientEligibilityClassifier.ReasonCrossCompanyMatch,
            "A card_key present in ASALA and ACAD excludes all clients on both sides.",
            eligibility.SkippedCrossCompanyMatches, eligibility);
        Add(report, ClientEligibilityClassifier.ReasonDuplicateIdNum,
            "The same ID_NUM and ID_TYPE belongs to multiple clients; all are excluded and never merged.",
            eligibility.SkippedDuplicateIdNums, eligibility);
    }

    static void Add(
        MigrationReportDto report,
        string reason,
        string message,
        int count,
        ClientEligibilityClassifier.Result eligibility) =>
        ReportBuilder.AddIssue(
            report,
            "Warning",
            reason,
            message,
            count,
            eligibility.SampleSkipped(reason));
}
