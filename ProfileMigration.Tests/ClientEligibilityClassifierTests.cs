using ProfileMigration.Application.Profiles;
using Xunit;
using static ProfileMigration.Application.Profiles.ClientEligibilityClassifier;

namespace ProfileMigration.Tests;

public class ClientEligibilityClassifierTests
{
    [Fact]
    public void BuildCardKey_RequiresBothTypeAndNumber()
    {
        Assert.Null(BuildCardKey(null, "123"));
        Assert.Null(BuildCardKey(1, null));
        Assert.Null(BuildCardKey(1, "  "));
        Assert.Equal("1|123", BuildCardKey(1, " 123 "));
    }

    [Fact]
    public void MissingIdCard_IsNeverEligible()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 10, null),
            new ClientIdentityRow("ACAD", 20, "1|AAA"),
        ]);

        Assert.Equal(2, result.TotalInput);
        Assert.False(result.IsEligible("ASALA", 10));
        Assert.True(result.IsEligible("ACAD", 20));
        Assert.Equal(1, result.SkippedMissingIdCard);
        Assert.Equal(ReasonMissingIdCard, result.Skipped[("ASALA", 10)]);
    }

    [Fact]
    public void InternalDuplicate_Asala_ExcludesAllClientIdsInGroup()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 1, "2|111"),
            new ClientIdentityRow("ASALA", 2, "2|111"),
            new ClientIdentityRow("ASALA", 3, "2|222"), // unique — eligible
        ]);

        Assert.False(result.IsEligible("ASALA", 1));
        Assert.False(result.IsEligible("ASALA", 2));
        Assert.True(result.IsEligible("ASALA", 3));
        Assert.Equal(2, result.SkippedInternalDuplicates);
        Assert.Equal(ReasonInternalDuplicateAsala, result.Skipped[("ASALA", 1)]);
        Assert.Equal(ReasonInternalDuplicateAsala, result.Skipped[("ASALA", 2)]);
    }

    [Fact]
    public void InternalDuplicate_Acad_ExcludesAllClientIdsInGroup()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ACAD", 1, "1|X"),
            new ClientIdentityRow("ACAD", 2, "1|X"),
        ]);

        Assert.False(result.IsEligible("ACAD", 1));
        Assert.False(result.IsEligible("ACAD", 2));
        Assert.Equal(ReasonInternalDuplicateAcad, result.Skipped[("ACAD", 1)]);
    }

    [Fact]
    public void CrossCompanyMatch_ExcludesBothSides_EvenAtPerfectCardMatch()
    {
        // Same card_key in ASALA and ACAD → exclude ALL involved CLIENT_IDs
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 100, "1|999"),
            new ClientIdentityRow("ACAD", 200, "1|999"),
            new ClientIdentityRow("ASALA", 101, "1|888"), // clean unique
        ]);

        Assert.False(result.IsEligible("ASALA", 100));
        Assert.False(result.IsEligible("ACAD", 200));
        Assert.True(result.IsEligible("ASALA", 101));
        Assert.Equal(2, result.SkippedCrossCompanyMatches);
        Assert.Equal(ReasonCrossCompanyMatch, result.Skipped[("ASALA", 100)]);
        Assert.Equal(ReasonCrossCompanyMatch, result.Skipped[("ACAD", 200)]);
    }

    [Fact]
    public void CrossCompany_ExcludesAllClientsSharingKey_WhenMultiplePerSide()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 1, "5|PASS"),
            new ClientIdentityRow("ASALA", 2, "5|PASS"),
            new ClientIdentityRow("ACAD", 3, "5|PASS"),
        ]);

        Assert.False(result.IsEligible("ASALA", 1));
        Assert.False(result.IsEligible("ASALA", 2));
        Assert.False(result.IsEligible("ACAD", 3));
        // Cross takes precedence over internal for ASALA 1/2
        Assert.Equal(ReasonCrossCompanyMatch, result.Skipped[("ASALA", 1)]);
        Assert.Equal(ReasonCrossCompanyMatch, result.Skipped[("ASALA", 2)]);
        Assert.Equal(ReasonCrossCompanyMatch, result.Skipped[("ACAD", 3)]);
        Assert.Equal(3, result.SkippedCrossCompanyMatches);
        Assert.Equal(0, result.SkippedInternalDuplicates);
    }

    [Fact]
    public void HappyPath_CleanUniqueClientsRemainEligible()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 1, "1|A"),
            new ClientIdentityRow("ACAD", 2, "1|B"),
            new ClientIdentityRow("asala", 3, "2|C"), // company normalized
        ]);

        Assert.Equal(3, result.TotalInput);
        Assert.Equal(3, result.EligibleCount);
        Assert.Equal(0, result.SkippedMissingIdCard);
        Assert.Equal(0, result.SkippedInternalDuplicates);
        Assert.Equal(0, result.SkippedCrossCompanyMatches);
        Assert.True(result.IsEligible("ASALA", 1));
        Assert.True(result.IsEligible("ACAD", 2));
        Assert.True(result.IsEligible("ASALA", 3));
    }

    [Fact]
    public void DuplicateIdNum_WithDifferentCardKeys_ExcludesEveryClient()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 1, "1|CARD-A", "SAME-ID", 1),
            new ClientIdentityRow("ASALA", 2, "2|CARD-B", "SAME-ID", 1),
            new ClientIdentityRow("ACAD", 3, "1|UNIQUE", "OTHER-ID", 1),
        ]);

        Assert.False(result.IsEligible("ASALA", 1));
        Assert.False(result.IsEligible("ASALA", 2));
        Assert.True(result.IsEligible("ACAD", 3));
        Assert.Equal(2, result.SkippedDuplicateIdNums);
        Assert.Equal(ReasonDuplicateIdNum, result.Skipped[("ASALA", 1)]);
        Assert.Equal(ReasonDuplicateIdNum, result.Skipped[("ASALA", 2)]);
    }

    [Fact]
    public void RepeatedSourceRows_ForSameClient_AreAllExcluded()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 7, "1|REPEATED", "REPEATED", 1),
            new ClientIdentityRow("ASALA", 7, "1|REPEATED", "REPEATED", 1),
        ]);

        Assert.Equal(2, result.TotalInput);
        Assert.Equal(0, result.EligibleCount);
        Assert.Equal(2, result.SkippedInternalDuplicates);
        Assert.Equal(ReasonInternalDuplicateAsala, result.Skipped[("ASALA", 7)]);
    }

    [Fact]
    public void SameIdNumber_WithDifferentTypes_RemainsDistinct()
    {
        var result = Classify(
        [
            new ClientIdentityRow("ASALA", 1, "1|SAME", "SAME", 1),
            new ClientIdentityRow("ASALA", 2, "2|SAME", "SAME", 2),
        ]);

        Assert.Equal(2, result.EligibleCount);
        Assert.Equal(0, result.SkippedDuplicateIdNums);
        Assert.NotEqual(
            BuildDbIdentityKey("SAME", 1),
            BuildDbIdentityKey("SAME", 2));
    }

    [Fact]
    public void IgnoresUnknownCompanies()
    {
        var result = Classify(
        [
            new ClientIdentityRow("OTHER", 1, "1|Z"),
            new ClientIdentityRow("ASALA", 2, "1|Y"),
        ]);

        Assert.Equal(1, result.TotalInput);
        Assert.True(result.IsEligible("ASALA", 2));
        Assert.False(result.IsEligible("OTHER", 1));
    }
}
