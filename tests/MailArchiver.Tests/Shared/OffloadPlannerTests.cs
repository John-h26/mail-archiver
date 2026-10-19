using MailArchiver.Models;
using MailArchiver.Services.Shared;

namespace MailArchiver.Tests.Shared;

/// <summary>
/// Grouping is shared by the IMAP and the M365 offload, so both must map the same source folder
/// onto the same target folder.
/// </summary>
public class OffloadPlannerTests
{
    private static OffloadCriteria Criteria(
        IEnumerable<string>? excluded = null, IDictionary<string, string>? rename = null) => new()
    {
        ExcludedSourceFolders = excluded?.ToList() ?? new List<string>(),
        FolderRenameMap = rename == null ? new Dictionary<string, string>() : new Dictionary<string, string>(rename),
    };

    [Fact]
    public void Group_PreservesStructureBelowRoot()
    {
        var groups = OffloadPlanner.GroupByTargetPath(
            new (int, string?)[] { (1, "INBOX"), (2, "INBOX/Projects"), (3, "Sent") },
            "Archive", preserveFolderStructure: true, Criteria(), out var excluded);

        Assert.Equal(0, excluded);
        Assert.Equal(new[] { 1 }, groups["Archive/INBOX"]);
        Assert.Equal(new[] { 2 }, groups["Archive/INBOX/Projects"]);
        Assert.Equal(new[] { 3 }, groups["Archive/Sent"]);
    }

    [Fact]
    public void Group_WithoutStructure_PutsEverythingInRoot()
    {
        var groups = OffloadPlanner.GroupByTargetPath(
            new (int, string?)[] { (1, "INBOX"), (2, "Sent") },
            "INBOX", preserveFolderStructure: false, Criteria(), out _);

        Assert.Single(groups);
        Assert.Equal(new[] { 1, 2 }, groups["INBOX"]);
    }

    [Fact]
    public void Group_CountsExcludedFoldersAndTheirChildren()
    {
        var groups = OffloadPlanner.GroupByTargetPath(
            new (int, string?)[] { (1, "Junk E-Mail"), (2, "Junk E-Mail/Old"), (3, "INBOX") },
            "INBOX", preserveFolderStructure: true, Criteria(excluded: new[] { "Junk E-Mail" }), out var excluded);

        Assert.Equal(2, excluded);
        Assert.Equal(new[] { 3 }, groups["INBOX"]);
    }

    [Fact]
    public void Group_RenamedFoldersCollapseOntoOneTarget()
    {
        var groups = OffloadPlanner.GroupByTargetPath(
            new (int, string?)[] { (1, "Sent Items"), (2, "Sent") },
            "INBOX", preserveFolderStructure: true,
            Criteria(rename: new Dictionary<string, string> { ["Sent Items"] = "Sent" }), out _);

        Assert.Single(groups);
        Assert.Equal(new[] { 1, 2 }, groups["Sent"]);
    }

    [Fact]
    public void Group_MissingSourceFolder_IsTreatedAsInbox()
    {
        var groups = OffloadPlanner.GroupByTargetPath(
            new (int, string?)[] { (1, null) },
            "INBOX", preserveFolderStructure: true, Criteria(), out _);

        Assert.Equal(new[] { 1 }, groups["INBOX"]);
    }
}
