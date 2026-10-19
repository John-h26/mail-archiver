using MailArchiver.Models;
using MailArchiver.Services.Providers.Graph;
using MailArchiver.Services.Providers.Imap;
using MailArchiver.Utilities;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;

namespace MailArchiver.Tests.Services;

/// <summary>
/// The Graph request loop needs a live tenant, but the mapping of a Graph message onto the index
/// does not. If that mapping drifts from the IMAP one, a repeated offload into M365 silently
/// duplicates mail, so it is pinned here against archived rows.
/// </summary>
public class GraphTargetMailboxIndexBuilderTests
{
    private static TargetMailboxIndex NewIndex(string timeZoneId = "Etc/UCT")
        => TargetMailboxIndex.CreateEmpty(
            new DateTimeHelper(Options.Create(new TimeZoneOptions { DisplayTimeZoneId = timeZoneId })));

    private static Message GraphMessage(
        string? internetMessageId, string? from, string[] to, string? subject, DateTimeOffset? sent, DateTimeOffset? received = null)
        => new()
        {
            InternetMessageId = internetMessageId,
            From = from == null ? null : new Recipient { EmailAddress = new EmailAddress { Address = from } },
            ToRecipients = to.Select(a => new Recipient { EmailAddress = new EmailAddress { Address = a } }).ToList(),
            Subject = subject,
            SentDateTime = sent,
            ReceivedDateTime = received,
        };

    private static ArchivedEmail Row(string messageId, string from, string to, string subject, DateTime sent)
        => new() { MessageId = messageId, From = from, To = to, Subject = subject, SentDate = sent };

    private static readonly DateTimeOffset SentUtc = new(2026, 2, 3, 10, 11, 12, TimeSpan.Zero);

    [Fact]
    public void AddMessage_MatchesArchivedRowByMessageId()
    {
        var index = NewIndex();
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage("<abc@example.com>", "other@x.com", new[] { "n@x.com" }, "Different", SentUtc.AddYears(-1)));

        // Graph returns the id with brackets; the archive may hold it without.
        Assert.Equal(OffloadMatchKind.MessageId,
            index.Match(Row("abc@example.com", "alice@x.com", "bob@y.com", "Hi", SentUtc.UtcDateTime)));
        Assert.Equal(1, index.IndexedMessages);
    }

    [Fact]
    public void AddMessage_MatchesByFingerprintWhenMessageIdMissing()
    {
        var index = NewIndex();
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage(null, "alice@x.com", new[] { "bob@y.com" }, "Hi", SentUtc.AddSeconds(1)));

        Assert.Equal(OffloadMatchKind.Fingerprint,
            index.Match(Row("", "alice@x.com", "bob@y.com", "Hi", SentUtc.UtcDateTime)));
    }

    [Fact]
    public void AddMessage_SeveralRecipients_JoinLikeTheArchive()
    {
        var index = NewIndex();
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage(null, "alice@x.com", new[] { "bob@y.com", "carol@z.com" }, "Hi", SentUtc));

        Assert.Equal(OffloadMatchKind.Fingerprint,
            index.Match(Row("", "alice@x.com", "bob@y.com, carol@z.com", "Hi", SentUtc.UtcDateTime)));
    }

    [Fact]
    public void AddMessage_ConvertsSendTimeIntoTheDisplayTimeZone()
    {
        // SentDate is stored display-local. In Berlin (+01:00 in February) 10:11:12 UTC is stored
        // as 11:11:12, and the fingerprint must still match.
        var index = NewIndex("Europe/Berlin");
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage(null, "alice@x.com", new[] { "bob@y.com" }, "Hi", SentUtc));

        Assert.Equal(OffloadMatchKind.Fingerprint,
            index.Match(Row("", "alice@x.com", "bob@y.com", "Hi", new DateTime(2026, 2, 3, 11, 11, 12))));
    }

    [Fact]
    public void AddMessage_FallsBackToReceivedTimeWithoutSendTime()
    {
        var index = NewIndex();
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage(null, "alice@x.com", new[] { "bob@y.com" }, "Hi", sent: null, received: SentUtc));

        Assert.Equal(OffloadMatchKind.Fingerprint,
            index.Match(Row("", "alice@x.com", "bob@y.com", "Hi", SentUtc.UtcDateTime)));
    }

    [Fact]
    public void AddMessage_DifferentMessage_DoesNotMatch()
    {
        var index = NewIndex();
        GraphTargetMailboxIndexBuilder.AddMessage(index,
            GraphMessage("<one@x>", "alice@x.com", new[] { "bob@y.com" }, "Hi", SentUtc));

        Assert.Equal(OffloadMatchKind.None,
            index.Match(Row("two@x", "alice@x.com", "bob@y.com", "Hi", SentUtc.UtcDateTime.AddMinutes(5))));
    }

    [Fact]
    public void SelectFields_CoverEveryPropertyTheMappingReads()
    {
        Assert.Equal(
            new[] { "internetMessageId", "from", "toRecipients", "subject", "sentDateTime", "receivedDateTime" },
            GraphTargetMailboxIndexBuilder.SelectFields);
        Assert.InRange(GraphTargetMailboxIndexBuilder.PageSize, 1, 1000);
    }
}
