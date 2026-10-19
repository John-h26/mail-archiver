using MailArchiver.Models;
using MailArchiver.Services.Providers.Imap;
using MailArchiver.Utilities;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace MailArchiver.Tests.Services;

/// <summary>
/// A restored reply must carry its In-Reply-To and References headers again, otherwise the target
/// client shows it as a conversation of its own.
/// </summary>
public class ImapMailRestorerThreadingTests
{
    private static ImapMailRestorer CreateRestorer() => new(
        context: null!,
        NullLogger<ImapMailRestorer>.Instance,
        connectionFactory: null!,
        new DateTimeHelper(Options.Create(new TimeZoneOptions { DisplayTimeZoneId = "UTC" })),
        Options.Create(new BatchOperationOptions()),
        Options.Create(new OffloadOptions()),
        folderService: null!);

    private static ArchivedEmail CreateEmail(string? rawHeaders) => new()
    {
        MessageId = "child@example.com",
        Subject = "Re: hello",
        From = "sender@example.com",
        To = "recipient@example.com",
        SentDate = new DateTime(2026, 7, 12, 10, 0, 0, DateTimeKind.Unspecified),
        RawHeaders = rawHeaders,
    };

    [Fact]
    public async Task CreateMimeMessage_RestoresThreadingHeadersFromRawHeaders()
    {
        var email = CreateEmail(
            "In-Reply-To: <parent@example.com>\r\nReferences: <root@example.com> <parent@example.com>\r\n");

        var message = await CreateRestorer().CreateMimeMessageFromArchivedEmailAsync(email, "target");

        Assert.Equal("parent@example.com", message.InReplyTo);
        Assert.Equal(new[] { "root@example.com", "parent@example.com" }, message.References.ToArray());
        Assert.Equal("child@example.com", message.MessageId);
    }

    [Fact]
    public async Task CreateMimeMessage_WithoutRawHeaders_EmitsNoThreadingHeaders()
    {
        var message = await CreateRestorer().CreateMimeMessageFromArchivedEmailAsync(CreateEmail(null), "target");

        Assert.Null(message.InReplyTo);
        Assert.Empty(message.References);
    }
}
