using MailArchiver.Models;
using MailArchiver.Services.Providers.Graph;
using MailArchiver.Utilities;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace MailArchiver.Tests.Services;

/// <summary>
/// The decisions in <see cref="GraphMailRestorer"/> that do not need a Graph endpoint: which
/// attachments need an upload session, the read flag, Content-ID normalisation and the delivery
/// time stamped on a restored message.
/// </summary>
public class GraphMailRestorerStaticTests
{
    [Theory]
    [InlineData(0, false)]
    [InlineData(1_024, false)]
    [InlineData(GraphMailRestorer.SimpleAttachmentMaxBytes, false)]
    [InlineData(GraphMailRestorer.SimpleAttachmentMaxBytes + 1, true)]
    [InlineData(25_000_000, true)]
    public void RequiresUploadSession_SwitchesAboveSimpleLimit(long size, bool expected)
    {
        Assert.Equal(expected, GraphMailRestorer.RequiresUploadSession(size));
    }

    [Fact]
    public void UploadSliceSize_IsMultipleOf320KiBAndBelow4MB()
    {
        Assert.Equal(0, GraphMailRestorer.UploadSliceSize % (320 * 1024));
        Assert.True(GraphMailRestorer.UploadSliceSize < 4 * 1024 * 1024);
    }

    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public void BuildMessageFlags_FollowsReadState(bool markAsRead, string expected)
    {
        Assert.Equal(expected, GraphMailRestorer.BuildMessageFlags(markAsRead));
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("  ", null)]
    [InlineData("<img1@mail>", "img1@mail")]
    [InlineData("img1@mail", "img1@mail")]
    [InlineData(" <img1@mail> ", "img1@mail")]
    public void NormalizeContentId_StripsBrackets(string? input, string? expected)
    {
        Assert.Equal(expected, GraphMailRestorer.NormalizeContentId(input));
    }

    [Fact]
    public void ResolveDeliveryTime_PrefersTopmostReceivedHeader()
    {
        var restorer = CreateRestorer("Europe/Berlin");
        var email = new ArchivedEmail
        {
            SentDate = new DateTime(2026, 7, 12, 11, 0, 0, DateTimeKind.Unspecified),
            RawHeaders = "Received: from mx2 by next; Sun, 12 Jul 2026 10:00:00 +0000\r\n" +
                         "Received: from mx1 by next; Sun, 12 Jul 2026 09:55:00 +0000\r\n",
        };

        Assert.Equal("2026-07-12T10:00:00.000Z", restorer.ResolveDeliveryTimeUtcIso8601(email));
    }

    [Fact]
    public void ResolveDeliveryTime_FallsBackToSentDateConvertedFromDisplayZone()
    {
        // SentDate is stored display-local: 11:00 in Berlin (+02:00 in July) is 09:00 UTC.
        var restorer = CreateRestorer("Europe/Berlin");
        var email = new ArchivedEmail
        {
            SentDate = new DateTime(2026, 7, 12, 11, 0, 0, DateTimeKind.Unspecified),
        };

        Assert.Equal("2026-07-12T09:00:00.000Z", restorer.ResolveDeliveryTimeUtcIso8601(email));
    }

    private static GraphMailRestorer CreateRestorer(string timeZoneId) => new(
        authFactory: null!,
        folderService: null!,
        context: null!,
        NullLogger<GraphMailRestorer>.Instance,
        new DateTimeHelper(Options.Create(new TimeZoneOptions { DisplayTimeZoneId = timeZoneId })),
        Options.Create(new OffloadOptions()));
}
