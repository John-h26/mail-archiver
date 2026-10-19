using MailArchiver.Services.Shared;

namespace MailArchiver.Tests.Shared;

public class ThreadingHeadersTests
{
    [Fact]
    public void Extract_ReadsInReplyToAndReferences()
    {
        var raw = string.Join("\r\n",
            "Received: from mx by next; Sun, 12 Jul 2026 10:00:00 +0000",
            "Message-ID: <child@example.com>",
            "In-Reply-To: <parent@example.com>",
            "References: <root@example.com> <parent@example.com>",
            "Subject: Re: hello");

        var result = ThreadingHeaders.Extract(raw);

        Assert.Equal("parent@example.com", result.InReplyTo);
        Assert.Equal(new[] { "root@example.com", "parent@example.com" }, result.References);
        Assert.False(result.IsEmpty);
    }

    [Fact]
    public void Extract_IsCaseInsensitiveOnFieldNames()
    {
        var result = ThreadingHeaders.Extract("in-reply-to: <a@x>\nREFERENCES: <b@x>");

        Assert.Equal("a@x", result.InReplyTo);
        Assert.Equal(new[] { "b@x" }, result.References);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Subject: no threading here")]
    public void Extract_ReturnsEmptyWhenAbsent(string? raw)
    {
        Assert.True(ThreadingHeaders.Extract(raw).IsEmpty);
    }

    [Fact]
    public void Extract_IgnoresValuesWithoutBracketedMsgId()
    {
        // A free-text In-Reply-To (some old clients wrote "Your message of ...") must not be
        // turned into an id: a wrong id can attach the message to an unrelated thread.
        var result = ThreadingHeaders.Extract("In-Reply-To: your message of yesterday\nReferences: garbage");

        Assert.True(result.IsEmpty);
    }

    [Fact]
    public void Extract_FirstOccurrenceWins()
    {
        var result = ThreadingHeaders.Extract("In-Reply-To: <first@x>\nIn-Reply-To: <second@x>");

        Assert.Equal("first@x", result.InReplyTo);
    }

    [Fact]
    public void Extract_RemovesDuplicateReferences()
    {
        var result = ThreadingHeaders.Extract("References: <a@x> <b@x> <a@x>");

        Assert.Equal(new[] { "a@x", "b@x" }, result.References);
    }

    [Fact]
    public void Extract_TrimsLongReferencesKeepingRootAndMostRecent()
    {
        var ids = Enumerable.Range(1, 120).Select(i => $"<id{i}@x>");
        var result = ThreadingHeaders.Extract("References: " + string.Join(" ", ids));

        Assert.Equal(ThreadingHeaders.MaxReferences, result.References.Count);
        Assert.Equal("id1@x", result.References[0]);
        Assert.Equal("id120@x", result.References[^1]);
        Assert.Equal("id72@x", result.References[1]);
    }
}
