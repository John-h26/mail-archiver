// Services/Shared/ThreadingHeaders.cs
using System.Text.RegularExpressions;

namespace MailArchiver.Services.Shared
{
    /// <summary>
    /// The conversation headers of an archived message, recovered from its stored
    /// <c>RawHeaders</c>.
    /// <para>
    /// A restored message is rebuilt from database columns, and neither In-Reply-To nor References
    /// has a column of its own. Without them every restored reply arrives as a conversation of its
    /// own: mail clients that thread by these headers (Thunderbird, Apple Mail, most webmail) show
    /// a restored mailbox as a flat list. <c>RawHeaders</c> keeps one "Field: Value" line per
    /// header for both IMAP and Graph archived mail, so the values can be read back from there.
    /// </para>
    /// </summary>
    public sealed class ThreadingHeaders
    {
        /// <summary>
        /// Upper bound on how many References ids are re-emitted. Long threads accumulate hundreds
        /// of ids; RFC 5322 section 3.6.4 allows trimming as long as the first id (the thread root)
        /// is kept, which is what <see cref="Extract"/> does.
        /// </summary>
        internal const int MaxReferences = 50;

        private static readonly Regex MsgIdPattern = new(@"<([^<>\s]+)>", RegexOptions.Compiled);

        public static readonly ThreadingHeaders Empty = new(null, Array.Empty<string>());

        /// <summary>The parent message id, without angle brackets, or null.</summary>
        public string? InReplyTo { get; }

        /// <summary>The thread ancestry, oldest first, without angle brackets.</summary>
        public IReadOnlyList<string> References { get; }

        public bool IsEmpty => InReplyTo == null && References.Count == 0;

        private ThreadingHeaders(string? inReplyTo, IReadOnlyList<string> references)
        {
            InReplyTo = inReplyTo;
            References = references;
        }

        /// <summary>
        /// Reads In-Reply-To and References from stored raw headers. The first occurrence of each
        /// header wins. Values without a bracketed msg-id are ignored rather than guessed at, since
        /// a malformed id in these headers is worse than none: it can attach a message to the
        /// wrong thread.
        /// </summary>
        public static ThreadingHeaders Extract(string? rawHeaders)
        {
            if (string.IsNullOrEmpty(rawHeaders)) return Empty;

            string? inReplyTo = null;
            List<string>? references = null;

            foreach (var rawLine in rawHeaders.Split('\n'))
            {
                var line = rawLine.TrimEnd('\r');

                if (inReplyTo == null && TryGetValue(line, "In-Reply-To:", out var inReplyToValue))
                {
                    inReplyTo = MsgIdPattern.Matches(inReplyToValue)
                        .Select(m => m.Groups[1].Value)
                        .FirstOrDefault();
                }
                else if (references == null && TryGetValue(line, "References:", out var referencesValue))
                {
                    var ids = MsgIdPattern.Matches(referencesValue)
                        .Select(m => m.Groups[1].Value)
                        .Distinct(StringComparer.Ordinal)
                        .ToList();
                    if (ids.Count > 0) references = Trim(ids);
                }

                if (inReplyTo != null && references != null) break;
            }

            if (inReplyTo == null && references == null) return Empty;
            return new ThreadingHeaders(inReplyTo, (IReadOnlyList<string>?)references ?? Array.Empty<string>());
        }

        /// <summary>Keeps the thread root and the most recent ancestors.</summary>
        private static List<string> Trim(List<string> ids)
        {
            if (ids.Count <= MaxReferences) return ids;
            var trimmed = new List<string>(MaxReferences) { ids[0] };
            trimmed.AddRange(ids.Skip(ids.Count - (MaxReferences - 1)));
            return trimmed;
        }

        private static bool TryGetValue(string line, string prefix, out string value)
        {
            if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                value = line.Substring(prefix.Length);
                return true;
            }

            value = string.Empty;
            return false;
        }
    }
}
