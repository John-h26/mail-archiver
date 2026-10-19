// Services/Providers/Graph/GraphTargetMailboxIndexBuilder.cs
using MailArchiver.Services.Providers.Imap;
using MailArchiver.Utilities;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MailArchiver.Services.Providers.Graph
{
    /// <summary>
    /// Fills a <see cref="TargetMailboxIndex"/> from an M365 mailbox through Microsoft Graph, so an
    /// offload into M365 is as repeatable as one into IMAP: a second run appends nothing the
    /// mailbox already holds.
    /// <para>
    /// Same scope rule as the IMAP builder: every folder of the target mailbox is read, not only
    /// the folders the offload writes to, so the index survives users reorganising their mail
    /// between two runs and a run repeated with a different rename map.
    /// </para>
    /// <para>
    /// Only the six properties both match criteria need are selected, in pages of 999 (the Graph
    /// maximum for messages), so indexing a mailbox of 100,000 messages is about a hundred
    /// requests and never transfers a body.
    /// </para>
    /// </summary>
    internal static class GraphTargetMailboxIndexBuilder
    {
        internal const int PageSize = 999;

        internal static readonly string[] SelectFields =
        {
            "internetMessageId", "from", "toRecipients", "subject", "sentDateTime", "receivedDateTime",
        };

        public static async Task<TargetMailboxIndex> BuildAsync(
            GraphServiceClient graphClient,
            string userPrincipalName,
            IReadOnlyCollection<MailFolder> folders,
            DateTimeHelper dateTimeHelper,
            ILogger logger,
            int prefetchMaxMessages,
            CancellationToken cancellationToken = default)
        {
            var index = TargetMailboxIndex.CreateEmpty(dateTimeHelper);

            foreach (var folder in folders)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrEmpty(folder.Id)) continue;

                if (index.IndexedMessages >= prefetchMaxMessages)
                {
                    index.MarkScopeReduced();
                    logger.LogWarning(
                        "Target mailbox index reached the configured limit of {Limit} messages after {Folders} folders. " +
                        "Duplicate detection no longer covers the whole mailbox, so a repeated offload may append mail " +
                        "that is already present in folders which were not indexed. Raise Offload:PrefetchMaxMessages to avoid this.",
                        prefetchMaxMessages, index.IndexedFolders);
                    break;
                }

                try
                {
                    await IndexFolderAsync(graphClient, userPrincipalName, folder, index, prefetchMaxMessages, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // One unreadable folder must not abort the job, but it narrows the scope.
                    index.MarkScopeReduced();
                    logger.LogWarning(ex,
                        "Could not index target folder '{Folder}'. Mail already present there will not be recognised, " +
                        "so a repeated offload may duplicate it.", folder.DisplayName);
                }
            }

            logger.LogInformation(
                "M365 target mailbox index built for {Mailbox}: {Messages} messages across {Folders} folders, " +
                "{MessageIdKeys} Message-ID keys, {Fingerprints} fingerprints, scope reduced: {Reduced}",
                userPrincipalName, index.IndexedMessages, index.IndexedFolders, index.MessageIdKeyCount,
                index.FingerprintCount, index.ScopeReduced);

            return index;
        }

        private static async Task IndexFolderAsync(
            GraphServiceClient graphClient,
            string userPrincipalName,
            MailFolder folder,
            TargetMailboxIndex index,
            int prefetchMaxMessages,
            CancellationToken cancellationToken)
        {
            index.CountFolder();
            if (folder.TotalItemCount == 0) return;

            var messages = graphClient.Users[userPrincipalName].MailFolders[folder.Id].Messages;
            var response = await messages.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = SelectFields;
                requestConfiguration.QueryParameters.Top = PageSize;
            }, cancellationToken);

            while (response?.Value != null)
            {
                foreach (var message in response.Value)
                {
                    if (index.IndexedMessages >= prefetchMaxMessages)
                    {
                        index.MarkScopeReduced();
                        return;
                    }

                    AddMessage(index, message);
                }

                if (string.IsNullOrEmpty(response.OdataNextLink)) break;

                cancellationToken.ThrowIfCancellationRequested();
                response = await messages.WithUrl(response.OdataNextLink).GetAsync(cancellationToken: cancellationToken);
            }
        }

        /// <summary>
        /// Feeds one Graph message into the index, mapping its properties onto the same inputs an
        /// IMAP ENVELOPE provides.
        /// </summary>
        internal static void AddMessage(TargetMailboxIndex index, Message message)
        {
            var from = message.From?.EmailAddress?.Address;

            index.AddEnvelope(
                message.InternetMessageId,
                string.IsNullOrEmpty(from) ? null : new[] { from },
                message.ToRecipients?
                    .Select(r => r.EmailAddress?.Address)
                    .Where(a => !string.IsNullOrEmpty(a))
                    .Select(a => a!),
                message.Subject,
                message.SentDateTime,
                message.ReceivedDateTime);
        }
    }
}
