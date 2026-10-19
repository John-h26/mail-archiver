using MailArchiver.Data;
using MailArchiver.Models;
using MailArchiver.Services.Shared;
using MailArchiver.Utilities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession;

namespace MailArchiver.Services.Providers.Graph
{
    /// <summary>
    /// Restores archived emails back to a Microsoft Graph (M365) mailbox, including
    /// attachment restoration, inline image processing, and folder hierarchy resolution.
    /// </summary>
    public class GraphMailRestorer
    {
        /// <summary>
        /// Largest attachment sent inline with a plain POST. Graph rejects requests above 4 MB and
        /// the payload is base64 encoded (4/3 overhead), so Microsoft documents 3 MB as the limit
        /// for the simple path. Anything larger has to go through an upload session.
        /// </summary>
        internal const long SimpleAttachmentMaxBytes = 3_000_000;

        /// <summary>
        /// Slice size for upload sessions. Graph requires a multiple of 320 KiB and caps a single
        /// request at 4 MB; ten slices of 320 KiB (3.125 MiB) stays safely below that.
        /// </summary>
        internal const int UploadSliceSize = 320 * 1024 * 10;

        /// <summary>PidTagMessageFlags bit MSGFLAG_READ.</summary>
        private const int MsgFlagRead = 0x1;

        private readonly GraphAuthClientFactory _authFactory;
        private readonly IGraphFolderService _folderService;
        private readonly MailArchiverDbContext _context;
        private readonly ILogger<GraphMailRestorer> _logger;
        private readonly DateTimeHelper _dateTimeHelper;
        private readonly OffloadOptions _offloadOptions;

        public GraphMailRestorer(
            GraphAuthClientFactory authFactory,
            IGraphFolderService folderService,
            MailArchiverDbContext context,
            ILogger<GraphMailRestorer> logger,
            DateTimeHelper dateTimeHelper,
            IOptions<OffloadOptions> offloadOptions)
        {
            _authFactory = authFactory;
            _folderService = folderService;
            _context = context;
            _logger = logger;
            _dateTimeHelper = dateTimeHelper;
            _offloadOptions = offloadOptions.Value;
        }

        /// <summary>
        /// Whether an attachment of the given size must be uploaded through an upload session
        /// instead of a single POST.
        /// </summary>
        internal static bool RequiresUploadSession(long sizeInBytes) => sizeInBytes > SimpleAttachmentMaxBytes;

        /// <summary>
        /// The PidTagMessageFlags value for a restored, non-draft message. MSGFLAG_UNSENT is never
        /// set, so Exchange treats the item as received mail; MSGFLAG_READ follows the requested
        /// read state so it cannot contradict <see cref="Message.IsRead"/>.
        /// </summary>
        internal static string BuildMessageFlags(bool markAsRead) => (markAsRead ? MsgFlagRead : 0).ToString(System.Globalization.CultureInfo.InvariantCulture);

        /// <summary>
        /// The delivery time stamped on a restored message: the topmost Received header when the
        /// archive holds one, otherwise the send time. Same rule as the IMAP INTERNALDATE, so a
        /// mailbox restored to M365 and one restored to IMAP sort identically by arrival.
        /// </summary>
        internal string ResolveDeliveryTimeUtcIso8601(ArchivedEmail email)
            => InternalDateResolver.Resolve(email, _dateTimeHelper)
                .UtcDateTime
                .ToString("yyyy-MM-ddTHH:mm:ss.fffZ", System.Globalization.CultureInfo.InvariantCulture);

        /// <summary>
        /// Restores a single archived email to a specific folder in the target M365 mailbox.
        /// </summary>
        public async Task<bool> RestoreEmailToFolderAsync(
            ArchivedEmail email,
            MailAccount targetAccount,
            string folderName,
            bool preserveFolderStructure = false)
        {
            _logger.LogInformation("RestoreEmailToFolderAsync via Graph API called for email {EmailId} to folder {FolderName}, preserveFolderStructure={Preserve}",
                email.Id, folderName, preserveFolderStructure);

            var graphClient = _authFactory.CreateGraphClient(targetAccount);
            var folders = await _folderService.GetAllMailFoldersAsync(graphClient, targetAccount.EmailAddress);

            return await RestoreEmailToFolderInternalAsync(email, targetAccount, folderName, preserveFolderStructure, graphClient, folders);
        }

        /// <summary>
        /// Internal implementation that accepts pre-fetched folders and Graph client to avoid
        /// redundant API calls when restoring many emails in a batch.
        /// </summary>
        private async Task<bool> RestoreEmailToFolderInternalAsync(
            ArchivedEmail email,
            MailAccount targetAccount,
            string folderName,
            bool preserveFolderStructure,
            GraphServiceClient graphClient,
            List<MailFolder> folders)
        {
            try
            {
                // Determine the desired full target path
                string targetFolderPath;
                if (preserveFolderStructure)
                {
                    var originalFolderName = email.FolderName ?? "INBOX";
                    targetFolderPath = string.IsNullOrWhiteSpace(folderName)
                        ? originalFolderName
                        : $"{folderName.Trim().TrimEnd('/', '\\')}/{originalFolderName.Trim().TrimStart('/', '\\')}";

                    _logger.LogInformation("Preserving folder structure: original folder '{OriginalFolder}' -> target '{TargetPath}'",
                        originalFolderName, targetFolderPath);
                }
                else
                {
                    targetFolderPath = folderName;
                }

                // Resolve / create the folder hierarchy
                MailFolder? targetFolder = await _folderService.EnsureFolderPathAsync(
                    graphClient,
                    targetAccount.EmailAddress,
                    targetFolderPath,
                    folders,
                    createIfMissing: preserveFolderStructure);

                if (targetFolder == null && preserveFolderStructure)
                {
                    _logger.LogWarning("Could not create folder structure for {TargetPath}, falling back to base folder {FolderName}",
                        targetFolderPath, folderName);
                    targetFolder = await _folderService.EnsureFolderPathAsync(
                        graphClient,
                        targetAccount.EmailAddress,
                        folderName,
                        folders,
                        createIfMissing: true);
                }

                if (targetFolder == null)
                {
                    targetFolder = await _folderService.GetWellKnownInboxAsync(graphClient, targetAccount.EmailAddress, folders);
                    if (targetFolder == null)
                    {
                        _logger.LogError("Could not find target folder {FolderName} or Inbox for account {AccountName}",
                            targetFolderPath, targetAccount.Name);
                        return false;
                    }
                    _logger.LogWarning("Target folder {FolderName} not found, using Inbox instead", targetFolderPath);
                }

                // Process the HTML body to ensure inline images are properly referenced
                var processedHtmlBody = email.HtmlBody;
                if (!string.IsNullOrEmpty(email.HtmlBody) && email.Attachments != null && email.Attachments.Any(a => !string.IsNullOrEmpty(a.ContentId)))
                {
                    processedHtmlBody = MailContentHelper.ProcessHtmlBodyForInlineImages(email.HtmlBody, email.Attachments);
                }

                // Create the message to restore.
                // Normalize the stored Message-ID (legacy Graph rows may carry surrounding
                // angle brackets) and re-emit it in the bracketed RFC 2822 form that Exchange
                // itself generates, so a restored message roundtrips to the same canonical id.
                var restorableMessageId = MailContentHelper.ToRestorableMessageId(email.MessageId);
                var internetMessageId = string.IsNullOrEmpty(restorableMessageId) ? null : "<" + restorableMessageId + ">";
                var markAsRead = _offloadOptions.MarkAsSeen;

                var message = new Message
                {
                    Subject = email.Subject ?? "(No Subject)",
                    Body = new ItemBody
                    {
                        ContentType = !string.IsNullOrEmpty(processedHtmlBody) ? BodyType.Html : BodyType.Text,
                        Content = !string.IsNullOrEmpty(processedHtmlBody) ? processedHtmlBody : (email.Body ?? "(No Content)")
                    },
                    From = new Recipient
                    {
                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                        {
                            Address = email.From ?? "unknown@unknown.com",
                            Name = !string.IsNullOrEmpty(email.FromDisplayName) ? email.FromDisplayName : (email.From ?? "Unknown Sender")
                        }
                    },
                    ToRecipients = ParseEmailAddressesWithDisplayNames(email.To, email.ToDisplayNames),
                    CcRecipients = ParseEmailAddressesWithDisplayNames(email.Cc, email.CcDisplayNames),
                    BccRecipients = ParseEmailAddressesWithDisplayNames(email.Bcc, email.BccDisplayNames),
                    SentDateTime = email.SentDate,
                    ReceivedDateTime = email.SentDate,
                    InternetMessageId = internetMessageId,
                    // Same switch as the IMAP path (Offload:MarkAsSeen, default true), so both
                    // providers restore with the same read state. Previously IsRead=false was sent
                    // together with MSGFLAG_READ, two contradicting instructions.
                    IsRead = markAsRead,
                    Importance = Importance.Normal,
                    InferenceClassification = InferenceClassificationType.Focused,
                    SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty>
                    {
                        new SingleValueLegacyExtendedProperty
                        {
                            Id = "Integer 0x0E07", // PidTagMessageFlags: no MSGFLAG_UNSENT -> not a draft
                            Value = BuildMessageFlags(markAsRead)
                        },
                        new SingleValueLegacyExtendedProperty
                        {
                            Id = "SystemTime 0x0039", // PidTagClientSubmitTime -> SentDateTime
                            Value = NormalizeToUtcIso8601(email.SentDate)
                        },
                        new SingleValueLegacyExtendedProperty
                        {
                            Id = "SystemTime 0x0E06", // PidTagMessageDeliveryTime -> ReceivedDateTime
                            Value = ResolveDeliveryTimeUtcIso8601(email)
                        }
                    }
                };

                _logger.LogDebug("Creating message with Subject: {Subject}, From: {From}, To: {To}, Body length: {BodyLength}",
                    message.Subject, message.From?.EmailAddress?.Address,
                    string.Join(", ", message.ToRecipients?.Select(r => r.EmailAddress?.Address) ?? new List<string>()),
                    message.Body?.Content?.Length ?? 0);

                var createdMessage = await graphClient.Users[targetAccount.EmailAddress].MailFolders[targetFolder.Id].Messages.PostAsync(message);

                if (createdMessage != null)
                {
                    _logger.LogInformation("Message created successfully with ID: {MessageId}", createdMessage.Id);

                    if (email.Attachments != null && email.Attachments.Any())
                    {
                        var attachmentsRestored = await RestoreAttachmentsAsync(graphClient, email, targetAccount, createdMessage.Id!);
                        if (!attachmentsRestored)
                        {
                            // A message without its attachments must not be counted as restored.
                            // Remove the partial copy so a retry starts clean instead of leaving a
                            // second, incomplete message next to the complete one.
                            await DeletePartialMessageAsync(graphClient, targetAccount, createdMessage.Id!, email.Id);
                            return false;
                        }
                    }

                    _logger.LogInformation("Successfully restored email {EmailId} to folder {FolderName} via Graph API.", 
                        email.Id, folderName);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring email {EmailId} to folder {FolderName} via Graph API: {Message}",
                    email.Id, folderName, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Restores multiple emails with progress tracking.
        /// Fetches the folder hierarchy once and reuses it, avoiding redundant Graph API calls.
        /// </summary>
        public async Task<(int Successful, int Failed)> RestoreMultipleEmailsWithProgressAsync(
            List<int> emailIds,
            int targetAccountId,
            string folderName,
            bool preserveFolderStructure,
            Action<int, int, int> progressCallback,
            CancellationToken cancellationToken = default)
        {
            int successCount = 0;
            int failCount = 0;

            var targetAccount = await _context.MailAccounts.FindAsync(targetAccountId);
            if (targetAccount == null)
            {
                _logger.LogError("Target account with ID {AccountId} not found", targetAccountId);
                return (0, emailIds.Count);
            }

            // Pre-fetch Graph client and folders once for the entire batch
            var graphClient = _authFactory.CreateGraphClient(targetAccount);
            var folders = await _folderService.GetAllMailFoldersAsync(graphClient, targetAccount.EmailAddress);

            foreach (var emailId in emailIds)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    // MEMORY FIX: Restore only reads the entity – load without tracking so the
                    // change tracker does not accumulate emails and attachment byte arrays
                    // across the entire batch.
                    var email = await _context.ArchivedEmails
                        .AsNoTracking()
                        .Include(e => e.Attachments)
                            .ThenInclude(a => a.AttachmentContent)
                        .FirstOrDefaultAsync(e => e.Id == emailId, cancellationToken);

                    if (email == null)
                    {
                        _logger.LogWarning("Email with ID {EmailId} not found during batch restore", emailId);
                        failCount++;
                        continue;
                    }

                    var result = await RestoreEmailToFolderInternalAsync(email, targetAccount, folderName, preserveFolderStructure, graphClient, folders);
                    if (result)
                        successCount++;
                    else
                        failCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    _logger.LogError(ex, "Error restoring email {EmailId} to account {AccountId}", emailId, targetAccountId);
                }

                progressCallback?.Invoke(successCount + failCount, successCount, failCount);
            }

            return (successCount, failCount);
        }

        /// <summary>
        /// Restores attachments for a newly created message. Attachments up to
        /// <see cref="SimpleAttachmentMaxBytes"/> are posted directly; larger ones go through an
        /// upload session, which Graph requires above roughly 3 MB. Before this, every large
        /// attachment failed with a 413 that was only logged, and the message was still reported
        /// as successfully restored.
        /// </summary>
        /// <returns>True only if every attachment was restored.</returns>
        private async Task<bool> RestoreAttachmentsAsync(
            GraphServiceClient graphClient,
            ArchivedEmail email,
            MailAccount targetAccount,
            string messageId)
        {
            foreach (var attachment in email.Attachments!)
            {
                try
                {
                    var content = attachment.Content;
                    var contentId = NormalizeContentId(attachment.ContentId);

                    if (RequiresUploadSession(content.LongLength))
                    {
                        await UploadLargeAttachmentAsync(graphClient, targetAccount, messageId, attachment, content, contentId);
                    }
                    else
                    {
                        var fileAttachment = new Microsoft.Graph.Models.FileAttachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = attachment.FileName,
                            ContentType = attachment.ContentType,
                            ContentBytes = content,
                            ContentId = contentId
                        };

                        await graphClient.Users[targetAccount.EmailAddress]
                            .Messages[messageId]
                            .Attachments
                            .PostAsync(fileAttachment);
                    }

                    _logger.LogInformation("Successfully restored attachment {AttachmentName} ({Size} bytes) for email {EmailId}",
                        attachment.FileName, content.LongLength, email.Id);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error restoring attachment {AttachmentName} for email {EmailId}; the message is rolled back",
                        attachment.FileName, email.Id);
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Uploads one attachment through a Graph upload session in
        /// <see cref="UploadSliceSize"/> chunks.
        /// </summary>
        private async Task UploadLargeAttachmentAsync(
            GraphServiceClient graphClient,
            MailAccount targetAccount,
            string messageId,
            EmailAttachment attachment,
            byte[] content,
            string? contentId)
        {
            var requestBody = new CreateUploadSessionPostRequestBody
            {
                AttachmentItem = new AttachmentItem
                {
                    AttachmentType = AttachmentType.File,
                    Name = attachment.FileName,
                    ContentType = attachment.ContentType,
                    Size = content.LongLength,
                    ContentId = contentId
                }
            };

            var uploadSession = await graphClient.Users[targetAccount.EmailAddress]
                .Messages[messageId]
                .Attachments
                .CreateUploadSession
                .PostAsync(requestBody);

            if (uploadSession == null)
            {
                throw new InvalidOperationException($"Graph returned no upload session for attachment '{attachment.FileName}'.");
            }

            using var stream = new MemoryStream(content, writable: false);
            var uploadTask = new LargeFileUploadTask<Microsoft.Graph.Models.FileAttachment>(
                uploadSession, stream, UploadSliceSize, graphClient.RequestAdapter);

            var result = await uploadTask.UploadAsync();
            if (!result.UploadSucceeded)
            {
                throw new InvalidOperationException($"Upload session for attachment '{attachment.FileName}' did not complete.");
            }
        }

        /// <summary>
        /// Deletes a message whose attachments could not all be restored. Failure to delete is
        /// logged loudly because it leaves an incomplete copy in the target mailbox.
        /// </summary>
        private async Task DeletePartialMessageAsync(
            GraphServiceClient graphClient,
            MailAccount targetAccount,
            string messageId,
            int emailId)
        {
            try
            {
                await graphClient.Users[targetAccount.EmailAddress].Messages[messageId].DeleteAsync();
                _logger.LogWarning("Rolled back partially restored message for email {EmailId}", emailId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex,
                    "Could not roll back partially restored message {MessageId} for email {EmailId}; an incomplete copy remains in {Mailbox}",
                    messageId, emailId, targetAccount.EmailAddress);
            }
        }

        /// <summary>Strips the angle brackets Graph does not accept around a Content-ID.</summary>
        internal static string? NormalizeContentId(string? contentId)
        {
            if (string.IsNullOrWhiteSpace(contentId)) return null;
            var trimmed = contentId.Trim();
            return trimmed.StartsWith("<") && trimmed.EndsWith(">") ? trimmed.Trim('<', '>') : trimmed;
        }

        /// <summary>
        /// Parses a comma-separated string of email addresses into a list of Graph Recipients.
        /// </summary>
        private static List<Recipient> ParseEmailAddresses(string? emailAddresses)
        {
            var recipients = new List<Recipient>();

            if (string.IsNullOrEmpty(emailAddresses))
                return recipients;

            var addresses = emailAddresses.Split(',', StringSplitOptions.RemoveEmptyEntries);

            foreach (var address in addresses)
            {
                var cleanAddress = address.Trim();
                if (!string.IsNullOrEmpty(cleanAddress))
                {
                    recipients.Add(new Recipient
                    {
                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                        {
                            Address = cleanAddress,
                            Name = cleanAddress
                        }
                    });
                }
            }

            return recipients;
        }

        /// <summary>
        /// Parses a comma-separated string of email addresses into a list of Graph Recipients,
        /// applying the stored display names when the count matches the number of addresses.
        /// Falls back to ParseEmailAddresses behavior when display names are missing or mismatched.
        /// </summary>
        private static List<Recipient> ParseEmailAddressesWithDisplayNames(string? emailAddresses, string? displayNamesCsv)
        {
            var recipients = ParseEmailAddresses(emailAddresses);

            if (string.IsNullOrEmpty(displayNamesCsv) || recipients.Count == 0)
                return recipients;

            var names = displayNamesCsv.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                       .Select(s => s.Trim())
                                       .ToArray();

            if (names.Length != recipients.Count)
                return recipients;

            for (int i = 0; i < recipients.Count; i++)
            {
                if (!string.IsNullOrEmpty(names[i]))
                    recipients[i].EmailAddress.Name = names[i];
            }

            return recipients;
        }
        private string NormalizeToUtcIso8601(DateTime value)
        {
            var utc = _dateTimeHelper.ConvertFromDisplayTimeZoneToUtc(value);
            return utc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}