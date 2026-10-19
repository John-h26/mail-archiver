using MailArchiver.Data;
using MailArchiver.Models;
using MailArchiver.Services.Providers.Imap;
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
        private readonly BatchOperationOptions _batchOptions;

        public GraphMailRestorer(
            GraphAuthClientFactory authFactory,
            IGraphFolderService folderService,
            MailArchiverDbContext context,
            ILogger<GraphMailRestorer> logger,
            DateTimeHelper dateTimeHelper,
            IOptions<OffloadOptions> offloadOptions,
            IOptions<BatchOperationOptions> batchOptions)
        {
            _authFactory = authFactory;
            _folderService = folderService;
            _context = context;
            _logger = logger;
            _dateTimeHelper = dateTimeHelper;
            _offloadOptions = offloadOptions.Value;
            _batchOptions = batchOptions.Value;
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

                var created = await CreateMessageInFolderAsync(
                    email, targetAccount, targetFolder, graphClient, _offloadOptions.MarkAsSeen);

                if (created)
                {
                    _logger.LogInformation("Successfully restored email {EmailId} to folder {FolderName} via Graph API.",
                        email.Id, folderName);
                }

                return created;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring email {EmailId} to folder {FolderName} via Graph API: {Message}",
                    email.Id, folderName, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Creates one archived message, with its attachments, in an already resolved folder.
        /// Shared by the restore and the offload path. Returns false, after removing any partial
        /// copy, when the message or one of its attachments could not be created; Graph errors on
        /// the message itself propagate to the caller.
        /// </summary>
        private async Task<bool> CreateMessageInFolderAsync(
            ArchivedEmail email,
            MailAccount targetAccount,
            MailFolder targetFolder,
            GraphServiceClient graphClient,
            bool markAsRead,
            CancellationToken cancellationToken = default)
        {
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
                // Same read state as the IMAP path: Offload:MarkAsSeen for a restore, the job's
                // own MarkAsSeen for an offload. Previously IsRead=false was sent together with
                // MSGFLAG_READ, two contradicting instructions.
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

            var createdMessage = await graphClient.Users[targetAccount.EmailAddress].MailFolders[targetFolder.Id].Messages.PostAsync(message, cancellationToken: cancellationToken);

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

                return true;
            }

            return false;
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
        /// Appends archived mail into an M365 mailbox, skipping anything the mailbox already
        /// holds. The M365 counterpart of <c>ImapMailRestorer.OffloadEmailsAsync</c>, with the
        /// same contract: folder exclusions and renames, a duplicate index over the whole target
        /// mailbox built before the first write, a dry run that writes nothing, and an outcome
        /// that reports appended, already present and failed mail per target folder.
        /// </summary>
        public async Task<OffloadOutcome> OffloadEmailsAsync(
            List<int> emailIds,
            int targetAccountId,
            string baseFolderName,
            bool preserveFolderStructure,
            OffloadCriteria criteria,
            Action<int, int, int>? progressCallback = null,
            CancellationToken cancellationToken = default)
        {
            var outcome = new OffloadOutcome { DryRun = criteria.DryRun };

            var targetAccount = await _context.MailAccounts.FindAsync(new object[] { targetAccountId }, cancellationToken);
            if (targetAccount == null || targetAccount.Provider != ProviderType.M365)
            {
                _logger.LogError("Offload target account {AccountId} not found or not an M365 account", targetAccountId);
                outcome.Failed = emailIds.Count;
                return outcome;
            }

            // Only the fields the mapping needs, so a large window does not load message bodies.
            var rows = await _context.ArchivedEmails
                .Where(e => emailIds.Contains(e.Id))
                .Select(e => new { e.Id, e.FolderName })
                .ToListAsync(cancellationToken);

            var groups = OffloadPlanner.GroupByTargetPath(
                rows.Select(r => (r.Id, (string?)r.FolderName)),
                baseFolderName, preserveFolderStructure, criteria, out var excludedCount);
            outcome.SkippedExcludedFolder = excludedCount;

            _logger.LogInformation(
                "M365 offload to account {AccountName}: {Considered} mails in {Groups} target folders, " +
                "{Excluded} skipped by folder exclusion, window {Window}, dry run: {DryRun}",
                targetAccount.Name, rows.Count - excludedCount, groups.Count,
                excludedCount, criteria.DescribeWindow(), criteria.DryRun);

            var processed = 0;
            var mailbox = targetAccount.EmailAddress;

            try
            {
                var graphClient = _authFactory.CreateGraphClient(targetAccount);
                var folders = await _folderService.GetAllMailFoldersAsync(graphClient, mailbox);

                // Built once, before the first write, over every folder of the mailbox. This is
                // what makes the offload repeatable.
                var index = await GraphTargetMailboxIndexBuilder.BuildAsync(
                    graphClient, mailbox, folders, _dateTimeHelper, _logger,
                    _offloadOptions.PrefetchMaxMessages, cancellationToken);
                outcome.DuplicateScopeReduced = index.ScopeReduced;

                // "INBOX" is the IMAP name. In M365 the Inbox carries a localised display name
                // ("Boîte de réception", "Posteingang"), so it is resolved through its well-known
                // name instead of creating a second, literal "INBOX" folder.
                var inbox = await _folderService.GetWellKnownInboxAsync(graphClient, mailbox, folders);

                foreach (var group in groups)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var targetPath = group.Key;
                    var folderOutcome = outcome.Folder(targetPath);

                    MailFolder? folder = null;
                    if (!criteria.DryRun)
                    {
                        folder = await ResolveOffloadFolderAsync(graphClient, mailbox, targetPath, folders, inbox);
                        if (folder == null)
                        {
                            _logger.LogError("Could not resolve or create target folder '{Folder}', skipping {Count} mails",
                                targetPath, group.Value.Count);
                            outcome.Failed += group.Value.Count;
                            folderOutcome.Failed += group.Value.Count;
                            processed += group.Value.Count;
                            progressCallback?.Invoke(processed, outcome.Appended, outcome.Failed);
                            continue;
                        }
                    }

                    for (var i = 0; i < group.Value.Count; i += _batchOptions.BatchSize)
                    {
                        var batchIds = group.Value.Skip(i).Take(_batchOptions.BatchSize).ToList();
                        var emails = await _context.ArchivedEmails
                            .AsNoTracking()
                            .Include(e => e.Attachments)
                                .ThenInclude(a => a.AttachmentContent)
                            .Where(e => batchIds.Contains(e.Id))
                            .ToDictionaryAsync(e => e.Id, cancellationToken);

                        foreach (var emailId in batchIds)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            if (!emails.TryGetValue(emailId, out var email))
                            {
                                outcome.Failed++;
                                folderOutcome.Failed++;
                                processed++;
                                continue;
                            }

                            var match = index.Match(email);
                            if (match != OffloadMatchKind.None)
                            {
                                outcome.SkippedAlreadyPresent++;
                                folderOutcome.SkippedAlreadyPresent++;
                                if (match == OffloadMatchKind.Fingerprint)
                                {
                                    outcome.MatchedByFingerprint++;
                                    folderOutcome.MatchedByFingerprint++;
                                }
                                processed++;
                                progressCallback?.Invoke(processed, outcome.Appended, outcome.Failed);
                                continue;
                            }

                            if (criteria.DryRun)
                            {
                                // Counts are real, only the write is skipped. The index is still
                                // updated so two copies inside the window are not both reported.
                                outcome.Appended++;
                                folderOutcome.Appended++;
                                index.Add(email);
                                processed++;
                                progressCallback?.Invoke(processed, outcome.Appended, outcome.Failed);
                                continue;
                            }

                            try
                            {
                                var ok = await CreateMessageInFolderAsync(
                                    email, targetAccount, folder!, graphClient, criteria.MarkAsSeen, cancellationToken);
                                if (ok)
                                {
                                    outcome.Appended++;
                                    folderOutcome.Appended++;
                                    index.Add(email);
                                }
                                else
                                {
                                    outcome.Failed++;
                                    folderOutcome.Failed++;
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                throw;
                            }
                            catch (Exception ex)
                            {
                                outcome.Failed++;
                                folderOutcome.Failed++;
                                _logger.LogError(ex, "Error creating email {EmailId} in '{Folder}': {Message}",
                                    emailId, targetPath, ex.Message);
                            }

                            processed++;
                            progressCallback?.Invoke(processed, outcome.Appended, outcome.Failed);

                            if (_batchOptions.PauseBetweenEmailsMs > 0)
                            {
                                await Task.Delay(_batchOptions.PauseBetweenEmailsMs, cancellationToken);
                            }
                        }
                    }

                    if (!criteria.DryRun && _batchOptions.PauseBetweenBatchesMs > 0)
                    {
                        await Task.Delay(_batchOptions.PauseBetweenBatchesMs, cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogInformation("M365 offload was cancelled after {Processed} of {Total} mails", processed, rows.Count);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Critical error during M365 offload: {Message}", ex.Message);
                outcome.Failed += Math.Max(0, rows.Count - excludedCount - processed);
            }

            _logger.LogInformation("M365 offload finished. {Summary}", outcome.Describe().Replace(Environment.NewLine, " | "));
            return outcome;
        }

        /// <summary>
        /// Resolves an offload target path in an M365 mailbox, creating missing folders.
        /// </summary>
        private async Task<MailFolder?> ResolveOffloadFolderAsync(
            GraphServiceClient graphClient,
            string mailbox,
            string targetPath,
            List<MailFolder> folders,
            MailFolder? inbox)
        {
            if (inbox != null && IsInboxRootOnly(targetPath))
            {
                return inbox;
            }

            var mappedPath = MapInboxRoot(targetPath, inbox?.DisplayName);
            return await _folderService.EnsureFolderPathAsync(graphClient, mailbox, mappedPath, folders, createIfMissing: true);
        }

        /// <summary>Whether a target path is the IMAP Inbox itself, without subfolders.</summary>
        internal static bool IsInboxRootOnly(string? path)
        {
            var segments = SplitPath(path);
            return segments.Length == 1 && string.Equals(segments[0], OffloadPlanner.DefaultSourceFolder, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Replaces a leading IMAP "INBOX" segment with the display name of the mailbox's real
        /// Inbox, so "INBOX/Projects" lands below the existing Inbox in a localised mailbox
        /// instead of under a new top-level folder literally named "INBOX". Other paths are
        /// returned normalised but unchanged.
        /// </summary>
        internal static string MapInboxRoot(string? path, string? inboxDisplayName)
        {
            var segments = SplitPath(path);
            if (segments.Length == 0) return string.Empty;

            if (!string.IsNullOrWhiteSpace(inboxDisplayName)
                && string.Equals(segments[0], OffloadPlanner.DefaultSourceFolder, StringComparison.OrdinalIgnoreCase))
            {
                segments[0] = inboxDisplayName.Trim();
            }

            return string.Join("/", segments);
        }

        private static string[] SplitPath(string? path)
            => string.IsNullOrWhiteSpace(path)
                ? Array.Empty<string>()
                : path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries)
                      .Select(p => p.Trim())
                      .Where(p => p.Length > 0)
                      .ToArray();

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