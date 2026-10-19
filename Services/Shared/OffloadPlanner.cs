// Services/Shared/OffloadPlanner.cs
using MailArchiver.Models;

namespace MailArchiver.Services.Shared
{
    /// <summary>
    /// The provider independent half of an offload: which archived mail goes into which target
    /// folder. Shared by the IMAP and the M365 offload so both apply exclusions, renames and the
    /// target root identically.
    /// </summary>
    public static class OffloadPlanner
    {
        /// <summary>The source folder assumed for a row that has none.</summary>
        public const string DefaultSourceFolder = "INBOX";

        /// <summary>
        /// Groups mail by target path, after exclusion and rename.
        /// <para>
        /// Grouping happens on the path AFTER the rename. Keying on the raw source name would open
        /// the same target folder twice when two source folders legitimately collapse onto one,
        /// for example "Sent Items" and "Sent" both becoming "Sent".
        /// </para>
        /// </summary>
        /// <param name="rows">Id and stored source folder of each archived message.</param>
        /// <param name="excludedCount">How many rows were dropped by a folder exclusion.</param>
        public static Dictionary<string, List<int>> GroupByTargetPath(
            IEnumerable<(int Id, string? FolderName)> rows,
            string? baseFolderName,
            bool preserveFolderStructure,
            OffloadCriteria criteria,
            out int excludedCount)
        {
            excludedCount = 0;
            var groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            foreach (var (id, folderName) in rows)
            {
                var source = folderName ?? DefaultSourceFolder;

                if (FolderMapper.IsExcluded(source, criteria.ExcludedSourceFolders))
                {
                    excludedCount++;
                    continue;
                }

                var renamed = FolderMapper.ApplyRenameMap(source, criteria.FolderRenameMap);
                var targetPath = FolderMapper.ResolveTargetPath(renamed, baseFolderName, preserveFolderStructure);

                if (!groups.TryGetValue(targetPath, out var list))
                {
                    list = new List<int>();
                    groups[targetPath] = list;
                }
                list.Add(id);
            }

            return groups;
        }
    }
}
