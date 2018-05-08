using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.Utilities
{
    public static class OutlookExtensions
    {
        private const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public static Outlook.MAPIFolder GetRootFolder(this Outlook.Application outlook)
        {
            return (Outlook.MAPIFolder) outlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent;
        }

        public static Outlook.MAPIFolder GetFolder(this Outlook.Folders folders, string path)
        {
            var pathSeparatorIndex = path.IndexOf("\\", StringComparison.Ordinal);
            var currentFolderName = pathSeparatorIndex < 0 ? path : path.Substring(0, pathSeparatorIndex);

            var matchingFolder =
                folders.OfType<Outlook.MAPIFolder>()
                    .FirstOrDefault(f => f.Name == currentFolderName) ??
                folders.Add(currentFolderName);

            return pathSeparatorIndex < 0
                ? matchingFolder
                : GetFolder(matchingFolder.Folders, path.Substring(pathSeparatorIndex + 1));
        }

        public static string GetEmailAddress(this Outlook.Recipient recipient)
            => recipient.Address == null ? null : recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;

        public static IEnumerable<T> ReverseList<T>(this IReadOnlyList<T> list)
        {
            for (var i = list.Count - 1; i >= 0; --i)
            {
                yield return list[i];
            }
        }
    }
}
