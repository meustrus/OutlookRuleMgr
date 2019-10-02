using System.Collections.Generic;

namespace OutlookRuleMgr.Models
{
    public class Rule
    {
        public string Name { get; set; }

        public List<Recipient> FromRecipients { get; set; }
        public List<Recipient> ToRecipients { get; set; }
        public List<string> SubjectContainsAny { get; set; }
        public List<string> BodyContainsAny { get; set; }
        public bool SentToOrCcMe { get; set; }
        public bool OnLocalMachineOnly { get; set; }

        public string MoveToFolder { get; set; }
        public List<string> AssignToCategories { get; set; }
        public bool ClearCategories { get; set; }
        public bool DeletePermanently { get; set; }
        public bool MarkAsRead { get; set; }
        public bool StopProcessingMoreRules { get; set; }

        public bool ExceptIfSentToOrCcMe { get; set; }
        public List<string> ExceptIfSubjectContainsAny { get; set; }
        public List<string> ExceptIfSubjectOrBodyContainsAny { get; set; }
    }
}
