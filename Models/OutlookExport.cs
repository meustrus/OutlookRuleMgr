using System.Collections.Generic;

namespace OutlookRuleMgr.Models
{
    public class OutlookExport
    {
        public List<Rule> ReceiveRules { get; set; }
        public List<Rule> SendRules { get; set; }
    }
}
