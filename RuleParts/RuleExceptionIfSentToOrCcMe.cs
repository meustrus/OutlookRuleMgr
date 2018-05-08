using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleExceptionIfSentToOrCcMe : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Exceptions.ToOrCc.Enabled;
        public bool IsEnabled(Models.Rule ruleModel) => ruleModel.ExceptIfSentToOrCcMe;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Exceptions.ToOrCc.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.ExceptIfSentToOrCcMe = true;
            return ruleModel;
        }
    }
}
