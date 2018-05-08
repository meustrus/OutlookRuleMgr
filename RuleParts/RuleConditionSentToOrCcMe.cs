using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionSentToOrCcMe : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.ToOrCc.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.SentToOrCcMe;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.ToOrCc.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.SentToOrCcMe = true;
            return ruleModel;
        }
    }
}
