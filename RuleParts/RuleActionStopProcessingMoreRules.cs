using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionStopProcessingMoreRules : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Actions.Stop.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.StopProcessingMoreRules;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions.Stop.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.StopProcessingMoreRules = true;
            return ruleModel;
        }
    }
}
