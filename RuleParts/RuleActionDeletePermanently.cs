using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionDeletePermanently : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Actions.DeletePermanently.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.DeletePermanently;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions.DeletePermanently.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.DeletePermanently = true;
            return ruleModel;
        }
    }
}
