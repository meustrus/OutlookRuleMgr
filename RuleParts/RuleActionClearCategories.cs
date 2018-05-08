using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionClearCategories : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Actions.ClearCategories.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.ClearCategories;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions.ClearCategories.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.ClearCategories = true;
            return ruleModel;
        }
    }
}
