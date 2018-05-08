using System.Collections.Generic;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionBodyContainsAny : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.Body.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.BodyContainsAny != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.Body.Enabled = true;
            rule.Conditions.Body.Text = ruleModel.BodyContainsAny.ToArray();
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.BodyContainsAny = new List<string>((string[]) rule.Conditions.Body.Text);
            return ruleModel;
        }
    }
}
