using System.Collections.Generic;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionSubjectContainsAny : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.Subject.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.SubjectContainsAny != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.Subject.Enabled = true;
            rule.Conditions.Subject.Text = ruleModel.SubjectContainsAny.ToArray();
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.SubjectContainsAny = new List<string>((string[]) rule.Conditions.Subject.Text);
            return ruleModel;
        }
    }
}
