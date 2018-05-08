using System.Collections.Generic;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionExceptIfSubjectContainsAny : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Exceptions.Subject.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.ExceptIfSubjectContainsAny != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Exceptions.Subject.Enabled = true;
            rule.Exceptions.Subject.Text = ruleModel.ExceptIfSubjectContainsAny.ToArray();
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.ExceptIfSubjectContainsAny = new List<string>((string[]) rule.Exceptions.Subject.Text);
            return ruleModel;
        }
    }
}
