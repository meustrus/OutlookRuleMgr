using System.Collections.Generic;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleExceptionIfSubjectOrBodyContainsAny : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Exceptions.BodyOrSubject.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.ExceptIfSubjectOrBodyContainsAny != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Exceptions.BodyOrSubject.Enabled = true;
            rule.Exceptions.BodyOrSubject.Text = ruleModel.ExceptIfSubjectOrBodyContainsAny.ToArray();
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.ExceptIfSubjectOrBodyContainsAny = new List<string>((string[]) rule.Exceptions.BodyOrSubject.Text);
            return ruleModel;
        }
    }
}
