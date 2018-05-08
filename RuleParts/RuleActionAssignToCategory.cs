using System;
using System.Collections.Generic;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionAssignToCategory : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Actions.AssignToCategory.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.AssignToCategories != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions.AssignToCategory.Enabled = true;
            rule.Actions.AssignToCategory.Categories = ruleModel.AssignToCategories.ToArray();
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.AssignToCategories = new List<string>((string[]) rule.Actions.AssignToCategory.Categories);
            return ruleModel;
        }
    }
}
