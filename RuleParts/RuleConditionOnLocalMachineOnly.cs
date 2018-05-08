using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionOnLocalMachineOnly : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.OnLocalMachine.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.OnLocalMachineOnly;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.OnLocalMachine.Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.OnLocalMachineOnly = true;
            return ruleModel;
        }
    }
}
