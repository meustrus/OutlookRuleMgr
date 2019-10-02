using System.Linq;
using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionMarkAsRead : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) =>
            rule.Actions
                .OfType<Outlook.RuleAction>()
                .Any(x => x.Enabled && x.ActionType == Outlook.OlRuleActionType.olRuleActionMarkRead);

        public bool IsEnabled(Rule ruleModel) => ruleModel.MarkAsRead;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions
                .OfType<Outlook.RuleAction>()
                .First(x => x.ActionType == Outlook.OlRuleActionType.olRuleActionMarkRead)
                .Enabled = true;
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.MarkAsRead = true;
            return ruleModel;
        }
    }
}
