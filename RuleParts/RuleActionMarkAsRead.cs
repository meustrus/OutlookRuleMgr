using System.Linq;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionMarkAsRead : IRulePart
    {
        private static readonly Logger Log = Logger.GetLogger<RuleActionMarkAsRead>();

        public bool IsEnabled(Outlook.Rule rule) =>
            rule.Actions
                .OfType<Outlook.RuleAction>()
                .Any(x => x.Enabled && x.ActionType == Outlook.OlRuleActionType.olRuleActionMarkRead);

        public bool IsEnabled(Rule ruleModel) => ruleModel.MarkAsRead;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            Log.Warn($"For rule '{rule.Name}', action 'MarkAsRead' cannot be enabled through the COM interface");
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.MarkAsRead = true;
            return ruleModel;
        }
    }
}
