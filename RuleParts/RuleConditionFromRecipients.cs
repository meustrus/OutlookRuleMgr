using System.Linq;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionFromRecipients : IRulePart
    {
        private readonly RecipientModelMapper _recipientModelMapper = new RecipientModelMapper();

        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.From.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.FromRecipients != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.From.Enabled = true;

            foreach (var recipientModel in ruleModel.FromRecipients)
            {
                _recipientModelMapper.AddToOutlook(recipientModel, rule.Conditions.From.Recipients);
            }

            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.FromRecipients = rule.Conditions.From.Recipients
                .OfType<Outlook.Recipient>()
                .Select(_recipientModelMapper.MapToModel)
                .ToList();
            return ruleModel;
        }
    }
}
