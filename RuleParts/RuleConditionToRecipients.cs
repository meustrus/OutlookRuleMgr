using System.Linq;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleConditionToRecipients : IRulePart
    {
        private readonly RecipientModelMapper _recipientModelMapper = new RecipientModelMapper();

        public bool IsEnabled(Outlook.Rule rule) => rule.Conditions.SentTo.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.ToRecipients != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Conditions.SentTo.Enabled = true;

            foreach (var recipientModel in ruleModel.ToRecipients)
            {
                _recipientModelMapper.AddToOutlook(recipientModel, rule.Conditions.SentTo.Recipients);
            }

            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            ruleModel.ToRecipients = rule.Conditions.SentTo.Recipients
                .OfType<Outlook.Recipient>()
                .Select(_recipientModelMapper.MapToModel)
                .ToList();
            return ruleModel;
        }
    }
}
