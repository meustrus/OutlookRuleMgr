using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public interface IRulePart
    {
        bool IsEnabled(Outlook.Rule rule);
        bool IsEnabled(Rule ruleModel);
        Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel);
        Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule);
    }
}
