using System.Linq;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.RuleParts;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.Commands
{
    public class Import : ICommand
    {
        private static readonly Logger Log = Logger.GetLogger<Import>();

        public string Command => "import";
        public string[] Args => new[] {"rulefile..."};

        public void Execute(Outlook.Application outlook, string[] args)
        {
            var rules = outlook.Session.DefaultStore.GetRules();
            var ruleParts = typeof(Import).Assembly.GetImplementations<IRulePart>().ToList();

            foreach (var rulefilename in args.ReverseList())
            {
                Log.Info($"Starting rules import from {rulefilename}");

                var rulefile = Json.ReadFile<OutlookExport>(rulefilename);

                foreach (var ruleModel in rulefile.ReceiveRules?.ReverseList() ?? Enumerable.Empty<Rule>())
                {
                    Log.Info($"Importing {Outlook.OlRuleType.olRuleReceive} rule \"{ruleModel.Name}\"...");
                    ruleParts.Where(r => r.IsEnabled(ruleModel))
                        .Aggregate(rules.Create(ruleModel.Name, Outlook.OlRuleType.olRuleReceive),
                            (newRule, r) => r.ApplyToOutlook(newRule, ruleModel));
                }

                foreach (var ruleModel in rulefile.SendRules?.ReverseList() ?? Enumerable.Empty<Rule>())
                {
                    Log.Info($"Importing {Outlook.OlRuleType.olRuleSend} rule \"{ruleModel.Name}\"...");
                    ruleParts.Where(r => r.IsEnabled(ruleModel))
                        .Aggregate(rules.Create(ruleModel.Name, Outlook.OlRuleType.olRuleSend),
                            (newRule, r) => r.ApplyToOutlook(newRule, ruleModel));
                }
            }

            Log.Info("Saving changes to rules...");
            rules.Save();
            Log.Info("Finished rules import");
        }
    }
}
