using System.Collections.Generic;
using System.Linq;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.RuleParts;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.Commands
{
    public class Export : ICommand
    {
        private static readonly Logger Log = Logger.GetLogger<Export>();

        public string Command => "export";
        public string[] Args => new[] {"rulefile"};

        public void Execute(Outlook.Application outlook, string[] args)
        {
            Log.Info("Starting rules export");

            var rulefile = new OutlookExport
            {
                ReceiveRules = new List<Rule>(),
                SendRules = new List<Rule>()
            };

            var rules = outlook.Session.DefaultStore.GetRules();
            var ruleParts = typeof(Export).Assembly.GetImplementations<IRulePart>().ToList();

            foreach (var rule in rules.OfType<Outlook.Rule>())
            {
                Log.Info($"Exporting rule \"{rule.Name}\"...");
                var ruleModel = ruleParts.Where(r => r.IsEnabled(rule))
                    .Aggregate(new Rule {Name = rule.Name}, (model, r) => r.ApplyToModel(model, rule));

                switch (rule.RuleType)
                {
                    case Outlook.OlRuleType.olRuleReceive:
                        rulefile.ReceiveRules.Add(ruleModel);
                        break;
                    case Outlook.OlRuleType.olRuleSend:
                        rulefile.SendRules.Add(ruleModel);
                        break;
                }
            }

            Log.Info("Writing rulefile...");
            Json.WriteFile(rulefile, args[0]);
            Log.Info("Finished rules export");
        }
    }
}
