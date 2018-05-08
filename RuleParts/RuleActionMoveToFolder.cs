using System;
using OutlookRuleMgr.Models;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.RuleParts
{
    public class RuleActionMoveToFolder : IRulePart
    {
        public bool IsEnabled(Outlook.Rule rule) => rule.Actions.MoveToFolder.Enabled;
        public bool IsEnabled(Rule ruleModel) => ruleModel.MoveToFolder != null;

        public Outlook.Rule ApplyToOutlook(Outlook.Rule rule, Rule ruleModel)
        {
            rule.Actions.MoveToFolder.Enabled = true;
            rule.Actions.MoveToFolder.Folder = rule.Application.GetRootFolder().Folders.GetFolder(ruleModel.MoveToFolder);
            return rule;
        }

        public Rule ApplyToModel(Rule ruleModel, Outlook.Rule rule)
        {
            var rootFolderPath = rule.Application.GetRootFolder().FolderPath;
            if (!rule.Actions.MoveToFolder.Folder.FolderPath.StartsWith(rootFolderPath,
                StringComparison.Ordinal))
            {
                throw new InvalidFolderException(
                    $"Folder {rule.Actions.MoveToFolder.Folder.FolderPath} is outside default root");
            }

            ruleModel.MoveToFolder = rule.Actions.MoveToFolder.Folder.FolderPath.Substring(rootFolderPath.Length + 1);
            return ruleModel;
        }
    }
}
