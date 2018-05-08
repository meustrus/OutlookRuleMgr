using Microsoft.Office.Interop.Outlook;
using OutlookRuleMgr.Utilities;

namespace OutlookRuleMgr.Commands
{
    public class Clear : ICommand
    {
        private static readonly Logger Log = Logger.GetLogger<Clear>();

        public string Command => "clear";
        public string[] Args => new string[0];

        public void Execute(Application outlook, string[] args)
        {
            Log.Info("Starting rules clear");

            var rules = outlook.Session.DefaultStore.GetRules();

            for (var i = rules.Count; i >= 1; --i)
            {
                rules.Remove(i);
            }

            Log.Info("Saving changes to rules...");
            rules.Save();
            Log.Info("Finished rules clear");
        }
    }
}
