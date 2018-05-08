using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.Commands
{
    public interface ICommand
    {
        string Command { get; }
        string[] Args { get; }
        void Execute(Outlook.Application outlook, string[] args);
    }
}
