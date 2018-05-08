using System;
using System.Linq;
using OutlookRuleMgr.Commands;
using OutlookRuleMgr.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr
{
    public static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            var commandName = args.FirstOrDefault();
            var commandArgs = args.Skip(1).ToArray();

            var commands = typeof(Program).Assembly.GetImplementations<ICommand>().ToArray();

            if (commandName == null || commandArgs.Length < commands.Min(c => c.Args.Length))
            {
                Console.WriteLine("Usage:");
                commands.Select(DescribeCommand).ToList().ForEach(Console.WriteLine);
                return;
            }

            var command = commands.FirstOrDefault(c => Matches(c, commandName, commandArgs));

            if (command == null)
            {
                Console.Error.WriteLine($"Command \"{commandName}\" with {commandArgs.Length} arguments is not valid.");
                Console.Error.WriteLine("Available commands:");
                commands.Select(DescribeCommand).ToList().ForEach(Console.WriteLine);
                return;
            }

            command.Execute(GetOutlook(), commandArgs);
        }

        private static string DescribeCommand(ICommand c) =>
            $"{AppDomain.CurrentDomain.FriendlyName} {c.Command} {string.Join(" ", c.Args.Select(a => $"{{{a}}}"))}";

        private static bool Matches(ICommand c, string commandName, string[] commandArgs)
            => c.Command == commandName
               && (c.Args.LastOrDefault()?.EndsWith("...") ?? false
                   ? commandArgs.Length >= c.Args.Length
                   : commandArgs.Length == c.Args.Length);

        private static Outlook.Application GetOutlook()
        {
            var outlook = new Outlook.Application();

            var mapi = outlook.GetNamespace("MAPI");
            mapi.Logon("", "", true, true);

            if (mapi.Offline)
            {
                outlook.ActiveExplorer().CommandBars.ExecuteMso("ToggleOnline");
            }

            return outlook;
        }
    }
}
