using System;
using System.IO;

namespace OutlookRuleMgr.Utilities
{
    public class Logger
    {
        private const int VisibleLevel = 7;
        private const int ErrorLevel = 3;

        private readonly string _className;

        private Logger(string className)
        {
            _className = className;
        }

        public static Logger GetLogger<T>() => new Logger(typeof(T).Name);

        public void Debug(string message) => LogMessage(message, LogLevel.Debug);
        public void Info(string message) => LogMessage(message, LogLevel.Info);
        public void Warn(string message) => LogMessage(message, LogLevel.Warn);

        private static void LogMessage(string message, LogLevel level)
            => GetTextWriter(level)?.WriteLine(message);

        private static TextWriter GetTextWriter(LogLevel level)
            => (int) level <= VisibleLevel
                ? (int) level <= ErrorLevel ? Console.Error : Console.Out
                : null;

        private enum LogLevel
        {
            Debug = 7,
            Info = 6,
            Warn = 4,
            Error = 3,
        }
    }
}
