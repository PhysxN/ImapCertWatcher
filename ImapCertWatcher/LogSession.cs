using System;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public static class LogSession
    {
        public static readonly string SessionId;
        public static readonly string DayDirectory;
        public static readonly string SessionLogFile;

        static LogSession()
        {
            SessionId = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            DayDirectory = Path.Combine(logDir, DateTime.Now.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(DayDirectory))
                Directory.CreateDirectory(DayDirectory);

            SessionLogFile = Path.Combine(DayDirectory, $"session_{SessionId}.log");

            // Заголовок новой сессии
            File.AppendAllText(
                SessionLogFile,
                $"=== Сессия запущена: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={Environment.NewLine}{Environment.NewLine}",
                Encoding.UTF8
            );
        }
    }
}
