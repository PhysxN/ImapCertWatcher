using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public class AppSettings
    {
        public string MailHost { get; set; }
        public int MailPort { get; set; }
        public bool MailUseSsl { get; set; }
        public string MailLogin { get; set; }
        public string MailPassword { get; set; }
        public string ImapFolder { get; set; } = "INBOX";
        public string FilterRecipient { get; set; }
        public string FilterSubjectPrefix { get; set; } = "Сертификат №";
        public string FirebirdDbPath { get; set; }
        public string FbServer { get; set; } = "127.0.0.1";
        public string FbUser { get; set; } = "SYSDBA";
        public string FbPassword { get; set; } = "masterkey";
        public int CheckIntervalHours { get; set; } = 1;
        public int FbDialect { get; set; } = 3;
    }

    public static class SettingsLoader
    {
        public static AppSettings Load(string path)
        {
            var s = new AppSettings();
            if (!File.Exists(path))
                throw new FileNotFoundException("settings.txt not found", path);

            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                var l = line.Trim();
                if (string.IsNullOrWhiteSpace(l) || l.StartsWith("#")) continue;
                var idx = l.IndexOf('=');
                if (idx <= 0) continue;
                var key = l.Substring(0, idx).Trim();
                var val = l.Substring(idx + 1).Trim();

                if (string.IsNullOrEmpty(val)) continue;

                switch (key)
                {
                    case "MailHost": s.MailHost = val; break;
                    case "MailPort": s.MailPort = int.Parse(val); break;
                    case "MailUseSsl": s.MailUseSsl = bool.Parse(val); break;
                    case "MailLogin": s.MailLogin = val; break;
                    case "MailPassword": s.MailPassword = val; break;
                    case "ImapFolder": s.ImapFolder = val; break;
                    case "FilterRecipient": s.FilterRecipient = val; break;
                    case "FilterSubjectPrefix": s.FilterSubjectPrefix = val; break;
                    case "FirebirdDbPath": s.FirebirdDbPath = val; break;
                    case "FbServer": s.FbServer = val; break;
                    case "FbUser": s.FbUser = val; break;
                    case "FbPassword": s.FbPassword = val; break;
                    case "CheckIntervalHours": s.CheckIntervalHours = int.Parse(val); break; // Изменили
                    case "FbDialect": s.FbDialect = int.Parse(val); break;
                }
            }
            return s;
        }
    }
}