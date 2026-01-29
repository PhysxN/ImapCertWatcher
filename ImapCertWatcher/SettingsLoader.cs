using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public static class SettingsLoader
    {
        // ==========================
        // PUBLIC API
        // ==========================

        public static ServerSettings LoadServer(string path)
        {
            var all = LoadRaw(path);

            var s = new ServerSettings();

            // ===== MAIL =====
            s.MailHost = Get(all, "MailHost");
            s.MailPort = GetInt(all, "MailPort", 993);
            s.MailUseSsl = GetBool(all, "MailUseSsl", true);
            s.MailLogin = Get(all, "MailLogin");
            s.MailPassword = Get(all, "MailPassword");

            // ===== IMAP =====
            s.ImapNewCertificatesFolder = Get(all, "ImapNewCertificatesFolder", "INBOX");
            s.ImapRevocationsFolder = Get(all, "ImapRevocationsFolder", "INBOX");

            // ===== FIREBIRD =====
            s.FirebirdDbPath = Get(all, "FirebirdDbPath");
            s.FbServer = Get(all, "FbServer", "127.0.0.1");
            s.FbUser = Get(all, "FbUser", "SYSDBA");
            s.FbPassword = Get(all, "FbPassword");
            s.FbDialect = GetInt(all, "FbDialect", 3);
            s.FbCharset = Get(all, "FbCharset", "UTF8");

            // ===== SERVER =====
            s.CheckIntervalMinutes = GetInt(all, "CheckIntervalMinutes", 60);
            s.NotifyDaysThreshold = GetInt(all, "NotifyDaysThreshold", 10);
            s.NotifyOnlyInWorkHours = GetBool(all, "NotifyOnlyInWorkHours", true);
            s.IsDevelopment = GetBool(all, "IsDevelopment", false);

            // ===== BIMOID =====
            s.BimoidAccountsKrasnoflotskaya =
                Get(all, "BimoidAccountsKrasnoflotskaya").Replace("\\n", Environment.NewLine);

            s.BimoidAccountsPionerskaya =
                Get(all, "BimoidAccountsPionerskaya").Replace("\\n", Environment.NewLine);

            return s;
        }


        public static ClientSettings LoadClient(string path)
        {
            var all = LoadRaw(path);

            var c = new ClientSettings();

            c.ServerIp = Get(all, "ServerIp", "127.0.0.1");
            c.ServerPort = GetInt(all, "ServerPort", 5050);

            c.AutoStart = GetBool(all, "AutoStart", false);
            c.MinimizeToTrayOnClose = GetBool(all, "MinimizeToTrayOnClose", true);
            c.DarkTheme = GetBool(all, "DarkTheme", false);

            return c;
        }


        // ==========================
        // INTERNAL PARSER
        // ==========================

        private static Dictionary<string, string> LoadRaw(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("settings.txt not found", path);

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();

                if (string.IsNullOrEmpty(line)) continue;
                if (line.StartsWith("#")) continue;

                int p = line.IndexOf('=');
                if (p <= 0) continue;

                string key = line.Substring(0, p).Trim();
                string val = line.Substring(p + 1).Trim();

                dict[key] = val;
            }

            return dict;
        }

        private static string Get(Dictionary<string, string> d, string key, string def = "")
        {
            return d.TryGetValue(key, out var v) ? v : def;
        }

        private static int GetInt(Dictionary<string, string> d, string key, int def)
        {
            return int.TryParse(Get(d, key), out var v) ? v : def;
        }

        private static bool GetBool(Dictionary<string, string> d, string key, bool def)
        {
            return bool.TryParse(Get(d, key), out var v) ? v : def;
        }
    }
}
