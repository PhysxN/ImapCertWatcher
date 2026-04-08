using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public static class SettingsLoader
    {
        public static ServerSettings LoadServer(string path)
        {
            var all = LoadRaw(path);

            var s = new ServerSettings();

            // ===== MAIL =====
            s.MailHost = Get(all, "MailHost");
            s.MailPort = GetInt(all, "MailPort", 993);
            s.MailUseSsl = GetBool(all, "MailUseSsl", true);
            s.MailLogin = Get(all, "MailLogin");

            var encMailPassword = Get(all, "MailPassword");
            s.MailPassword = UnprotectIfNotEmpty(encMailPassword);

            // ===== IMAP =====
            s.ImapNewCertificatesFolder = Get(all, "ImapNewCertificatesFolder", "");
            s.ImapRevocationsFolder = Get(all, "ImapRevocationsFolder", "");

            // ===== FIREBIRD =====
            s.FirebirdDbPath = Get(all, "FirebirdDbPath");
            s.FbServer = Get(all, "FbServer", "127.0.0.1");
            s.FbUser = Get(all, "FbUser", "SYSDBA");
            s.FbPassword = Get(all, "FbPassword", "");
            s.FbDialect = GetInt(all, "FbDialect", 3);
            s.FbCharset = Get(all, "FbCharset", "UTF8");
            s.IsDevelopment = GetBool(all, "IsDevelopment", false);

            // ===== SERVER =====
            s.ServerPort = GetInt(all, "ServerPort", 5050);
            s.CheckIntervalMinutes = GetInt(all, "CheckIntervalMinutes", 60);
            s.NotifyDaysThreshold = GetInt(all, "NotifyDaysThreshold", 10);
            s.NotifyOnlyInWorkHours = GetBool(all, "NotifyOnlyInWorkHours", true);
            s.AutoStartServer = GetBool(all, "AutoStartServer", false);
            s.MinimizeToTrayOnClose = GetBool(all, "MinimizeToTrayOnClose", true);

            // ===== BIMOID =====
            s.BimoidSenderExePath = Get(all, "BimoidSenderExePath", @"BimoidBroadcastSender\BimoidBroadcastSender.exe");
            s.BimoidJobDirectory = Get(all, "BimoidJobDirectory", "BimoidJobs");
            s.BimoidServer = Get(all, "BimoidServer", "");
            s.BimoidPort = GetInt(all, "BimoidPort", 7023);
            s.BimoidLogin = Get(all, "BimoidLogin", "");

            var encBimoidPassword = Get(all, "BimoidPassword");
            s.BimoidPassword = UnprotectIfNotEmpty(encBimoidPassword);

            s.BimoidDelayBetweenMessagesMs = GetInt(all, "BimoidDelayBetweenMessagesMs", 300);

            s.BimoidAccountsKrasnoflotskaya =
                NormalizeMultilineSetting(Get(all, "BimoidAccountsKrasnoflotskaya", ""));

            s.BimoidAccountsPionerskaya =
                NormalizeMultilineSetting(Get(all, "BimoidAccountsPionerskaya", ""));

            return s;
        }

        public static ClientSettings LoadClient(string path)
        {
            var all = LoadRaw(path);

            var c = new ClientSettings
            {
                ServerIp = Get(all, "ServerIp", "127.0.0.1"),
                ServerPort = GetInt(all, "ServerPort", 5050),
                DarkTheme = GetBool(all, "DarkTheme", false)
            };

            return c;
        }

        private static Dictionary<string, string> LoadRaw(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Путь к файлу настроек пустой.", nameof(path));

            if (!File.Exists(path))
                throw new FileNotFoundException($"Settings file not found: {path}", path);

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();

                if (string.IsNullOrEmpty(line))
                    continue;

                if (line.StartsWith("#"))
                    continue;

                int p = line.IndexOf('=');
                if (p <= 0)
                    continue;

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

        private static string NormalizeMultilineSetting(string value)
        {
            return string.IsNullOrEmpty(value)
                ? string.Empty
                : value.Replace("\\n", Environment.NewLine);
        }

        private static string UnprotectIfNotEmpty(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return CryptoHelper.Unprotect(value);
        }
    }
}