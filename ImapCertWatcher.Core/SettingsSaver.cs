using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public static class SettingsSaver
    {
        public static void SaveClient(string path, ClientSettings client)
        {
            if (client == null)
                throw new ArgumentNullException(nameof(client));

            EnsureDirectoryExists(path);

            var lines = new List<string>
            {
                "# ===== CLIENT =====",
                "",
                $"ServerIp={client.ServerIp ?? string.Empty}",
                $"ServerPort={client.ServerPort}",
                $"DarkTheme={client.DarkTheme}"
            };

            File.WriteAllLines(path, lines, Encoding.UTF8);
        }

        public static void SaveServer(string path, ServerSettings server)
        {
            if (server == null)
                throw new ArgumentNullException(nameof(server));

            EnsureDirectoryExists(path);

            var lines = new List<string>
            {
                "# ===== MAIL =====",
                $"MailHost={server.MailHost ?? string.Empty}",
                $"MailPort={server.MailPort}",
                $"MailUseSsl={server.MailUseSsl}",
                $"MailLogin={server.MailLogin ?? string.Empty}",
                $"MailPassword={ProtectIfNotEmpty(server.MailPassword)}",
                "",

                "# ===== IMAP =====",
                $"ImapNewCertificatesFolder={server.ImapNewCertificatesFolder ?? string.Empty}",
                $"ImapRevocationsFolder={server.ImapRevocationsFolder ?? string.Empty}",
                "",

                "# ===== FIREBIRD =====",
                $"FirebirdDbPath={server.FirebirdDbPath ?? string.Empty}",
                $"FbServer={server.FbServer ?? string.Empty}",
                $"FbUser={server.FbUser ?? string.Empty}",
                $"FbPassword={server.FbPassword ?? string.Empty}",
                $"FbDialect={server.FbDialect}",
                $"FbCharset={server.FbCharset ?? string.Empty}",
                $"IsDevelopment={server.IsDevelopment}",
                "",

                "# ===== SERVER =====",
                $"ServerPort={server.ServerPort}",
                $"CheckIntervalMinutes={server.CheckIntervalMinutes}",
                $"NotifyDaysThreshold={server.NotifyDaysThreshold}",
                $"NotifyOnlyInWorkHours={server.NotifyOnlyInWorkHours}",
                $"AutoStartServer={server.AutoStartServer}",
                $"MinimizeToTrayOnClose={server.MinimizeToTrayOnClose}",
                "",

                "# ===== BIMOID =====",
                $"BimoidSenderExePath={server.BimoidSenderExePath ?? string.Empty}",
                $"BimoidJobDirectory={server.BimoidJobDirectory ?? string.Empty}",
                $"BimoidServer={server.BimoidServer ?? string.Empty}",
                $"BimoidPort={server.BimoidPort}",
                $"BimoidLogin={server.BimoidLogin ?? string.Empty}",
                $"BimoidPassword={ProtectIfNotEmpty(server.BimoidPassword)}",
                $"BimoidDelayBetweenMessagesMs={server.BimoidDelayBetweenMessagesMs}",
                $"BimoidAccountsKrasnoflotskaya={NormalizeMultilineSetting(server.BimoidAccountsKrasnoflotskaya)}",
                $"BimoidAccountsPionerskaya={NormalizeMultilineSetting(server.BimoidAccountsPionerskaya)}"
            };

            File.WriteAllLines(path, lines, Encoding.UTF8);
        }

        private static void EnsureDirectoryExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Путь к файлу настроек пустой.", nameof(path));

            var dir = Path.GetDirectoryName(path);

            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        private static string NormalizeMultilineSetting(string value)
        {
            return string.IsNullOrEmpty(value)
                ? string.Empty
                : value.Replace(Environment.NewLine, "\\n");
        }

        private static string ProtectIfNotEmpty(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return CryptoHelper.Protect(value);
        }
    }
}