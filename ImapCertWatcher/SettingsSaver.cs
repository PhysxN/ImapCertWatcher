using ImapCertWatcher.Utils;
using System.Collections.Generic;
using System.Text;
using System.IO;

public static class SettingsSaver
{
    public static void SaveClient(string path, ClientSettings client)
    {
        var lines = new List<string>
        {
            "# ===== CLIENT =====",
            "",
            $"ServerIp={client.ServerIp}",
            $"ServerPort={client.ServerPort}",
            $"DarkTheme={client.DarkTheme}"
        };

        File.WriteAllLines(path, lines, Encoding.UTF8);
    }

    public static void SaveServer(string path, ServerSettings server)
    {
        var lines = new List<string>
        {
            "# ===== SERVER =====",
            "",
            $"AutoStartServer={server.AutoStartServer}",

            $"MailHost={server.MailHost}",
            $"MailPort={server.MailPort}",
            $"MailUseSsl={server.MailUseSsl}",
            $"MailLogin={server.MailLogin}",
            $"MailPassword={server.MailPassword}",

            $"ImapNewCertificatesFolder={server.ImapNewCertificatesFolder}",
            $"ImapRevocationsFolder={server.ImapRevocationsFolder}",

            $"FirebirdDbPath={server.FirebirdDbPath}",
            $"FbServer={server.FbServer}",
            $"FbUser={server.FbUser}",
            $"FbPassword={server.FbPassword}",
            $"FbDialect={server.FbDialect}",
            $"FbCharset={server.FbCharset}",

            $"CheckIntervalMinutes={server.CheckIntervalMinutes}",
            $"NotifyDaysThreshold={server.NotifyDaysThreshold}",
            $"NotifyOnlyInWorkHours={server.NotifyOnlyInWorkHours}",

            $"BimoidAccountsKrasnoflotskaya={server.BimoidAccountsKrasnoflotskaya}",
            $"BimoidAccountsPionerskaya={server.BimoidAccountsPionerskaya}"
        };

        File.WriteAllLines(path, lines, Encoding.UTF8);
    }
}
