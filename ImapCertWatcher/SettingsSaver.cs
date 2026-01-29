using ImapCertWatcher.Utils;
using System.IO;
using System.Text;

public static class SettingsSaver
{
    public static void Save(string path, ClientSettings client, ServerSettings server)
    {
        var lines = new[]
        {
            // ===== CLIENT =====
            $"ServerIp={client.ServerIp}",
            $"ServerPort={client.ServerPort}",
            $"AutoStart={client.AutoStart}",
            $"MinimizeToTrayOnClose={client.MinimizeToTrayOnClose}",

            // ===== SERVER =====
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
