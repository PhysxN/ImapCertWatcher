using ImapCertWatcher.Utils;
using System.Text;
using System.IO;

public static class SettingsSaver
{
    public static void Save(string path, AppSettings s)
    {
        var lines = new[]
        {
            $"MailHost={s.MailHost}",
            $"MailPort={s.MailPort}",
            $"MailUseSsl={s.MailUseSsl}",
            $"MailLogin={s.MailLogin}",
            $"MailPassword={s.MailPassword}",

            $"ImapFolder={s.ImapFolder}",

            $"FirebirdDbPath={s.FirebirdDbPath}",
            $"FbServer={s.FbServer}",
            $"FbUser={s.FbUser}",
            $"FbPassword={s.FbPassword}",
            $"FbCharset={s.FbCharset}",

            $"CheckIntervalMinutes={s.CheckIntervalMinutes}",

            $"NotifyDaysThreshold={s.NotifyDaysThreshold}",
            $"BimoidAccountsKrasnoflotskaya={s.BimoidAccountsKrasnoflotskaya}",
            $"BimoidAccountsPionerskaya={s.BimoidAccountsPionerskaya}",
        };

        File.WriteAllLines(path, lines, Encoding.UTF8);
    }
}
