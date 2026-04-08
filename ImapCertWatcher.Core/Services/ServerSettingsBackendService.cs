using ImapCertWatcher.Data;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ImapCertWatcher.Services
{
    public class ServerSettingsBackendService
    {
        public async Task<string[]> LoadFoldersAsync(ServerSettings settings, string mailPassword)
        {
            var result = new List<string>();

            using (var client = new ImapClient())
            {
                await client.ConnectAsync(
                    settings.MailHost,
                    settings.MailPort,
                    settings.MailUseSsl
                        ? SecureSocketOptions.SslOnConnect
                        : SecureSocketOptions.StartTlsWhenAvailable);

                await client.AuthenticateAsync(
                    settings.MailLogin,
                    mailPassword);

                var root = client.GetFolder(client.PersonalNamespaces[0]);
                var folders = await root.GetSubfoldersAsync(false);

                foreach (var folder in folders)
                    await LoadFoldersRecursive(folder, result);

                await client.DisconnectAsync(false);
            }

            return result.ToArray();
        }

        private async Task LoadFoldersRecursive(IMailFolder folder, List<string> acc)
        {
            acc.Add(folder.FullName);

            var subs = await folder.GetSubfoldersAsync(false);

            foreach (var sub in subs)
                await LoadFoldersRecursive(sub, acc);
        }

        public async Task TestMailAsync(ServerSettings settings, string mailPassword)
        {
            using (var client = new ImapClient())
            {
                await client.ConnectAsync(
                    settings.MailHost,
                    settings.MailPort,
                    settings.MailUseSsl
                        ? SecureSocketOptions.SslOnConnect
                        : SecureSocketOptions.StartTlsWhenAvailable);

                await client.AuthenticateAsync(
                    settings.MailLogin,
                    mailPassword);

                await client.DisconnectAsync(false);
            }
        }

        public void TestDb(ServerSettings settings, string dbPassword)
        {
            var testSettings = new ServerSettings
            {
                FbServer = settings.FbServer,
                FirebirdDbPath = settings.FirebirdDbPath,
                FbUser = settings.FbUser,
                FbPassword = dbPassword,
                FbDialect = settings.FbDialect,
                FbCharset = settings.FbCharset,
                IsDevelopment = settings.IsDevelopment
            };

            var db = new DbHelper(testSettings);
            db.TestConnection();
        }
    }
}