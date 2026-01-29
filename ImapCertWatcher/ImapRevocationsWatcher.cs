using ImapCertWatcher.Data;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ImapCertWatcher.Services
{
    public class ImapRevocationsWatcher
    {
        private readonly ServerSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        // Универсальный regex для номера сертификата
        private static readonly Regex CertNumberRegex = new Regex(
            @"сертификат\s*№?\s*([0-9A-Fa-f]{20,})",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ImapRevocationsWatcher(
            ServerSettings settings,
            DbHelper db,
            Action<string> log)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
        }

        private void Log(string message)
        {
            _log.Invoke("[REVOKE] " + message);

            try
            {
                string file = LogSession.SessionLogFile;
                File.AppendAllText(
                    file,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [REVOKE] {message}{Environment.NewLine}");
            }
            catch { }
        }

        /// <summary>
        /// Основной метод. Счётчики ведутся внутри watcher’а.
        /// </summary>
        public void ProcessRevocations(bool checkAllMessages, CancellationToken token)
        {
            if (token.IsCancellationRequested)
            {
                Log("Обработка аннулирований прервана пользователем.");
                return;
            }

            int processed = 0;
            int applied = 0;

            Log("Старт обработки аннулирований");

            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    client.Connect(
                        _settings.MailHost,
                        _settings.MailPort,
                        _settings.MailUseSsl
                            ? SecureSocketOptions.SslOnConnect
                            : SecureSocketOptions.StartTlsWhenAvailable);
                    Log($"IMAP подключение: {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");

                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    string folderName = string.IsNullOrWhiteSpace(_settings.ImapRevocationsFolder) ? "INBOX" : _settings.ImapRevocationsFolder;
                    IMailFolder folder = client.GetFolder(folderName) ?? client.Inbox;

                    folder.Open(FolderAccess.ReadWrite);

                    long lastUid = checkAllMessages
                        ? 0
                        : _db.GetLastUid(folder.FullName + "_REVOKE");

                    var uids = checkAllMessages
                        ? folder.Search(SearchQuery.All)
                        : folder.Search(
                            SearchQuery.Uids(
                                new UniqueIdRange(
                                    new UniqueId((uint)(lastUid + 1)),
                                    UniqueId.MaxValue
                                )
                            )
                          );

                    Log($"Найдено писем: {uids.Count}");

                    foreach (var uid in uids)
                    {
                        if (token.IsCancellationRequested)
                        {
                            Log("Обработка аннулирований отменена");
                            break;
                        }

                        string uidStr = uid.ToString();

                        if (_db.IsMailProcessed(folder.FullName, uidStr, "REVOKE"))
                        {
                            Log($"UID={uidStr}: уже обработано, пропуск");
                            continue;
                        }

                        processed++;

                        MimeMessage message;
                        try
                        {
                            message = folder.GetMessage(uid);
                        }
                        catch
                        {
                            continue;
                        }

                        string body = GetBodyText(message);
                        if (string.IsNullOrWhiteSpace(body))
                        {
                            MarkProcessed(folder, uidStr);
                            continue;
                        }

                        var match = CertNumberRegex.Match(body);
                        if (!match.Success)
                        {
                            MarkProcessed(folder, uidStr);
                            continue;
                        }

                        string certNumber = match.Groups[1].Value.Trim();

                        bool ok = _db.FindAndMarkAsRevokedByCertNumber(certNumber, null, folder.FullName, null);

                        if (ok)
                        {
                            applied++;
                            Log($"Аннулирован сертификат: {certNumber}");
                        }

                        MarkProcessed(folder, uidStr);

                        
                            
                    }
                    if (!checkAllMessages && uids.Count > 0)
                    {
                        long maxUid = uids.Max(u => u.Id);
                        _db.UpdateLastUid(folder.FullName + "_REVOKE", maxUid);
                    }
                    folder.Close();
                    client.Disconnect(true);
                }
            }
            catch (OperationCanceledException)
            {
                Log("Обработка аннулирований отменена");
            }
            catch (Exception ex)
            {
                Log("Ошибка: " + ex.Message);
            }

            Log($"Завершено. Обработано писем={processed}, аннулировано={applied}");
        }


        private void MarkProcessed(IMailFolder folder, string uid)
        {
            try
            {
                _db.MarkMailProcessed(folder.FullName, uid, "REVOKE");
            }
            catch (Exception ex)
            {
                Log($"Ошибка MarkMailProcessed: {ex.Message}");
            }
        }

        public Task CheckRevocationsFastAsync(CancellationToken token)
        {
            return Task.Run(() =>
            {
                ProcessRevocations(false, token);
            }, token);
        }

        private static string GetBodyText(MimeMessage msg)
        {
            if (!string.IsNullOrEmpty(msg.TextBody))
                return msg.TextBody;

            if (!string.IsNullOrEmpty(msg.HtmlBody))
                return Regex.Replace(msg.HtmlBody, "<.*?>", " ");

            return string.Empty;
        }
    }
}
