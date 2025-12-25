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

namespace ImapCertWatcher.Services
{
    public class ImapRevocationsWatcher
    {
        private readonly AppSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        // Универсальный regex для номера сертификата
        private static readonly Regex CertNumberRegex = new Regex(
            @"Сертификат\s*№\s*([0-9A-Fa-f]{20,})",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ImapRevocationsWatcher(
            AppSettings settings,
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
        public void ProcessRevocations(bool checkAllMessages)
        {
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

                    client.Authenticate(
                        _settings.MailLogin,
                        _settings.MailPassword);

                    var folder = client.GetFolder(_settings.ImapFolder);
                    folder.Open(FolderAccess.ReadOnly);

                    var uids = checkAllMessages
                        ? folder.Search(SearchQuery.All)
                        : folder.Search(SearchQuery.DeliveredAfter(DateTime.Now.AddDays(-7)));

                    Log($"Найдено писем: {uids.Count}");

                    foreach (var uid in uids)
                    {
                        string uidStr = uid.ToString();

                        if (_db.IsMailProcessed(folder.FullName, uidStr, "REVOKE"))
                            continue;

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

                        bool ok = _db.FindAndMarkAsRevokedByCertNumber(
                            certNumber,
                            null,
                            folder.FullName,
                            null);

                        if (ok)
                        {
                            applied++;
                            Log($"Аннулирован сертификат: {certNumber}");
                        }
                        else
                        {
                            Log($"Сертификат не найден в БД: {certNumber}");
                        }

                        MarkProcessed(folder, uidStr);

                        if (!checkAllMessages)
                            break;
                    }

                    folder.Close();
                    client.Disconnect(true);
                }
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
