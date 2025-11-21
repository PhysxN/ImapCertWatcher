using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ImapCertWatcher.Services
{
    public class ImapWatcher
    {
        private readonly AppSettings _settings;
        private readonly DbHelper _db;

        // регекс для извлечения ФИО и дат
        // ФИО: Овчинникова Светлана Анатольевна
        private static Regex fioRegex = new Regex(@"ФИО:\s*(?<fio>[\p{L}\s\-]+)", RegexOptions.Compiled);
        // Срок действия сертификата: с 21.11.2025 09:45:18 по 14.02.2027 09:45:18
        private static Regex datesRegex = new Regex(@"с\s*(?<ds>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})\s*по\s*(?<de>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public ImapWatcher(AppSettings settings, DbHelper db)
        {
            _settings = settings;
            _db = db;
        }

        public List<CertEntry> CheckMail()
        {
            var results = new List<CertEntry>();
            using (var client = new ImapClient())
            {
                client.Timeout = 30 * 1000;
                if (_settings.MailUseSsl)
                    client.Connect(_settings.MailHost, _settings.MailPort, MailKit.Security.SecureSocketOptions.SslOnConnect);
                else
                    client.Connect(_settings.MailHost, _settings.MailPort, MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);

                client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                var inbox = client.GetFolder(_settings.ImapFolder);
                inbox.Open(FolderAccess.ReadOnly);

                // получим все письма, фильтруя по теме и получателю
                // лучший вариант: использовать поиск по SUBJECT и TO
                var query = SearchQuery.SubjectContains(_settings.FilterSubjectPrefix)
                              .And(SearchQuery.DeliveredTo(_settings.FilterRecipient));
                var uids = inbox.Search(query);

                foreach (var uid in uids)
                {
                    var msg = inbox.GetMessage(uid);
                    var subject = msg.Subject ?? "";
                    var recipients = string.Join(", ", msg.To.Mailboxes.Select(m => m.Address));
                    var body = msg.TextBody ?? msg.HtmlBody ?? (msg.Body?.ToString() ?? "");

                    // проверим точнее, что тема начинается с нужного префикса
                    if (!subject.StartsWith(_settings.FilterSubjectPrefix, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // поиск ФИО и дат
                    var fioMatch = fioRegex.Match(body);
                    var datesMatch = datesRegex.Match(body);

                    if (!fioMatch.Success || !datesMatch.Success)
                        continue; // не найдено необходимых данных

                    var fio = fioMatch.Groups["fio"].Value.Trim();

                    if (!DateTime.TryParseExact(datesMatch.Groups["ds"].Value.Trim(),
                        "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"),
                        System.Globalization.DateTimeStyles.None, out DateTime dateStart))
                        continue;
                    if (!DateTime.TryParseExact(datesMatch.Groups["de"].Value.Trim(),
                        "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"),
                        System.Globalization.DateTimeStyles.None, out DateTime dateEnd))
                        continue;

                    var daysLeft = (int)Math.Ceiling((dateEnd - DateTime.Now).TotalDays);
                    if (daysLeft < 0) daysLeft = 0;

                    var entry = new CertEntry
                    {
                        MailUid = uid.ToString(),
                        Fio = fio,
                        DateStart = dateStart,
                        DateEnd = dateEnd,
                        DaysLeft = daysLeft,
                        Subject = subject,
                        Received = msg.Date.UtcDateTime.ToLocalTime()
                    };

                    // сохраняем в БД
                    try
                    {
                        _db.InsertOrUpdate(entry);
                    }
                    catch (Exception ex)
                    {
                        // логирование — для простоты просто игнорируем или можно добавить лог-файл
                    }

                    results.Add(entry);
                }

                client.Disconnect(true);
            }

            return results;
        }
    }
}
