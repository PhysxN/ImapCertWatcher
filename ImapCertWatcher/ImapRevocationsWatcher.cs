using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Globalization;
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

        // Строгая тема аннулирования:
        // "Сертификат № 008954001058EE95915C0F731E29907571 аннулирован или прекратил действие"
        private static readonly Regex SubjectRevokeRegex = new Regex(
            @"^Сертификат\s*№\s*([0-9A-Fa-f]+)\s+(аннулирован|прекратил\s+действие)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ImapRevocationsWatcher(AppSettings settings, DbHelper db, Action<string> log = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
        }

        // Добавляем метод для файлового логирования
        private void FileLog(string message)
        {
            try
            {
                string logFile = ImapCertWatcher.Utils.LogSession.SessionLogFile;
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [REVOKE] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);
            }
            catch
            {
                // игнорируем ошибки логирования
            }
        }

        private void Log(string msg)
        {
            _log?.Invoke(msg); // Мини-лог
            FileLog(msg);      // Файловый лог
        }

        /// <summary>
        /// Основной публичный метод.
        /// Возвращает количество сертификатов, которые удалось пометить как аннулированные.
        /// </summary>
        public int ProcessRevocations(bool checkAllMessages = false)
        {
            int applied = 0;

            Log("=== ImapRevocationsWatcher: старт поиска аннулированных сертификатов ===");

            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    if (_settings.MailUseSsl)
                        client.Connect(_settings.MailHost, _settings.MailPort, SecureSocketOptions.SslOnConnect);
                    else
                        client.Connect(_settings.MailHost, _settings.MailPort, SecureSocketOptions.StartTlsWhenAvailable);

                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    IMailFolder rootFolder = null;
                    try
                    {
                        rootFolder = client.GetFolder(_settings.ImapFolder);
                    }
                    catch (Exception ex)
                    {
                        Log($"[REVOKE] Не удалось получить корневую папку '{_settings.ImapFolder}': {ex.Message}");
                    }

                    if (rootFolder == null)
                    {
                        Log($"[REVOKE] Папка '{_settings.ImapFolder}' не найдена. Проверка отменена.");
                        client.Disconnect(true);
                        return applied;
                    }

                    Log($"[REVOKE] Корневая папка: '{rootFolder.FullName}' — начинаем рекурсивный обход (checkAll={checkAllMessages}).");

                    ProcessFolderRecursive(client, rootFolder, checkAllMessages, ref applied);

                    client.Disconnect(true);
                }
            }
            catch (Exception exTop)
            {
                Log($"[REVOKE] Общая ошибка ImapRevocationsWatcher: {exTop}");
            }

            Log($"=== ImapRevocationsWatcher: завершено. Аннулирований применено={applied} ===");
            return applied;
        }

        /// <summary>
        /// Рекурсивная обработка одной папки и её подпапок.
        /// </summary>
        private void ProcessFolderRecursive(ImapClient client, IMailFolder folder, bool checkAllMessages, ref int applied)
        {
            if (folder == null) return;

            Log($"[REVOKE] ---> Обработка папки '{folder.FullName}'");

            try
            {
                folder.Open(FolderAccess.ReadOnly);
            }
            catch (Exception exOpen)
            {
                Log($"[REVOKE] Не удалось открыть папку '{folder.FullName}': {exOpen.Message}");
                return;
            }

            IList<UniqueId> uids;
            try
            {
                if (checkAllMessages)
                    uids = folder.Search(SearchQuery.All);
                else
                    // Аннулирования могут приходить задним числом — берём за последний год
                    uids = folder.Search(SearchQuery.DeliveredAfter(DateTime.Now.AddDays(-365)));

                Log($"[REVOKE][Search] Папка='{folder.FullName}', checkAll={checkAllMessages}, найдено UID={uids?.Count ?? 0}");
            }
            catch (Exception exSearch)
            {
                Log($"[REVOKE] Ошибка Search в папке '{folder.FullName}': {exSearch.Message}");
                try { folder.Close(); } catch { }
                return;
            }

            var summaries = new List<IMessageSummary>();
            if (uids != null && uids.Count > 0)
            {
                try
                {
                    summaries.AddRange(
                        folder.Fetch(
                            uids,
                            MessageSummaryItems.Envelope |
                            MessageSummaryItems.InternalDate |
                            MessageSummaryItems.UniqueId
                        )
                    );
                    Log($"[REVOKE][Fetch] Папка='{folder.FullName}', summaries={summaries.Count}");

                    int preview = 0;
                    foreach (var s in summaries)
                    {
                        var subj = s.Envelope?.Subject ?? "(no subject)";
                        var dt = s.InternalDate?.ToString() ?? "(no date)";
                        Log($"[REVOKE][Preview] Folder='{folder.FullName}' UID={s.UniqueId} Date={dt} Subject='{subj}'");
                        if (++preview >= 20) break;
                    }
                }
                catch (Exception exFetch)
                {
                    Log($"[REVOKE] Ошибка Fetch summaries в папке '{folder.FullName}': {exFetch.Message}");
                }
            }

            int folderApplied = 0;

            foreach (var summary in summaries)
            {
                var uid = summary.UniqueId;
                var uidStr = uid.ToString();
                var folderPath = folder.FullName;

                try
                {
                    // 1) Проверка: письмо уже обрабатывали как REVOKE?
                    if (_db.IsMailProcessed(folderPath, uidStr, "REVOKE"))
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: уже помечено как обработанное (REVOKE), пропускаем.");
                        continue;
                    }

                    string subject = summary.Envelope?.Subject ?? "";
                    if (string.IsNullOrWhiteSpace(subject))
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: пустая тема, пропуск.");
                        continue;
                    }

                    var mSubj = SubjectRevokeRegex.Match(subject);
                    if (!mSubj.Success)
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: тема не соответствует шаблону аннулирования, пропуск. Subject='{subject}'");
                        continue;
                    }

                    string certNumberFromSubject = mSubj.Groups[1].Value.Trim();
                    Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: тема подходит, номер из темы={certNumberFromSubject}");

                    MimeMessage message;
                    try
                    {
                        message = folder.GetMessage(uid);
                    }
                    catch (Exception exMsg)
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: ошибка GetMessage: {exMsg.Message}");
                        continue;
                    }

                    var bodyText = GetMessageBody(message) ?? string.Empty;

                    // Извлекаем ФИО и дату отзыва
                    string fio = ExtractFio(bodyText);
                    var revokeDate = ExtractRevokeDate(bodyText);
                    string revokeDateShort = revokeDate?.ToString("dd.MM.yyyy") ?? "";

                    // Номер сертификата из тела (на случай, если в теме другая форма)
                    string certNumberFromBody = ExtractCertNumber(bodyText);
                    string certNumber = !string.IsNullOrEmpty(certNumberFromBody)
                        ? certNumberFromBody
                        : certNumberFromSubject;

                    if (string.IsNullOrEmpty(certNumber) && string.IsNullOrEmpty(fio))
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: не удалось извлечь ни номер сертификата, ни ФИО — пропуск.");
                        continue;
                    }

                    Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: найдено cert='{certNumber}', fio='{fio}', revokeDate='{revokeDateShort}'");

                    bool ok = _db.FindAndMarkAsRevokedByCertNumber(
                        certNumber,
                        fio,
                        folderPath,
                        revokeDateShort
                    );

                    if (ok)
                    {
                        applied++;
                        folderApplied++;
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: аннулирован сертификат {certNumber} — {fio}");
                    }
                    else
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: сертификат для аннулирования не найден в БД: {certNumber} — {fio}");
                    }

                    // Мы письмо обработали (даже если в БД не нашли сертификат) — помечаем UID
                    try
                    {
                        _db.MarkMailProcessed(folderPath, uidStr, "REVOKE");
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: помечено как обработанное (REVOKE) в PROCESSED_MAILS.");
                    }
                    catch (Exception exMark)
                    {
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: ошибка MarkMailProcessed(REVOKE): {exMark.Message}");
                    }

                    if (!checkAllMessages)
                    {
                        // В "быстром" режиме можно выходить после первого подходящего письма
                        Log($"[REVOKE] Folder='{folderPath}' UID={uidStr}: быстрый режим, выходим после первого письма.");
                        break;
                    }
                }
                catch (Exception exItem)
                {
                    Log($"[REVOKE] Ошибка обработки письма Folder='{folder.FullName}' UID={uidStr}: {exItem}");
                }
            }

            Log($"[REVOKE] Папка='{folder.FullName}': применено аннулирований={folderApplied}");

            try { folder.Close(); } catch { }

            // Рекурсивно обрабатываем подпапки
            try
            {
                var subfolders = folder.GetSubfolders(false).ToList();
                Log($"[REVOKE] Папка='{folder.FullName}': найдено подпапок={subfolders.Count}");
                foreach (var sub in subfolders)
                {
                    ProcessFolderRecursive(client, sub, checkAllMessages, ref applied);
                }
            }
            catch (Exception exSub)
            {
                Log($"[REVOKE] Ошибка получения подпапок для '{folder.FullName}': {exSub.Message}");
            }
        }

        #region Helpers

        private static string GetMessageBody(MimeMessage message)
        {
            if (message == null) return string.Empty;
            if (!string.IsNullOrEmpty(message.TextBody)) return message.TextBody;
            if (!string.IsNullOrEmpty(message.HtmlBody))
            {
                return Regex.Replace(message.HtmlBody, "<.*?>", string.Empty)
                    .Replace("&nbsp;", " ");
            }

            var textParts = message.BodyParts.OfType<TextPart>().ToList();
            if (textParts.Any()) return textParts.First().Text;
            return message.Body?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// ФИО из блока:
        /// "ФИО: Кулькова Алина Алексеевна. Срок действия сертификата..."
        /// Берём до первой точки или перевода строки.
        /// </summary>
        private static string ExtractFio(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            // 1) Основной вариант: строго три "слова" ФИО
            var m = Regex.Match(text,
                @"ФИО[:\s]*([А-ЯЁа-яё\-]+\s+[А-ЯЁа-яё\-]+\s+[А-ЯЁа-яё\-]+)",
                RegexOptions.IgnoreCase);
            if (m.Success)
                return m.Groups[1].Value.Trim().Trim('.', ',', ';');

            // 2) Запасной вариант – до конца строки/предложения, потом режем по ключевым словам
            m = Regex.Match(text,
                @"ФИО[:\s]*([^\r\n]+)",
                RegexOptions.IgnoreCase);
            if (m.Success)
            {
                var fio = m.Groups[1].Value;

                // Отрезаем служебные фразы, если они попали в захват
                string[] stopMarkers =
                {
            "Срок действия сертификата",
            "Срок действия",
            "Дата отзыва сертификата",
            "Дата отзыва",
            "Организация"
        };

                foreach (var marker in stopMarkers)
                {
                    var idx = fio.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0)
                    {
                        fio = fio.Substring(0, idx);
                        break;
                    }
                }

                fio = fio.Trim().Trim('.', ',', ';');
                // На всякий случай ограничим ФИО максимум 3 "словами"
                var parts = fio.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2 && parts.Length <= 4)
                    fio = string.Join(" ", parts.Take(3));

                return fio;
            }

            return null;
        }


        /// <summary>
        /// Номер сертификата из текста (на всякий случай, если в теме другая форма).
        /// </summary>
        private static string ExtractCertNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            var m = Regex.Match(text,
                @"Сертификат\s*№\s*([0-9A-Fa-f]+)",
                RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();

            m = Regex.Match(text,
                @"Файл\s+сертификата\s*№\s*([0-9A-Fa-f]+)",
                RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();

            m = Regex.Match(text,
                @"\b([0-9A-Fa-f]{20,})\b",
                RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();

            return null;
        }

        /// <summary>
        /// Дата отзыва сертификата из текста:
        /// "Дата отзыва сертификата: 17.04.2025 10:43:29 (по московскому времени)."
        /// </summary>
        private static DateTime? ExtractRevokeDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            var m = Regex.Match(text,
                @"Дата\s+отзыва\s+сертификата[:\s]*([\d\.:\s]+)",
                RegexOptions.IgnoreCase);

            if (m.Success)
            {
                var val = m.Groups[1].Value.Trim();
                if (DateTime.TryParseExact(
                        val,
                        new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime dt))
                {
                    return dt;
                }
            }

            // запасной вариант: первая дата-строка формата "dd.MM.yyyy HH:mm:ss"
            m = Regex.Match(text,
                @"\b\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}\b");
            if (m.Success &&
                DateTime.TryParseExact(
                    m.Value,
                    "dd.MM.yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal,
                    out DateTime dt2))
            {
                return dt2;
            }

            return null;
        }

        #endregion
    }
}