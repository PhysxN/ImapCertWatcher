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
    /// <summary>
    /// Поиск писем с новыми сертификатами:
    /// - одна корневая папка _settings.ImapFolder + все вложенные;
    /// - первый проход: парсим тему/тело, вызываем InsertOrUpdateAndGetId;
    /// - второй проход: только для помеченных писем качаем ZIP и сохраняем в БД;
    /// - письма помечаются в PROCESSED_MAILS (KIND='NEW') по UID+папка.
    /// </summary>
    public class ImapNewCertificatesWatcher
    {
        private readonly AppSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;
        private readonly string _subjectPrefix;

        public ImapNewCertificatesWatcher(AppSettings settings, DbHelper db, Action<string> log = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
            _subjectPrefix = string.IsNullOrWhiteSpace(_settings.FilterSubjectPrefix)
                ? "Сертификат №"
                : _settings.FilterSubjectPrefix;
        }

        // Добавляем метод для файлового логирования
        private void FileLog(string message)
        {
            try
            {
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                // Используем тот же файл сессии, что и в MainWindow
                var sessionLogs = Directory.GetFiles(logDirectory, "session_*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .ToList();

                string sessionLogFile;
                if (sessionLogs.Any())
                {
                    sessionLogFile = sessionLogs.First().FullName;
                }
                else
                {
                    sessionLogFile = Path.Combine(logDirectory, $"session_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                }

                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [NEW] {message}{Environment.NewLine}";
                File.AppendAllText(sessionLogFile, logEntry, System.Text.Encoding.UTF8);
            }
            catch { /* игнорируем ошибки логирования */ }
        }

        private void Log(string msg)
        {
            _log?.Invoke(msg); // Мини-лог
            FileLog(msg);      // Файловый лог
        }

        /// <summary>
        /// Основной метод.
        /// Возвращает список успешно разобранных CertEntry и счётчики добавленных/обновлённых.
        /// </summary>
        public (List<CertEntry> processedEntries, int updatedCount, int addedCount) ProcessNewCertificates(bool checkAllMessages = false)
        {
            var results = new List<CertEntry>();
            int updatedCount = 0;
            int addedCount = 0;

            Log("=== ImapNewCertificatesWatcher: старт проверки новых сертификатов ===");

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
                        Log($"[NEW] Не удалось получить корневую папку '{_settings.ImapFolder}': {ex.Message}");
                    }

                    if (rootFolder == null)
                    {
                        Log($"[NEW] Папка '{_settings.ImapFolder}' не найдена. Проверка отменена.");
                        client.Disconnect(true);
                        return (results, updatedCount, addedCount);
                    }

                    Log($"[NEW] Корневая папка: '{rootFolder.FullName}' — начинаем рекурсивный обход (checkAll={checkAllMessages}).");

                    // Рекурсивный обход корневой и всех вложенных папок
                    ProcessFolderRecursive(client, rootFolder, checkAllMessages, results, ref updatedCount, ref addedCount);

                    client.Disconnect(true);
                }
            }
            catch (Exception exTop)
            {
                Log($"[NEW] Общая ошибка ImapNewCertificatesWatcher: {exTop}");
            }

            Log($"=== ImapNewCertificatesWatcher: завершено. Добавлено={addedCount}, обновлено={updatedCount}, всего записей={results.Count} ===");
            return (results, updatedCount, addedCount);
        }

        /// <summary>
        /// Рекурсивная обработка одной папки и всех её подпапок.
        /// </summary>
        private void ProcessFolderRecursive(
            ImapClient client,
            IMailFolder folder,
            bool checkAllMessages,
            List<CertEntry> globalResults,
            ref int updatedCount,
            ref int addedCount)
        {
            if (folder == null) return;

            Log($"[NEW] ---> Обработка папки '{folder.FullName}'");

            try
            {
                folder.Open(FolderAccess.ReadOnly);
            }
            catch (Exception exOpen)
            {
                Log($"[NEW] Не удалось открыть папку '{folder.FullName}': {exOpen.Message}");
                return;
            }

            // Получаем UIDs
            IList<UniqueId> uids;
            try
            {
                if (checkAllMessages)
                    uids = folder.Search(SearchQuery.All);
                else
                    // За последние 10 дней (можно увеличить при необходимости)
                    uids = folder.Search(SearchQuery.DeliveredAfter(DateTime.Now.AddDays(-10)));

                Log($"[NEW][Search] Папка='{folder.FullName}', checkAll={checkAllMessages}, найдено UID: {uids?.Count ?? 0}");
            }
            catch (Exception exSearch)
            {
                Log($"[NEW] Ошибка Search в папке '{folder.FullName}': {exSearch.Message}");
                try { folder.Close(); } catch { }
                return;
            }

            // Получаем summary
            var summaries = new List<IMessageSummary>();
            if (uids != null && uids.Count > 0)
            {
                try
                {
                    summaries.AddRange(
                        folder.Fetch(
                            uids,
                            MessageSummaryItems.Envelope |
                            MessageSummaryItems.BodyStructure |
                            MessageSummaryItems.InternalDate
                        )
                    );
                    Log($"[NEW][Fetch] Папка='{folder.FullName}', summaries={summaries.Count}");

                    int preview = 0;
                    foreach (var s in summaries)
                    {
                        var subj = s.Envelope?.Subject ?? "(no subject)";
                        var dt = s.InternalDate?.ToString() ?? "(no date)";
                        Log($"[NEW][Preview] Folder='{folder.FullName}' UID={s.UniqueId} Date={dt} Subject='{subj}'");
                        if (++preview >= 20) break;
                    }
                }
                catch (Exception exFetch)
                {
                    Log($"[NEW] Ошибка Fetch summaries в папке '{folder.FullName}': {exFetch.Message}");
                }
            }

            var subjectStrictRegex = new Regex(
                @"^" + Regex.Escape(_subjectPrefix) + @"\s*([0-9A-Fa-f]+)\s*$",
                RegexOptions.Compiled
            );

            var zipCandidates = new List<(UniqueId uid, int certId)>();
            int folderProcessed = 0;

            // ---------- 1-й проход: анализируем summary + тело, но не качаем ZIP ----------

            foreach (var s in summaries)
            {
                try
                {
                    var subject = s.Envelope?.Subject ?? string.Empty;
                    var uid = s.UniqueId;
                    var uidStr = uid.ToString();
                    var folderPath = folder.FullName;

                    // Проверка: уже обрабатывали это письмо как NEW?
                    if (_db.IsMailProcessed(folderPath, uidStr, "NEW"))
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uidStr}: уже помечено как обработанное (NEW), пропускаем.");
                        continue;
                    }

                    // Строгая проверка темы
                    var m = subjectStrictRegex.Match(subject);
                    if (!m.Success)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: тема не соответствует шаблону, пропуск. Subject='{subject}'");
                        continue;
                    }

                    MimeMessage message;
                    try
                    {
                        message = folder.GetMessage(uid);
                    }
                    catch (Exception exMsg)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: ошибка загрузки письма: {exMsg.Message}");
                        continue;
                    }

                    var bodyText = GetMessageBody(message) ?? string.Empty;
                    var fromAddress = message.From.Mailboxes.FirstOrDefault()?.Address ?? string.Empty;

                    var fio = ExtractFio(bodyText);
                    var certFromSubject = m.Groups[1].Value.Trim();
                    var certNumber = !string.IsNullOrEmpty(certFromSubject)
                        ? certFromSubject
                        : ExtractCertNumber(bodyText);
                    var dates = ExtractDates(bodyText);

                    if (string.IsNullOrWhiteSpace(fio)
                        || string.IsNullOrWhiteSpace(certNumber)
                        || dates.start == DateTime.MinValue
                        || dates.end == DateTime.MinValue)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: недостаточно данных (FIO/Cert/Dates). " +
                            $"fio='{fio}', cert='{certNumber}', start={dates.start}, end={dates.end}");
                        continue;
                    }

                    var entry = new CertEntry
                    {
                        Fio = fio,
                        CertNumber = certNumber,
                        DateStart = dates.start,
                        DateEnd = dates.end,
                        DaysLeft = (int)(dates.end - DateTime.Now).TotalDays,
                        FromAddress = fromAddress,
                        FolderPath = folderPath,
                        MailUid = uidStr,
                        MessageDate = s.InternalDate?.UtcDateTime ?? message.Date.UtcDateTime,
                        Subject = subject
                    };

                    try
                    {
                        var (wasUpdated, wasAdded, certId) = _db.InsertOrUpdateAndGetId(entry);
                        if (wasAdded) addedCount++;
                        if (wasUpdated) updatedCount++;

                        Log($"[NEW] Folder='{folderPath}' UID={uid}: InsertOrUpdateAndGetId -> certId={certId}, added={wasAdded}, updated={wasUpdated}");

                        globalResults.Add(entry);
                        folderProcessed++;

                        bool hasZipInSummary = BodyStructureHasZip(s);
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: BodyStructureHasZip={hasZipInSummary}");

                        if (hasZipInSummary && certId > 0)
                        {
                            zipCandidates.Add((uid, certId));
                            Log($"[NEW] Folder='{folderPath}' UID={uid}: добавлен в очередь скачивания ZIP (certId={certId})");
                        }
                    }
                    catch (Exception exDb)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: ошибка InsertOrUpdateAndGetId: {exDb.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"[NEW] Ошибка обработки summary в папке '{folder.FullName}': {ex}");
                }
            }

            Log($"[NEW] Папка='{folder.FullName}': 1-й проход завершён. Обработано писем={folderProcessed}, zipCandidates={zipCandidates.Count}");

            // ---------- 2-й проход: скачиваем и сохраняем zip-вложения ----------

            foreach (var (uid, certId) in zipCandidates)
            {
                var uidStr = uid.ToString();
                var folderPath = folder.FullName;

                try
                {
                    MimeMessage message;
                    try
                    {
                        message = folder.GetMessage(uid);
                    }
                    catch (Exception ex)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: ошибка загрузки письма для вложений: {ex.Message}");
                        continue;
                    }

                    var attachments = message.Attachments?.ToList() ?? new List<MimeEntity>();
                    if (!attachments.Any())
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: вложений не найдено (несмотря на BodyStructure).");
                        continue;
                    }

                    bool savedAny = false;
                    foreach (var attach in attachments)
                    {
                        string fname = null;

                        if (attach is MimePart mp)
                        {
                            fname = mp.FileName ?? mp.ContentDisposition?.FileName ?? mp.ContentType?.Name;
                        }
                        else if (attach is MessagePart mpart)
                        {
                            fname = mpart.ContentDisposition?.FileName ?? mpart.ContentType?.Name;
                        }
                        else
                        {
                            fname = attach.ContentDisposition?.FileName ?? attach.ContentType?.Name ?? Guid.NewGuid().ToString();
                        }

                        if (string.IsNullOrEmpty(fname))
                            fname = Guid.NewGuid().ToString() + ".bin";

                        bool isZip = fname.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)
                                     || (attach.ContentType != null &&
                                         attach.ContentType.MediaType.Equals("application", StringComparison.OrdinalIgnoreCase) &&
                                         attach.ContentType.MediaSubtype.Equals("zip", StringComparison.OrdinalIgnoreCase));

                        if (!isZip)
                        {
                            Log($"[NEW] Folder='{folderPath}' UID={uid}: вложение '{fname}' не ZIP, пропускаем.");
                            continue;
                        }

                        string tempDir = Path.Combine(Path.GetTempPath(), "ImapCertWatcher");
                        if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                        string tempFilePath = Path.Combine(tempDir, $"{Guid.NewGuid()}_{MakeSafeFileName(fname)}");

                        try
                        {
                            using (var stream = File.Create(tempFilePath))
                            {
                                if (attach is MimePart mimePart)
                                {
                                    mimePart.Content.DecodeTo(stream);
                                }
                                else if (attach is MessagePart msgPart)
                                {
                                    msgPart.Message.WriteTo(stream);
                                }
                                else
                                {
                                    var txt = attach.ToString();
                                    var b = System.Text.Encoding.UTF8.GetBytes(txt);
                                    stream.Write(b, 0, b.Length);
                                }
                            }

                            try
                            {
                                _db.DeleteArchiveFromDb(certId);
                                Log($"[NEW] CertID={certId}: старые архивы удалены перед сохранением нового.");
                            }
                            catch (Exception exDel)
                            {
                                Log($"[NEW] CertID={certId}: ошибка удаления старых архивов: {exDel.Message}");
                            }

                            bool ok = _db.SaveArchiveToDbTransactional(certId, tempFilePath, Path.GetFileName(tempFilePath));
                            if (ok)
                            {
                                savedAny = true;
                                Log($"[NEW] Folder='{folderPath}' UID={uid}: архив сохранён в БД, certId={certId}, file='{Path.GetFileName(tempFilePath)}'");
                            }
                            else
                            {
                                Log($"[NEW] Folder='{folderPath}' UID={uid}: ошибка SaveArchiveToDbTransactional для certId={certId}");
                            }
                        }
                        catch (Exception exSave)
                        {
                            Log($"[NEW] Folder='{folderPath}' UID={uid}: ошибка сохранения вложения '{fname}': {exSave.Message}");
                        }
                        finally
                        {
                            try { File.Delete(tempFilePath); } catch { }
                        }
                    }

                    if (!savedAny)
                    {
                        Log($"[NEW] Folder='{folderPath}' UID={uid}: ZIP-вложение не найдено или не сохранено.");
                    }
                    else
                    {
                        try
                        {
                            _db.MarkMailProcessed(folderPath, uidStr, "NEW");
                            Log($"[NEW] Folder='{folderPath}' UID={uidStr}: помечено как обработанное (NEW) в PROCESSED_MAILS.");
                        }
                        catch (Exception exMark)
                        {
                            Log($"[NEW] Folder='{folderPath}' UID={uidStr}: ошибка MarkMailProcessed(NEW): {exMark.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[NEW] Ошибка 2-го прохода в папке '{folder.FullName}' UID={uid}: {ex}");
                }
            }

            try { folder.Close(); } catch { }

            // Рекурсивно обрабатываем подпапки
            try
            {
                var subfolders = folder.GetSubfolders(false).ToList();
                Log($"[NEW] Папка='{folder.FullName}': найдено подпапок={subfolders.Count}");
                foreach (var sub in subfolders)
                {
                    ProcessFolderRecursive(client, sub, checkAllMessages, globalResults, ref updatedCount, ref addedCount);
                }
            }
            catch (Exception exSub)
            {
                Log($"[NEW] Ошибка получения подпапок для '{folder.FullName}': {exSub.Message}");
            }
        }

        #region Helpers

        private static string GetMessageBody(MimeMessage message)
        {
            if (message == null) return string.Empty;
            if (!string.IsNullOrEmpty(message.TextBody)) return message.TextBody;
            if (!string.IsNullOrEmpty(message.HtmlBody)) return StripHtml(message.HtmlBody);
            var textParts = message.BodyParts.OfType<TextPart>().ToList();
            if (textParts.Any()) return textParts.First().Text;
            return message.Body?.ToString() ?? string.Empty;
        }

        private static string StripHtml(string html) =>
            string.IsNullOrEmpty(html)
                ? ""
                : Regex.Replace(html, "<.*?>", string.Empty).Replace("&nbsp;", " ");

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

        private static string ExtractCertNumber(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;

            var m = Regex.Match(text, @"Файл\s+сертификата\s*№\s*([0-9A-Fa-f]+)", RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();

            m = Regex.Match(text, @"Сертификат\s*№\s*([0-9A-Fa-f]+)", RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();

            m = Regex.Match(text, @"\b([0-9A-Fa-f]{20,})\b");
            if (m.Success) return m.Groups[1].Value.Trim();

            return null;
        }

        private static (DateTime start, DateTime end) ExtractDates(string text)
        {
            if (string.IsNullOrEmpty(text)) return (DateTime.MinValue, DateTime.MinValue);

            var m = Regex.Match(text,
                @"Срок\s+действия\s+сертификата[:\s]*с\s*([\d\.:\s]+)\s*по\s*([\d\.:\s]+)",
                RegexOptions.IgnoreCase);

            if (m.Success)
            {
                if (DateTime.TryParseExact(
                        m.Groups[1].Value.Trim(),
                        new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime s)
                    && DateTime.TryParseExact(
                        m.Groups[2].Value.Trim(),
                        new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime e))
                {
                    return (s, e);
                }
            }

            var allDates = Regex.Matches(text, @"\b\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}\b");
            if (allDates.Count >= 2)
            {
                if (DateTime.TryParseExact(allDates[0].Value, "dd.MM.yyyy HH:mm:ss",
                        CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime s2)
                    && DateTime.TryParseExact(allDates[1].Value, "dd.MM.yyyy HH:mm:ss",
                        CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime e2))
                {
                    return (s2, e2);
                }
            }

            return (DateTime.MinValue, DateTime.MinValue);
        }

        private static bool BodyStructureHasZip(IMessageSummary summary)
        {
            if (summary == null || summary.Body == null) return false;
            try { return CheckBodyPartForZip(summary.Body); } catch { return false; }
        }

        private static bool CheckBodyPartForZip(MailKit.BodyPart part)
        {
            if (part == null) return false;

            if (part is MailKit.BodyPartBasic basic)
            {
                var fn = basic.FileName ?? basic.ContentType?.Name;
                if (!string.IsNullOrEmpty(fn) && fn.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    return true;

                if (basic.ContentType != null &&
                    basic.ContentType.MediaType.Equals("application", StringComparison.OrdinalIgnoreCase) &&
                    basic.ContentType.MediaSubtype.Equals("zip", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            if (part is MailKit.BodyPartMultipart multi && multi.BodyParts != null)
            {
                foreach (var p in multi.BodyParts)
                    if (CheckBodyPartForZip(p)) return true;
            }

            return false;
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        #endregion
    }
}