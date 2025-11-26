using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ImapCertWatcher.Utils;
using ImapCertWatcher.Data;
using ImapCertWatcher.Models;

namespace ImapCertWatcher.Services
{
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

        private void Log(string msg) => _log?.Invoke(msg);

        /// <summary>
        /// Основной метод:
        ///  - подключается к IMAP,
        ///  - находит корневую папку _settings.ImapFolder,
        ///  - рекурсивно обходит её и все подпапки,
        ///  - в каждой папке:
        ///      1) находит письма с корректной темой;
        ///      2) извлекает ФИО/номер/даты и делает InsertOrUpdateAndGetId;
        ///      3) если есть ZIP – скачивает и сохраняет в БД транзакционно.
        /// </summary>
        public (List<CertEntry> processedEntries, int updatedCount, int addedCount)
            ProcessNewCertificates(bool checkAllMessages = false)
        {
            var results = new List<CertEntry>();
            int updatedCount = 0;
            int addedCount = 0;

            Log($"=== ImapNewCertificatesWatcher: старт проверки (checkAllMessages={checkAllMessages}) ===");

            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    Log($"[Connect] Подключение к {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");
                    if (_settings.MailUseSsl)
                        client.Connect(_settings.MailHost, _settings.MailPort, MailKit.Security.SecureSocketOptions.SslOnConnect);
                    else
                        client.Connect(_settings.MailHost, _settings.MailPort, MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);

                    Log($"[Auth] Аутентификация как {_settings.MailLogin}");
                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    // Находим корневую папку (ту, что указана в настройках)
                    var rootFolder = ResolveRootFolder(client, _settings.ImapFolder);
                    if (rootFolder == null)
                    {
                        Log($"[Root] Не удалось найти папку '{_settings.ImapFolder}' ни напрямую, ни через личный namespace.");
                        client.Disconnect(true);
                        return (results, updatedCount, addedCount);
                    }

                    Log($"[Root] Старт обработки с папки '{rootFolder.FullName}'");

                    // Рекурсивный обход: сама папка + все подпапки
                    ProcessFolderRecursive(rootFolder, checkAllMessages, results, ref updatedCount, ref addedCount);

                    Log("[Disconnect] Отключение от IMAP-сервера");
                    client.Disconnect(true);
                }
            }
            catch (Exception exTop)
            {
                Log($"[ERROR] Ошибка ImapNewCertificatesWatcher: {exTop}");
            }

            Log($"=== ImapNewCertificatesWatcher: завершено. Всего обработано записей={results.Count}, добавлено={addedCount}, обновлено={updatedCount} ===");
            return (results, updatedCount, addedCount);
        }

        /// <summary>
        /// Пытается найти корневую папку по имени:
        ///  1) Прямой вызов client.GetFolder("Имя");
        ///  2) Рекурсивный поиск по личному namespace (PersonalNamespaces).
        /// </summary>
        private IMailFolder ResolveRootFolder(ImapClient client, string folderName)
        {
            IMailFolder folder = null;

            try
            {
                folder = client.GetFolder(folderName);
                if (folder != null)
                {
                    Log($"[ResolveRootFolder] Найдена папка напрямую: '{folder.FullName}'");
                    return folder;
                }
            }
            catch (Exception ex)
            {
                Log($"[ResolveRootFolder] Не удалось открыть папку '{folderName}' напрямую: {ex.Message}");
            }

            try
            {
                var ns = client.PersonalNamespaces.FirstOrDefault();
                if (ns != null)
                {
                    var root = client.GetFolder(ns);
                    Log($"[ResolveRootFolder] Поиск папки '{folderName}' рекурсивно под '{root.FullName}'");
                    var found = FindFolderRecursive(root, folderName);
                    if (found != null)
                    {
                        Log($"[ResolveRootFolder] Найдена папка рекурсивно: '{found.FullName}'");
                        return found;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[ResolveRootFolder] Ошибка при рекурсивном поиске: {ex.Message}");
            }

            return null;
        }

        private IMailFolder FindFolderRecursive(IMailFolder current, string targetName)
        {
            if (current.FullName.Equals(targetName, StringComparison.OrdinalIgnoreCase) ||
                current.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
            {
                return current;
            }

            try
            {
                foreach (var sub in current.GetSubfolders(false))
                {
                    var found = FindFolderRecursive(sub, targetName);
                    if (found != null) return found;
                }
            }
            catch (Exception ex)
            {
                Log($"[FindFolderRecursive] Ошибка при обходе подпапок '{current.FullName}': {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Рекурсивно обходит папку и все её подпапки.
        /// Для каждой папки вызывает ProcessSingleFolder.
        /// </summary>
        private void ProcessFolderRecursive(
            IMailFolder folder,
            bool checkAllMessages,
            List<CertEntry> results,
            ref int updatedCount,
            ref int addedCount)
        {
            if (folder == null) return;

            Log($"[Folder] === Обработка папки '{folder.FullName}' ===");

            // Сначала обрабатываем саму папку
            ProcessSingleFolder(folder, checkAllMessages, results, ref updatedCount, ref addedCount);

            // Потом рекурсивно все подпапки
            try
            {
                var subs = folder.GetSubfolders(false).ToList();
                Log($"[Folder] Папка '{folder.FullName}' имеет подпапок: {subs.Count}");
                foreach (var sub in subs)
                {
                    ProcessFolderRecursive(sub, checkAllMessages, results, ref updatedCount, ref addedCount);
                }
            }
            catch (Exception ex)
            {
                Log($"[Folder] Ошибка при получении подпапок '{folder.FullName}': {ex.Message}");
            }
        }

        /// <summary>
        /// Полная логика обработки одной IMAP-папки:
        ///  - поиск UID сообщений;
        ///  - Fetch summary (Envelope + BodyStructure + InternalDate);
        ///  - первый проход: парсинг, InsertOrUpdateAndGetId, отбор кандидатов на ZIP;
        ///  - второй проход: загрузка ZIP-вложений и сохранение в БД.
        /// </summary>
        private void ProcessSingleFolder(
            IMailFolder folder,
            bool checkAllMessages,
            List<CertEntry> results,
            ref int updatedCount,
            ref int addedCount)
        {
            try
            {
                folder.Open(FolderAccess.ReadOnly);
            }
            catch (Exception exOpen)
            {
                Log($"[Folder] Не удалось открыть папку '{folder.FullName}': {exOpen.Message}");
                return;
            }

            try
            {
                // Получаем UIDs (в зависимости от режима)
                IList<UniqueId> uids;
                if (checkAllMessages)
                {
                    uids = folder.Search(SearchQuery.All);
                }
                else
                {
                    // Можно изменить окно при необходимости
                    var dateLimit = DateTime.Now.AddDays(-10);
                    uids = folder.Search(SearchQuery.DeliveredAfter(dateLimit));
                }

                Log($"[Search] Папка='{folder.FullName}', checkAll={checkAllMessages}, найдено UID: {uids?.Count ?? 0}");

                // Получаем summary для быстрого определения структуры/наличия вложений
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
                        Log($"[Fetch] Папка='{folder.FullName}', summaries count={summaries.Count}");

                        // Превью первых 30 писем (UID, дата, тема)
                        int preview = 0;
                        foreach (var s in summaries)
                        {
                            var subj = s.Envelope?.Subject ?? "(no subject)";
                            var dt = s.InternalDate?.ToString() ?? "(no date)";
                            Log($"[Preview] Folder='{folder.FullName}' UID={s.UniqueId} Date={dt} Subject='{subj}'");
                            if (++preview >= 30) break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[Fetch] Ошибка при Fetch summaries в '{folder.FullName}': {ex.Message}");
                    }
                }
                else
                {
                    Log($"[Fetch] Папка='{folder.FullName}': UIDs нет, summaries не запрошены.");
                }

                // Регекс: тема строго начинается с префикса и затем только номер (без лишнего текста)
                var subjectStrictRegex = new Regex(
                    @"^" + Regex.Escape(_subjectPrefix) + @"\s*([0-9A-Fa-f]+)\s*$",
                    RegexOptions.Compiled
                );

                // Накапливаем кандидатов для второго прохода (UID + certId)
                var zipCandidates = new List<(UniqueId uid, int certId)>();

                // ---------- ПЕРВЫЙ ПРОХОД ----------
                foreach (var s in summaries)
                {
                    try
                    {
                        var subject = s.Envelope?.Subject ?? string.Empty;
                        var uid = s.UniqueId;

                        // Быстрая фильтрация: тема должна строго совпадать по шаблону
                        var m = subjectStrictRegex.Match(subject);
                        if (!m.Success)
                        {
                            Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: тема не соответствует шаблону, пропуск. Subject='{subject}'");
                            continue;
                        }

                        // Загружаем полное письмо (нужно для парсинга тела и вложений)
                        MimeMessage message;
                        try
                        {
                            message = folder.GetMessage(uid);
                            Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: письмо загружено полностью.");
                        }
                        catch (Exception ex)
                        {
                            Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: ошибка загрузки письма: {ex.Message}");
                            continue;
                        }

                        var bodyText = GetMessageBody(message) ?? string.Empty;
                        var fromAddress = message.From.Mailboxes.FirstOrDefault()?.Address ?? string.Empty;

                        // Извлекаем ФИО, номер сертификата, даты
                        var fio = ExtractFio(bodyText);
                        var certFromSubject = m.Groups[1].Value.Trim();
                        var certNumber = !string.IsNullOrEmpty(certFromSubject)
                            ? certFromSubject
                            : ExtractCertNumber(bodyText);
                        var dates = ExtractDates(bodyText);

                        Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: fio='{fio}', cert='{certNumber}', start={dates.start}, end={dates.end}");

                        if (string.IsNullOrWhiteSpace(fio)
                            || string.IsNullOrWhiteSpace(certNumber)
                            || dates.start == DateTime.MinValue
                            || dates.end == DateTime.MinValue)
                        {
                            Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: недостаточно данных (FIO/Cert/Dates) — пропускаем.");
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
                            FolderPath = folder.FullName,
                            MailUid = uid.ToString(),
                            MessageDate = s.InternalDate?.UtcDateTime ?? message.Date.UtcDateTime,
                            Subject = subject
                        };

                        // Сохраняем/обновляем запись атомарно и получаем certId
                        try
                        {
                            var (wasUpdated, wasAdded, certId) = _db.InsertOrUpdateAndGetId(entry);
                            if (wasAdded) addedCount++;
                            if (wasUpdated) updatedCount++;

                            Log($"[DB] Folder='{folder.FullName}' UID={uid}: InsertOrUpdateAndGetId -> certId={certId}, added={wasAdded}, updated={wasUpdated}");

                            // Для результатов UI — кладём запись (без привязанного архива)
                            results.Add(entry);

                            // Проверка: есть ли ZIP во summary (быстро, без загрузки вложений)
                            bool hasZipInSummary = BodyStructureHasZip(s);
                            Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: BodyStructureHasZip={hasZipInSummary}");

                            if (hasZipInSummary && certId > 0)
                            {
                                zipCandidates.Add((uid, certId));
                                Log($"[FirstPass] Folder='{folder.FullName}' UID={uid}: помечено для скачивания архива (certId={certId})");
                            }
                        }
                        catch (Exception exDb)
                        {
                            Log($"[DB] Folder='{folder.FullName}' UID={uid}: ошибка InsertOrUpdateAndGetId: {exDb.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[FirstPass] Folder='{folder.FullName}': ошибка при обработке summary: {ex.Message}");
                    }
                } // foreach summaries

                Log($"[FirstPass] Папка='{folder.FullName}': завершено. Обработано={summaries.Count}, zipCandidates={zipCandidates.Count}");

                // ---------- ВТОРОЙ ПРОХОД: скачиваем ZIP-вложения ----------
                foreach (var (uid, certId) in zipCandidates)
                {
                    try
                    {
                        MimeMessage message;
                        try
                        {
                            message = folder.GetMessage(uid);
                            Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: письмо загружено для скачивания вложений.");
                        }
                        catch (Exception ex)
                        {
                            Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: ошибка загрузки письма при скачивании вложений: {ex.Message}");
                            continue;
                        }

                        var attachments = message.Attachments?.ToList() ?? new List<MimeEntity>();
                        if (!attachments.Any())
                        {
                            Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: вложений не найдено (несмотря на summary).");
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
                                // fallback
                                fname = attach.ContentDisposition?.FileName ?? attach.ContentType?.Name ?? Guid.NewGuid().ToString();
                            }

                            if (string.IsNullOrEmpty(fname))
                                fname = Guid.NewGuid().ToString() + ".bin";

                            bool isZip =
                                fname.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ||
                                (attach.ContentType != null &&
                                 attach.ContentType.MediaType.Equals("application", StringComparison.OrdinalIgnoreCase) &&
                                 attach.ContentType.MediaSubtype.Equals("zip", StringComparison.OrdinalIgnoreCase));

                            if (!isZip)
                            {
                                Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: вложение '{fname}' не ZIP, пропуск.");
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
                                        // fallback: raw текст
                                        var txt = attach.ToString();
                                        var b = System.Text.Encoding.UTF8.GetBytes(txt);
                                        stream.Write(b, 0, b.Length);
                                    }
                                }

                                // Транзакционное сохранение архива и обновление ARCHIVE_PATH
                                bool ok = _db.SaveArchiveToDbTransactional(certId, tempFilePath, Path.GetFileName(tempFilePath));
                                if (ok)
                                {
                                    savedAny = true;
                                    Log($"[SecondPass][DB] Folder='{folder.FullName}' UID={uid}: архив сохранён для certId={certId}: {Path.GetFileName(tempFilePath)}");
                                }
                                else
                                {
                                    Log($"[SecondPass][DB] Folder='{folder.FullName}' UID={uid}: ошибка сохранения архива в БД для certId={certId}");
                                }
                            }
                            catch (Exception exSave)
                            {
                                Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: ошибка сохранения вложения '{fname}': {exSave.Message}");
                            }
                            finally
                            {
                                try { File.Delete(tempFilePath); } catch { }
                            }
                        } // foreach attachments

                        if (!savedAny)
                        {
                            Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: zip-вложение не найдено или не сохранено.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[SecondPass] Folder='{folder.FullName}' UID={uid}: ошибка во 2-м проходе: {ex.Message}");
                    }
                } // foreach zipCandidates
            }
            finally
            {
                try { folder.Close(); } catch { }
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
            if (string.IsNullOrWhiteSpace(text))
                return null;

            // Ищем строку, где упоминается "ФИО:"
            // Берём всё, что после "ФИО:" до конца строки
            var m = Regex.Match(
                text,
                @"^.*ФИО[:\s]*(.+)$",
                RegexOptions.IgnoreCase | RegexOptions.Multiline);

            if (!m.Success)
                return null;

            var fioRaw = m.Groups[1].Value.Trim();

            // Отрезаем всё после первой точки (обычно ФИО заканчивается точкой)
            int dotIdx = fioRaw.IndexOf('.');
            if (dotIdx >= 0)
                fioRaw = fioRaw.Substring(0, dotIdx);

            // На всякий случай — отрезаем хвост, если в ту же строку попало "Срок действия ..."
            var cutMarkers = new[]
            {
        "Срок действия сертификата",
        "Срок действия",
        "Срок действия сертифката",
        "Срок"
    };

            foreach (var marker in cutMarkers)
            {
                int idx = fioRaw.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (idx > 0)
                {
                    fioRaw = fioRaw.Substring(0, idx);
                    break;
                }
            }

            // Чистим лишние знаки
            fioRaw = fioRaw.Trim(' ', '\t', ',', ';', ':');

            // Оставляем только «словные» части (буквы/дефис), как правило Фамилия Имя Отчество
            var parts = Regex.Matches(fioRaw, @"[\p{L}\-]+")
                             .Cast<Match>()
                             .Select(mm => mm.Value)
                             .ToList();

            if (parts.Count == 0)
                return null;

            if (parts.Count > 3)
                parts = parts.Take(3).ToList();

            var fio = string.Join(" ", parts).Trim();
            return string.IsNullOrWhiteSpace(fio) ? null : fio;
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

            var m = Regex.Match(text, @"Срок\s+действия\s+сертификата[:\s]*с\s*([\d\.:\s]+)\s*по\s*([\d\.:\s]+)", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                if (DateTime.TryParseExact(m.Groups[1].Value.Trim(),
                        new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime s)
                    && DateTime.TryParseExact(m.Groups[2].Value.Trim(),
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
