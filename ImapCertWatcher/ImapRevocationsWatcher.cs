using ImapCertWatcher.Data;
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

        public ImapRevocationsWatcher(AppSettings settings, DbHelper db, Action<string> log = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
        }

        private void Log(string msg) => _log?.Invoke(msg);

        /// <summary>
        /// Основной метод:
        ///  - подключается к IMAP;
        ///  - находит корневую папку _settings.ImapFolder;
        ///  - рекурсивно обходит её и все подпапки;
        ///  - по теме "Сертификат № ... аннулирован (или прекратил действие)" + телу письма
        ///    ищет сертификат в БД и помечает его аннулированным.
        /// </summary>
        public int ProcessRevocations(bool checkAllMessages = false)
        {
            int applied = 0; // количество успешно применённых аннулирований

            Log($"=== ImapRevocationsWatcher: старт проверки (checkAllMessages={checkAllMessages}) ===");

            // Строгая тема:
            // "Сертификат № 008954001058EE95915C0F731E29907571 аннулирован"
            // или
            // "Сертификат № 008954001058EE95915C0F731E29907571 аннулирован или прекратил действие"
            var pattern = @"^Сертификат\s*№\s*([0-9A-Fa-f]+)\s+аннулирован(?:\s+или\s+прекратил\s+действие)?\.?$";
            var subjectRegex = new Regex(pattern, RegexOptions.IgnoreCase);

            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    Log($"[Connect] Подключение к {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");
                    if (_settings.MailUseSsl)
                        client.Connect(_settings.MailHost, _settings.MailPort, SecureSocketOptions.SslOnConnect);
                    else
                        client.Connect(_settings.MailHost, _settings.MailPort, SecureSocketOptions.StartTlsWhenAvailable);

                    Log($"[Auth] Аутентификация как {_settings.MailLogin}");
                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    // Стартуем из той же папки, что и новые сертификаты
                    var rootFolder = ResolveRootFolder(client, _settings.ImapFolder);
                    if (rootFolder == null)
                    {
                        Log($"[Root] Не удалось найти папку '{_settings.ImapFolder}' для аннулирований.");
                        client.Disconnect(true);
                        return applied;
                    }

                    Log($"[Root] Старт обработки аннулирований из папки '{rootFolder.FullName}'");

                    ProcessFolderRecursive(rootFolder, subjectRegex, checkAllMessages, ref applied);

                    Log("[Disconnect] Отключение от IMAP-сервера");
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Log($"[Revoke][ERROR] ProcessRevocations error: {ex}");
            }

            Log($"=== ImapRevocationsWatcher: завершено. Всего применено аннулирований: {applied} ===");
            return applied;
        }

        #region Поиск корневой папки и рекурсивный обход

        private IMailFolder ResolveRootFolder(ImapClient client, string folderName)
        {
            IMailFolder folder = null;

            try
            {
                folder = client.GetFolder(folderName);
                if (folder != null)
                {
                    Log($"[ResolveRootFolder] (Revoke) Найдена папка напрямую: '{folder.FullName}'");
                    return folder;
                }
            }
            catch (Exception ex)
            {
                Log($"[ResolveRootFolder] (Revoke) Не удалось открыть папку '{folderName}' напрямую: {ex.Message}");
            }

            try
            {
                var ns = client.PersonalNamespaces.FirstOrDefault();
                if (ns != null)
                {
                    var root = client.GetFolder(ns);
                    Log($"[ResolveRootFolder] (Revoke) Поиск папки '{folderName}' рекурсивно под '{root.FullName}'");
                    var found = FindFolderRecursive(root, folderName);
                    if (found != null)
                    {
                        Log($"[ResolveRootFolder] (Revoke) Найдена папка рекурсивно: '{found.FullName}'");
                        return found;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[ResolveRootFolder] (Revoke) Ошибка при рекурсивном поиске: {ex.Message}");
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
                Log($"[FindFolderRecursive] (Revoke) Ошибка при обходе подпапок '{current.FullName}': {ex.Message}");
            }

            return null;
        }

        private void ProcessFolderRecursive(
            IMailFolder folder,
            Regex subjectRegex,
            bool checkAllMessages,
            ref int applied)
        {
            if (folder == null) return;

            Log($"[Revoke][Folder] === Обработка папки '{folder.FullName}' ===");

            ProcessSingleFolder(folder, subjectRegex, checkAllMessages, ref applied);

            try
            {
                var subs = folder.GetSubfolders(false).ToList();
                Log($"[Revoke][Folder] Папка '{folder.FullName}' имеет подпапок: {subs.Count}");
                foreach (var sub in subs)
                {
                    ProcessFolderRecursive(sub, subjectRegex, checkAllMessages, ref applied);
                }
            }
            catch (Exception ex)
            {
                Log($"[Revoke][Folder] Ошибка при получении подпапок '{folder.FullName}': {ex.Message}");
            }
        }

        #endregion

        #region Обработка одной папки (без ZIP!)

        /// <summary>
        /// Обработка одной IMAP-папки:
        ///  - ищем письма с темой "Сертификат № ... аннулирован (или прекратил действие)";
        ///  - по тексту письма вытаскиваем ФИО и дату отзыва;
        ///  - вызываем FindAndMarkAsRevokedByCertNumber.
        /// </summary>
        private void ProcessSingleFolder(
            IMailFolder folder,
            Regex subjectRegex,
            bool checkAllMessages,
            ref int applied)
        {
            try
            {
                folder.Open(FolderAccess.ReadOnly);
            }
            catch (Exception exOpen)
            {
                Log($"[Revoke][Folder] Не удалось открыть папку '{folder.FullName}': {exOpen.Message}");
                return;
            }

            try
            {
                // Собираем UID писем
                IList<UniqueId> uids;
                if (checkAllMessages)
                {
                    uids = folder.Search(SearchQuery.All);
                }
                else
                {
                    // для аннулирований логичнее смотреть более длинный период, чем 10 дней
                    var dateLimit = DateTime.Now.AddDays(-365);
                    uids = folder.Search(SearchQuery.DeliveredAfter(dateLimit));
                }

                Log($"[Revoke][Search] Папка='{folder.FullName}', checkAll={checkAllMessages}, найдено UID: {uids?.Count ?? 0}");

                var summaries = new List<IMessageSummary>();
                if (uids != null && uids.Count > 0)
                {
                    summaries.AddRange(
                        folder.Fetch(
                            uids,
                            MessageSummaryItems.Envelope |
                            MessageSummaryItems.InternalDate |
                            MessageSummaryItems.UniqueId
                        )
                    );
                }

                Log($"[Revoke][Fetch] Папка='{folder.FullName}', summaries count={summaries.Count}");

                int preview = 0;
                foreach (var s in summaries)
                {
                    var subj = s.Envelope?.Subject ?? "(no subject)";
                    var dt = s.InternalDate?.ToString() ?? "(no date)";
                    Log($"[Revoke][Preview] Folder='{folder.FullName}' UID={s.UniqueId} Date={dt} Subject='{subj}'");
                    if (++preview >= 30) break;
                }

                foreach (var summary in summaries)
                {
                    var uid = summary.UniqueId;
                    string subject = summary.Envelope?.Subject ?? "";
                    if (string.IsNullOrWhiteSpace(subject))
                    {
                        Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: пустая тема, пропуск.");
                        continue;
                    }

                    var m = subjectRegex.Match(subject);
                    if (!m.Success)
                    {
                        Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: тема не подходит под шаблон аннулирования, пропуск. Subject='{subject}'");
                        continue;
                    }

                    string certFromSubject = m.Groups[1].Value.Trim();
                    Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: тема подходит, certFromSubject='{certFromSubject}'");

                    // Загружаем полное письмо, чтобы прочитать тело
                    MimeMessage message;
                    try
                    {
                        message = folder.GetMessage(uid);
                        Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: письмо загружено полностью.");
                    }
                    catch (Exception exMsg)
                    {
                        Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: ошибка загрузки письма: {exMsg.Message}");
                        continue;
                    }

                    string bodyText = GetMessageBody(message);
                    Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: длина текста письма={bodyText?.Length ?? 0}");

                    // ФИО из тела
                    string fio = ExtractFio(bodyText) ?? "";
                    // Дата отзыва из тела
                    DateTime? revokeDt = ExtractRevokeDate(bodyText);
                    string revokeDateShort = revokeDt?.ToString("dd.MM.yyyy") ?? "";

                    Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: parsed fio='{fio}', revokeDate='{revokeDateShort}'");

                    if (string.IsNullOrEmpty(certFromSubject) && string.IsNullOrEmpty(fio))
                    {
                        Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: ни номера из темы, ни ФИО — пропуск.");
                        continue;
                    }

                    // На всякий случай: попробуем взять номер сертификата из тела, если в теме не удалось (теоретический случай)
                    if (string.IsNullOrEmpty(certFromSubject))
                    {
                        string certFromBody = ExtractCertNumber(bodyText);
                        if (!string.IsNullOrEmpty(certFromBody))
                        {
                            certFromSubject = certFromBody;
                            Log($"[Revoke] Folder='{folder.FullName}' UID={uid}: номер сертификата взят из тела: '{certFromSubject}'");
                        }
                    }

                    // Наконец — помечаем сертификат в базе
                    try
                    {
                        bool ok = _db.FindAndMarkAsRevokedByCertNumber(
                            certFromSubject,
                            fio,
                            folder.FullName,
                            revokeDateShort
                        );

                        if (ok)
                        {
                            applied++;
                            Log($"[Revoke][APPLY] Folder='{folder.FullName}' UID={uid}: аннулирован сертификат: {certFromSubject} — {fio}, дата='{revokeDateShort}'");
                        }
                        else
                        {
                            Log($"[Revoke][MISS] Folder='{folder.FullName}' UID={uid}: НЕ найден сертификат для аннулирования: {certFromSubject} — {fio}, дата='{revokeDateShort}'");
                        }
                    }
                    catch (Exception exDb)
                    {
                        Log($"[Revoke][DB] Folder='{folder.FullName}' UID={uid}: ошибка FindAndMarkAsRevokedByCertNumber: {exDb.Message}");
                    }
                }
            }
            finally
            {
                try { folder.Close(); } catch { }
            }
        }

        #endregion

        #region Helpers (парсинг текста письма)

        private static string GetMessageBody(MimeMessage message)
        {
            if (message == null) return string.Empty;
            if (!string.IsNullOrEmpty(message.TextBody)) return message.TextBody;
            if (!string.IsNullOrEmpty(message.HtmlBody)) return Regex.Replace(message.HtmlBody, "<.*?>", string.Empty);

            var textParts = message.BodyParts.OfType<TextPart>().ToList();
            if (textParts.Any()) return textParts.First().Text;

            return message.Body?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Ищем "ФИО: Кулькова Алина Алексеевна."
        /// или похожие конструкции.
        /// </summary>
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

        /// <summary>
        /// Номер сертификата из текста письма (на случай fallback).
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

            m = Regex.Match(text, @"\b([0-9A-Fa-f]{20,})\b");
            if (m.Success) return m.Groups[1].Value.Trim();

            return null;
        }

        /// <summary>
        /// Ищем "Дата отзыва сертификата: 17.04.2025 10:43:29 (по московскому времени)."
        /// Возвращаем DateTime? (локальное время).
        /// </summary>
        private static DateTime? ExtractRevokeDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            // Основной шаблон
            var m = Regex.Match(text,
                @"Дата\s+отзыва\s+сертификата[:\s]*([\d\.:\s]+)",
                RegexOptions.IgnoreCase);
            if (m.Success)
            {
                var raw = m.Groups[1].Value.Trim();
                if (DateTime.TryParseExact(
                        raw,
                        new[] { "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime dt))
                {
                    return dt;
                }
            }

            // Запасной вариант: первая дата после фразы "Дата отзыва"
            m = Regex.Match(text,
                @"Дата\s+отзыва.*?([0-3]?\d\.[0-1]?\d\.\d{4})",
                RegexOptions.IgnoreCase);
            if (m.Success)
            {
                var raw = m.Groups[1].Value.Trim();
                if (DateTime.TryParseExact(
                        raw,
                        new[] { "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal,
                        out DateTime dt))
                {
                    return dt;
                }
            }

            return null;
        }

        #endregion
    }
}
