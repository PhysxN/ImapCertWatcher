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
using System.Net;
using System.Text;
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

        private static readonly Regex SubjectCertRegex = new Regex(
            @"сертификат\s*№\s*([0-9A-Fa-f][0-9A-Fa-f\s\-]{15,})",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex BodyFioRegex = new Regex(
            @"ФИО(?:\s+владельца(?:\s+сертификата)?)?\s*:\s*([А-ЯЁA-Z][^\r\n\.]{3,120})",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex BodyRevokeDateRegex = new Regex(
            @"Дата\s+отзыва\s+сертификата\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2}:\d{2})?)",
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
            try
            {
                _log("[REVOKE] " + message);
            }
            catch
            {
            }
        }

        public void ProcessRevocations(bool checkAllMessages, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            int processed = 0;
            int applied = 0;

            Log("Старт обработки аннулирований");

            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;

                    for (int attempt = 1; attempt <= 2; attempt++)
                    {
                        try
                        {
                            client.Connect(
                                _settings.MailHost,
                                _settings.MailPort,
                                _settings.MailUseSsl
                                    ? SecureSocketOptions.SslOnConnect
                                    : SecureSocketOptions.StartTlsWhenAvailable);

                            break;
                        }
                        catch (Exception ex)
                        {
                            Log($"IMAP connect attempt {attempt} failed: {ex.Message}");

                            if (attempt == 2)
                                throw;

                            Task.Delay(3000, token).GetAwaiter().GetResult();
                        }
                    }

                    token.ThrowIfCancellationRequested();

                    Log($"IMAP подключение: {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");

                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    token.ThrowIfCancellationRequested();

                    string folderName = (_settings.ImapRevocationsFolder ?? string.Empty).Trim();

                    if (string.IsNullOrWhiteSpace(folderName))
                    {
                        throw new InvalidOperationException(
                            "Не задана папка ImapRevocationsFolder в server.settings.txt.");
                    }

                    IMailFolder folder = client.GetFolder(folderName);

                    if (folder == null)
                    {
                        throw new InvalidOperationException(
                            $"Папка IMAP для аннулирований не найдена: '{folderName}'.");
                    }

                    folder.Open(FolderAccess.ReadOnly);

                    try
                    {
                        long previousLastUid = checkAllMessages
                            ? 0
                            : _db.GetLastUid(folder.FullName + "_REVOKE");

                        var uids = checkAllMessages
                            ? folder.Search(SearchQuery.All)
                            : folder.Search(
                                SearchQuery.Uids(
                                    new UniqueIdRange(
                                        new UniqueId((uint)(previousLastUid + 1)),
                                        UniqueId.MaxValue)));

                        var orderedUids = uids.OrderBy(u => u.Id).ToList();

                        if (!checkAllMessages)
                        {
                            const int MaxPerRun = 200;

                            if (orderedUids.Count > MaxPerRun)
                            {
                                orderedUids = orderedUids
                                    .Take(MaxPerRun)
                                    .ToList();

                                Log($"Ограничение: обрабатываем первые {MaxPerRun} новых писем за прогон. Остальные будут обработаны в следующих запусках");
                            }
                        }

                        long lastCompletedUid = previousLastUid;

                        Log($"Найдено писем: {orderedUids.Count}");

                        foreach (var uid in orderedUids)
                        {
                            token.ThrowIfCancellationRequested();

                            string uidStr = uid.ToString();
                            bool finalized = false;

                            if (!checkAllMessages &&
                                _db.IsMailProcessed(folder.FullName, uidStr, "REVOKE"))
                            {
                                Log($"UID={uidStr}: уже обработано, пропуск");
                                finalized = true;
                            }
                            else
                            {
                                try
                                {
                                    processed++;

                                    var message = folder.GetMessage(uid);

                                    if (!TryExtractStrictRevocationData(
                                            message,
                                            out string certNumber,
                                            out string fio,
                                            out DateTime? revokeDate,
                                            out string skipReason))
                                    {
                                        if (!MarkProcessed(folder, uidStr))
                                        {
                                            Log($"UID={uidStr}: пропуск, {skipReason}, но письмо не удалось пометить обработанным");
                                            finalized = false;
                                        }
                                        else
                                        {
                                            Log($"UID={uidStr}: пропуск, {skipReason}, помечено обработанным");
                                            finalized = true;
                                        }
                                    }
                                    else
                                    {
                                        Log($"UID={uidStr}: extracted cert={certNumber}, fio={fio}, revokeDate={revokeDate:dd.MM.yyyy HH:mm:ss}");

                                        string dateText = revokeDate.HasValue
                                            ? revokeDate.Value.ToString("dd.MM.yyyy HH:mm:ss")
                                            : "<без даты>";

                                        bool ok = _db.FindAndMarkAsRevokedByCertAndFio(
                                            certNumber,
                                            fio,
                                            folder.FullName,
                                            revokeDate?.ToString("dd.MM.yyyy HH:mm:ss"),
                                            out string resultMessage);

                                        if (ok)
                                            applied++;

                                        Log($"UID={uidStr}: {resultMessage}. Сертификат: {certNumber}, ФИО: {fio}, дата {dateText}");

                                        if (!MarkProcessed(folder, uidStr))
                                        {
                                            Log($"UID={uidStr}: письмо обработано, но не удалось пометить его как обработанное");
                                            finalized = false;
                                        }
                                        else
                                        {
                                            finalized = true;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"UID={uidStr}: ошибка обработки письма: {ex.Message}");
                                    finalized = false;
                                }
                            }

                            if (!checkAllMessages)
                            {
                                if (finalized)
                                {
                                    lastCompletedUid = uid.Id;
                                }
                                else
                                {
                                    Log($"UID={uidStr}: остановка прохода, чтобы не перепрыгнуть проблемное письмо");
                                    break;
                                }
                            }
                        }

                        if (!checkAllMessages && lastCompletedUid > previousLastUid)
                        {
                            _db.UpdateLastUid(folder.FullName + "_REVOKE", lastCompletedUid);
                            Log($"LAST_UID обновлён для '{folder.FullName}_REVOKE' -> {lastCompletedUid}");
                        }
                    }
                    finally
                    {
                        try { folder.Close(); } catch { }

                        if (client.IsConnected)
                            client.Disconnect(false);
                    }
                }

                Log($"ИТОГО: обработано={processed}, применено={applied}");
                Log("Завершение обработки аннулирований");
            }
            catch (OperationCanceledException)
            {
                Log("Обработка аннулирований отменена");
                throw;
            }
            catch (Exception ex)
            {
                Log("Ошибка: " + ex.Message);
                throw;
            }
        }

        private bool MarkProcessed(IMailFolder folder, string uid)
        {
            try
            {
                _db.MarkMailProcessed(folder.FullName, uid, "REVOKE");
                return true;
            }
            catch (Exception ex)
            {
                Log($"Ошибка MarkMailProcessed: {ex.Message}");
                return false;
            }
        }

        public Task CheckRevocationsFastAsync(CancellationToken token)
        {
            ProcessRevocations(false, token);
            return Task.CompletedTask;
        }


        private static string StripHtml(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return string.Empty;

            return Regex.Replace(html, "<.*?>", " ");
        }

        private static string NormalizeForMatch(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var s = text.Replace('\u00A0', ' ');
            s = WebUtility.HtmlDecode(s);
            s = Regex.Replace(s, @"\s+", " ");
            return s.Trim().ToLowerInvariant();
        }

        private static bool ContainsRevokePhrase(string text)
        {
            var s = NormalizeForMatch(text);

            return s.Contains("аннулир") ||
                   s.Contains("прекратил действие") ||
                   s.Contains("прекращен") ||
                   s.Contains("прекращён") ||
                   s.Contains("отозван");
        }


        private static bool IsLikelyFio(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            var normalized = value.Trim();

            if (normalized.Length < 5 || normalized.Length > 120)
                return false;

            var lower = normalized.ToLowerInvariant();

            if (lower.Contains("сведения") ||
                lower.Contains("сертификат") ||
                lower.Contains("организац") ||
                lower.Contains("дата") ||
                lower.Contains("причина") ||
                lower.Contains("казнач"))
            {
                return false;
            }

            var parts = normalized
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length < 2 || parts.Length > 4)
                return false;

            foreach (var part in parts)
            {
                if (!Regex.IsMatch(part, @"^[А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z\-]+$"))
                    return false;
            }

            return true;
        }

        private static string GetBodyText(MimeMessage msg)
        {
            var sb = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(msg.TextBody))
                sb.AppendLine(msg.TextBody);

            if (!string.IsNullOrWhiteSpace(msg.HtmlBody))
                sb.AppendLine(WebUtility.HtmlDecode(StripHtml(msg.HtmlBody)));

            var result = sb.ToString();
            result = result.Replace('\u00A0', ' ');
            result = Regex.Replace(result, @"[ \t]+", " ");

            return result;
        }

        private static bool TryExtractStrictRevocationData(
            MimeMessage message,
            out string certNumber,
            out string fio,
            out DateTime? revokeDate,
            out string skipReason)
        {
            certNumber = null;
            fio = null;
            revokeDate = null;
            skipReason = null;

            if (message == null)
            {
                skipReason = "message == null";
                return false;
            }

            string subjectRaw = message.Subject ?? "";
            string bodyRaw = GetBodyText(message);

            if (string.IsNullOrWhiteSpace(subjectRaw))
            {
                skipReason = "пустая тема";
                return false;
            }

            if (string.IsNullOrWhiteSpace(bodyRaw))
            {
                skipReason = "пустое тело";
                return false;
            }

            if (!ContainsRevokePhrase(subjectRaw))
            {
                skipReason = "в теме нет обязательной фразы об аннулировании";
                return false;
            }

            var certMatch = SubjectCertRegex.Match(subjectRaw);
            if (!certMatch.Success)
            {
                skipReason = "номер сертификата не найден в теме";
                return false;
            }

            certNumber = DbHelper.NormalizeCertNumber(certMatch.Groups[1].Value);
            if (string.IsNullOrWhiteSpace(certNumber))
            {
                skipReason = "номер сертификата пуст после нормализации";
                return false;
            }

            var fioMatch = BodyFioRegex.Match(bodyRaw);
            if (!fioMatch.Success)
            {
                skipReason = "ФИО не найдено в теле";
                return false;
            }

            fio = fioMatch.Groups[1].Value.Trim();
            fio = Regex.Replace(fio, @"\s+", " ");
            fio = fio.Trim('.', ';', ',', ':');

            if (string.IsNullOrWhiteSpace(fio))
            {
                skipReason = "ФИО пустое";
                return false;
            }

            if (!IsLikelyFio(fio))
            {
                skipReason = "извлечённое значение не похоже на ФИО";
                return false;
            }

            var dateMatch = BodyRevokeDateRegex.Match(bodyRaw);

            if (dateMatch.Success)
            {
                var rawDate = dateMatch.Groups[1].Value.Trim();

                if (DateTime.TryParseExact(
                        rawDate,
                        "dd.MM.yyyy HH:mm:ss",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime parsedDateTime))
                {
                    revokeDate = parsedDateTime;
                    return true;
                }

                if (DateTime.TryParseExact(
                        rawDate,
                        "dd.MM.yyyy",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime parsedDate))
                {
                    revokeDate = parsedDate;
                    return true;
                }
            }

            if (message.Date != DateTimeOffset.MinValue)
            {
                revokeDate = message.Date.LocalDateTime;
                return true;
            }

            revokeDate = DateTime.Now;
            return true;
        }
    }
}