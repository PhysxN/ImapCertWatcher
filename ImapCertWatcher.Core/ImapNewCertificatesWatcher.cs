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
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ImapCertWatcher.Services
{
    /// <summary>
    /// Обработка новых сертификатов:
    /// - обходит IMAP-папку и подпапки
    /// - ищет письма с ZIP-вложениями
    /// - извлекает CER из ZIP
    /// - читает данные из CER
    /// - сохраняет CER в БД
    /// - помечает письмо как обработанное (PROCESSED_MAILS, kind=NEW)
    /// </summary>
    public class ImapNewCertificatesWatcher
    {
        private readonly ServerSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        public ImapNewCertificatesWatcher(
            ServerSettings settings,
            DbHelper db,
            Action<string> log = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
        }

        private void Log(string msg)
        {
            try
            {
                _log("[NEW] " + msg);
            }
            catch
            {
            }
        }

        // =====================================================================
        // PUBLIC API
        // =====================================================================

        private void CleanupTempFolder()
        {
            try
            {
                string dir = Path.Combine(Path.GetTempPath(), "ImapCertWatcher");

                if (!Directory.Exists(dir))
                    return;

                foreach (var file in Directory.GetFiles(dir))
                {
                    try { File.Delete(file); } catch { }
                }
            }
            catch { }
        }

        public void ProcessNewCertificates(bool checkAllMessages, CancellationToken token)
        {
            CleanupTempFolder();
            token.ThrowIfCancellationRequested();

            int total = 0;
            int addedOrUpdated = 0;
            int skipped = 0;

            Log("=== ImapNewCertificatesWatcher: старт ===");

            if (string.IsNullOrWhiteSpace(_settings.MailHost))
                throw new InvalidOperationException("MailHost пустой.");

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

                    Log($"IMAP подключение: {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");

                    try
                    {
                        client.Authenticate(_settings.MailLogin, _settings.MailPassword);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Ошибка авторизации IMAP: " + ex.Message, ex);
                    }

                    token.ThrowIfCancellationRequested();

                    string folderName = (_settings.ImapNewCertificatesFolder ?? string.Empty).Trim();

                    if (string.IsNullOrWhiteSpace(folderName))
                        throw new InvalidOperationException(
                            "Не задана папка ImapNewCertificatesFolder в server.settings.txt.");

                    IMailFolder root = client.GetFolder(folderName);

                    if (root == null)
                        throw new InvalidOperationException(
                            $"Папка IMAP для новых сертификатов не найдена: '{folderName}'.");

                    try
                    {
                        ProcessFolderRecursive(
                            client,
                            root,
                            checkAllMessages,
                            ref total,
                            ref addedOrUpdated,
                            ref skipped,
                            token);
                    }
                    finally
                    {
                        if (client.IsConnected)
                            client.Disconnect(false);
                    }
                }

                Log($"ИТОГО: обработано={total}, добавлено/обновлено={addedOrUpdated}, пропущено={skipped}");
                Log("=== ImapNewCertificatesWatcher: завершено ===");
            }
            catch (OperationCanceledException)
            {
                Log("Обработка новых сертификатов отменена");
                throw;
            }
            catch (Exception ex)
            {
                Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");
                throw;
            }
        }

        // =====================================================================
        // CORE LOGIC
        // =====================================================================

        private void ProcessFolderRecursive(
            ImapClient client,
            IMailFolder folder,
            bool checkAllMessages,
            ref int total,
            ref int addedOrUpdated,
            ref int skipped,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            if (folder == null)
                return;

            try
            {
                folder.Open(FolderAccess.ReadOnly);
                Log($"Папка '{folder.FullName}' открыта. Всего писем: {folder.Count}");
            }
            catch (Exception ex)
            {
                Log($"Не удалось открыть папку '{folder.FullName}': {ex.Message}");
                return;
            }

            IList<UniqueId> uids;
            long previousLastUid = 0;
            long lastCompletedUid = 0;

            try
            {
                previousLastUid = checkAllMessages
                    ? 0
                    : _db.GetLastUid(folder.FullName);

                if (checkAllMessages)
                {
                    uids = folder.Search(SearchQuery.All)
                        .OrderBy(u => u.Id)
                        .ToList();
                }
                else
                {
                    var foundUids = folder.Search(
                            SearchQuery.Uids(
                                new UniqueIdRange(
                                    new UniqueId((uint)(previousLastUid + 1)),
                                    UniqueId.MaxValue)))
                        .OrderBy(u => u.Id)
                        .ToList();

                    const int MaxPerRun = 200;

                    if (foundUids.Count > MaxPerRun)
                    {
                        foundUids = foundUids
                            .Take(MaxPerRun)
                            .ToList();

                        Log($"Ограничение: обрабатываем первые {MaxPerRun} новых писем за прогон. Остальные будут обработаны в следующих запусках");
                    }

                    uids = foundUids;
                }

                lastCompletedUid = previousLastUid;
            }
            catch (Exception ex)
            {
                Log($"Ошибка поиска писем: {ex.Message}");
                try { folder.Close(); } catch { }
                return;
            }

            IList<IMessageSummary> summaries;

            try
            {
                summaries =
                    uids != null && uids.Count > 0
                        ? folder.Fetch(uids, MessageSummaryItems.BodyStructure)
                        : new List<IMessageSummary>();
            }
            catch (Exception ex)
            {
                Log($"Ошибка получения summaries для папки '{folder.FullName}': {ex.Message}");
                try { folder.Close(); } catch { }
                return;
            }

            foreach (var uid in uids)
            {
                token.ThrowIfCancellationRequested();
                total++;

                bool finalized = false;

                try
                {
                    var summary = summaries.FirstOrDefault(s => s.UniqueId == uid);

                    if (summary == null)
                    {
                        Log($"UID={uid}: summary не получен");
                        skipped++;
                        finalized = false;
                    }
                    else
                    {
                        bool hasZip = summary.Attachments?
                            .Any(a => a.FileName != null &&
                                      a.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                            ?? false;

                        if (!hasZip)
                        {
                            _db.MarkMailProcessed(folder.FullName, uid.ToString(), "NEW");
                            Log($"UID={uid}: пропуск, нет ZIP");
                            skipped++;
                            finalized = true;
                        }
                        else
                        {
                            var message = folder.GetMessage(uid);

                            finalized = ProcessMessage(
                                folder,
                                message,
                                uid.ToString(),
                                ref addedOrUpdated,
                                ref skipped);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"UID={uid}: ошибка обработки письма: {ex.Message}");
                    finalized = false;
                }

                if (!checkAllMessages)
                {
                    if (finalized)
                    {
                        lastCompletedUid = uid.Id;
                    }
                    else
                    {
                        Log($"UID={uid}: остановка прохода, чтобы не перепрыгнуть проблемное письмо");
                        break;
                    }
                }
            }

            if (!checkAllMessages && lastCompletedUid > previousLastUid)
            {
                _db.UpdateLastUid(folder.FullName, lastCompletedUid);
                Log($"LAST_UID обновлён для '{folder.FullName}' -> {lastCompletedUid}");
            }

            Log(
                $"ИТОГО [{folder.FullName}]: " +
                $"обработано всего={total}, добавлено/обновлено={addedOrUpdated}, пропущено={skipped}");

            try { folder.Close(); } catch { }

            IEnumerable<IMailFolder> subfolders;

            try
            {
                subfolders = folder.GetSubfolders(false);
            }
            catch (Exception ex)
            {
                Log($"Ошибка получения подпапок для '{folder.FullName}': {ex.Message}");
                return;
            }

            foreach (var sub in subfolders)
            {
                try
                {
                    ProcessFolderRecursive(
                        client,
                        sub,
                        checkAllMessages,
                        ref total,
                        ref addedOrUpdated,
                        ref skipped,
                        token);
                }
                catch (Exception ex)
                {
                    Log($"Ошибка подпапки {sub.FullName}: {ex.Message}");
                }
            }
        }

        private bool ProcessMessage(
            IMailFolder folder,
            MimeMessage message,
            string uidStr,
            ref int addedOrUpdated,
            ref int skipped)
        {
            string folderPath = folder.FullName;
            string subject = string.IsNullOrWhiteSpace(message.Subject)
                ? "(без темы)"
                : message.Subject.Trim();

            if (_db.IsMailProcessed(folderPath, uidStr, "NEW"))
            {
                skipped++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (уже обработано)");
                return true;
            }

            var zipAttachments = message.Attachments
                .Where(IsZipAttachment)
                .ToList();

            if (!zipAttachments.Any())
            {
                skipped++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (нет ZIP)");
                _db.MarkMailProcessed(folderPath, uidStr, "NEW");
                return true;
            }

            bool anySaved = false;
            bool hadProcessingError = false;

            foreach (var attach in zipAttachments)
            {
                var started = DateTime.UtcNow;

                string zipPath = SaveAttachmentToTemp(attach);
                if (zipPath == null)
                {
                    hadProcessingError = true;
                    Log($"UID={uidStr} | не удалось сохранить ZIP во временный файл");
                    continue;
                }

                try
                {
                    string fromAddress = message.From.Mailboxes.FirstOrDefault()?.Address ?? "";
                    var zipResult = ProcessZip(zipPath, folderPath, uidStr, fromAddress);

                    if (zipResult.savedAny)
                        anySaved = true;

                    if (zipResult.hadErrors)
                        hadProcessingError = true;

                    var elapsed = DateTime.UtcNow - started;
                    if (elapsed.TotalSeconds > 30)
                    {
                        Log($"UID={uidStr} | предупреждение: обработка вложения заняла {elapsed.TotalSeconds:F1} сек");
                    }
                }
                catch (Exception ex)
                {
                    hadProcessingError = true;
                    Log($"UID={uidStr} | ошибка обработки ZIP: {ex.Message}");
                }
                finally
                {
                    try { File.Delete(zipPath); } catch { }
                }
            }

            if (hadProcessingError)
            {
                Log($"UID={uidStr} | Тема=\"{subject}\" | НЕ ПОМЕЧЕНО обработанным из-за ошибок обработки вложений");
                return false;
            }

            _db.MarkMailProcessed(folderPath, uidStr, "NEW");

            if (anySaved)
            {
                addedOrUpdated++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ДОБАВЛЕНО / ОБНОВЛЕНО");
            }
            else
            {
                skipped++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (новых данных не найдено)");
            }

            return true;
        }
        // =====================================================================
        // ZIP → CER → DB
        // =====================================================================

        private (bool savedAny, bool hadErrors) ProcessZip(string zipPath, string folderPath, string uidStr, string fromAddress)
        {
            bool savedAny = false;
            bool hadErrors = false;

            var processedCertNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var cerFiles = ZipCerExtractor.ExtractCerFiles(zipPath, Log);

            if (cerFiles.Count == 0)
                return (false, false);

            foreach (var cerPath in cerFiles)
            {
                try
                {
                    if (!CerCertificateParser.TryParse(cerPath, Log, out var certInfo))
                    {
                        hadErrors = true;
                        continue;
                    }

                    string certNumberNorm = DbHelper.NormalizeCertNumber(certInfo.CertNumber);

                    if (string.IsNullOrWhiteSpace(certNumberNorm))
                    {
                        hadErrors = true;
                        Log($"CER '{Path.GetFileName(cerPath)}': пустой номер сертификата после нормализации");
                        continue;
                    }

                    if (!processedCertNumbers.Add(certNumberNorm))
                        continue;

                    var entry = new CertEntry
                    {
                        Fio = certInfo.Fio,
                        CertNumber = certInfo.CertNumber,
                        DateStart = certInfo.DateStart,
                        DateEnd = certInfo.DateEnd,
                        FromAddress = fromAddress,
                        FolderPath = folderPath,
                        MailUid = uidStr,
                        MessageDate = DateTime.UtcNow
                    };

                    bool wasUpdated;
                    bool wasAdded;
                    int certId;

                    try
                    {
                        (wasUpdated, wasAdded, certId) = _db.InsertOrUpdateAndGetId(entry);
                    }
                    catch (Exception ex)
                    {
                        hadErrors = true;
                        Log($"CER '{Path.GetFileName(cerPath)}': ошибка записи в БД: {ex.Message}");
                        continue;
                    }

                    if (certId <= 0)
                    {
                        hadErrors = true;
                        Log($"CER '{Path.GetFileName(cerPath)}': InsertOrUpdateAndGetId вернул certId <= 0");
                        continue;
                    }

                    // Если запись не добавилась и не обновилась, архив не пересохраняем
                    if (!wasAdded && !wasUpdated)
                    {
                        Log($"CertID={certId}: без изменений");
                        continue;
                    }

                    bool archiveSaved = _db.SaveArchiveToDbTransactional(certId, cerPath);

                    if (!archiveSaved)
                    {
                        hadErrors = true;
                        Log($"CertID={certId}: архив CER не сохранён в БД");
                        continue;
                    }

                    savedAny = true;

                    if (wasAdded)
                        Log($"CertID={certId}: добавлен");
                    else if (wasUpdated)
                        Log($"CertID={certId}: обновлён");
                }
                catch (Exception ex)
                {
                    hadErrors = true;
                    Log($"Ошибка обработки CER '{Path.GetFileName(cerPath)}': {ex.Message}");
                }
                finally
                {
                    try { File.Delete(cerPath); } catch { }
                }
            }

            return (savedAny, hadErrors);
        }

        public Task CheckNewCertificatesFastAsync(CancellationToken token)
        {
            ProcessNewCertificates(false, token);
            return Task.CompletedTask;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================

        private static bool IsZipAttachment(MimeEntity entity)
        {
            string name =
                (entity as MimePart)?.FileName ??
                entity.ContentDisposition?.FileName ??
                entity.ContentType?.Name;

            if (string.IsNullOrEmpty(name))
                return false;

            return name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
        }

        private string SaveAttachmentToTemp(MimeEntity attach)
        {
            if (attach is MimePart part &&
                    part.ContentDisposition != null &&
                    part.ContentDisposition.Size.HasValue &&
                    part.ContentDisposition.Size.Value > 50_000_000)
            {
                Log("ZIP слишком большой — пропущен");
                return null;
            }
            try
            {
                string dir = Path.Combine(Path.GetTempPath(), "ImapCertWatcher");
                Directory.CreateDirectory(dir);

                string filePath = Path.Combine(
                    dir,
                    Guid.NewGuid().ToString("N") + ".zip");

                using (var stream = File.Create(filePath))
                {
                    if (attach is MimePart mp)
                        mp.Content.DecodeTo(stream);
                    else if (attach is MessagePart msg)
                        msg.Message.WriteTo(stream);
                }

                return filePath;
            }
            catch (Exception ex)
            {
                Log($"Ошибка сохранения ZIP: {ex.Message}");
                return null;
            }
        }
    }
}
