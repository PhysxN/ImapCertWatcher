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
        private readonly AppSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        public ImapNewCertificatesWatcher(
            AppSettings settings,
            DbHelper db,
            Action<string> log = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });
        }

        private void Log(string msg)
        {
            _log(msg);
            try
            {
                File.AppendAllText(
                    LogSession.SessionLogFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [NEW] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        // =====================================================================
        // PUBLIC API
        // =====================================================================

        public void ProcessNewCertificates(bool checkAllMessages = false)
        {
            int total = 0;
            int addedOrUpdated = 0;
            int skipped = 0;

            Log("=== ImapNewCertificatesWatcher: старт ===");
            if (string.IsNullOrWhiteSpace(_settings.MailHost))
            {
                Log("ОШИБКА: MailHost пустой");
                return;
            }
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
                    _log($"IMAP подключение: {_settings.MailHost}:{_settings.MailPort}, SSL={_settings.MailUseSsl}");

                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    var root = client.GetFolder(_settings.ImapFolder);
                    ProcessFolderRecursive(
                            client,
                            root,
                            checkAllMessages,
                            ref total,
                            ref addedOrUpdated,
                            ref skipped);

                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex}");
            }
            Log($"ИТОГО: обработано={total}, добавлено/обновлено={addedOrUpdated}, пропущено={skipped}");            
            Log("=== ImapNewCertificatesWatcher: завершено ===");
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
    ref int skipped)
        {
            if (folder == null) return;

            Log($"Папка '{folder.FullName}' открыта. Всего писем: {folder.Count}");

            try
            {
                folder.Open(FolderAccess.ReadOnly);
            }
            catch (Exception ex)
            {
                Log($"Не удалось открыть папку '{folder.FullName}': {ex.Message}");
                return;
            }

            IList<UniqueId> uids;

            try
            {
                long lastUid = checkAllMessages
                    ? 0
                    : _db.GetLastUid(folder.FullName);

                var allUids = folder.Search(SearchQuery.All);

                uids = checkAllMessages
                    ? allUids
                    : allUids.Where(u => u.Id > lastUid)
                             .OrderBy(u => u.Id)
                             .ToList();
            }
            catch (Exception ex)
            {
                Log($"Ошибка поиска писем: {ex.Message}");
                return;
            }

            foreach (var uid in uids)
            {
                try
                {
                    var message = folder.GetMessage(uid);
                    ProcessMessage(
                        folder,
                        message,
                        uid,
                        uid.ToString(),
                        ref total,
                        ref addedOrUpdated,
                        ref skipped);
                }
                catch (Exception ex)
                {
                    Log($"UID={uid}: ошибка обработки письма: {ex.Message}");
                }
            }

            if (!checkAllMessages && uids.Count > 0)
            {
                long maxUid = uids.Max(u => u.Id);
                _db.UpdateLastUid(folder.FullName, maxUid);
            }

            // ✅ СВОДКА
            Log(
                $"ИТОГО [{folder.FullName}]: " +
                $"писем={total}, добавлено/обновлено={addedOrUpdated}, пропущено={skipped}");

            try { folder.Close(); } catch { }

            foreach (var sub in folder.GetSubfolders(false))
                ProcessFolderRecursive(
                    client,
                    sub,
                    checkAllMessages,
                    ref total,
                    ref addedOrUpdated,
                    ref skipped);
        }


        private void ProcessMessage(
    IMailFolder folder,
    MimeMessage message,
    UniqueId _,
    string uidStr,
    ref int total,
    ref int addedOrUpdated,
    ref int skipped)
        {
            total++;

            string folderPath = folder.FullName;
            string subject = string.IsNullOrWhiteSpace(message.Subject)
                ? "(без темы)"
                : message.Subject.Trim();

            if (_db.IsMailProcessed(folderPath, uidStr, "NEW"))
            {
                skipped++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (уже обработано)");
                return;
            }

            var zipAttachments = message.Attachments
                .Where(IsZipAttachment)
                .ToList();

            if (!zipAttachments.Any())
            {
                skipped++;
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (нет ZIP)");
                _db.MarkMailProcessed(folderPath, uidStr, "NEW");
                return;
            }

            bool anySaved = false;

            foreach (var attach in zipAttachments)
            {
                string zipPath = SaveAttachmentToTemp(attach);
                if (zipPath == null)
                    continue;

                try
                {
                    if (ProcessZip(zipPath, folderPath, uidStr))
                        anySaved = true;
                }
                finally
                {
                    try { File.Delete(zipPath); } catch { }
                }
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
                Log($"UID={uidStr} | Тема=\"{subject}\" | ПРОПУЩЕНО (ZIP без CER)");
            }
        }

        // =====================================================================
        // ZIP → CER → DB
        // =====================================================================

        private bool ProcessZip(string zipPath, string folderPath, string uidStr)
        {
            bool savedAny = false; // ⬅ КЛЮЧЕВОЕ

            var cerFiles = ZipCerExtractor.ExtractCerFiles(zipPath, Log);

            if (cerFiles.Count == 0)
            {
                // ❌ УБИРАЕМ шумный лог
                // Log($"ZIP '{Path.GetFileName(zipPath)}': CER не найден");
                return false;
            }

            foreach (var cerPath in cerFiles)
            {
                try
                {
                    if (!CerCertificateParser.TryParse(cerPath, Log, out var certInfo))
                        continue;

                    var entry = new CertEntry
                    {
                        Fio = certInfo.Fio,
                        CertNumber = certInfo.CertNumber,
                        DateStart = certInfo.DateStart,
                        DateEnd = certInfo.DateEnd,
                        FolderPath = folderPath,
                        MailUid = uidStr,
                        MessageDate = DateTime.Now
                    };

                    var (wasUpdated, wasAdded, certId) = _db.InsertOrUpdateAndGetId(entry);

                    if (certId <= 0)
                        continue;

                    _db.DeleteArchiveFromDb(certId);
                    _db.SaveArchiveToDbTransactional(certId, cerPath);

                    savedAny = true; // ⬅ ВОТ РАДИ ЭТОГО ВСЁ

                    // ❗ УКОРОЧЕННЫЙ лог
                    if (wasAdded)
                        Log($"CertID={certId}: добавлен");
                    else if (wasUpdated)
                        Log($"CertID={certId}: обновлён");
                }
                catch (Exception ex)
                {
                    Log($"Ошибка обработки CER '{Path.GetFileName(cerPath)}': {ex.Message}");
                }
                finally
                {
                    try { File.Delete(cerPath); } catch { }
                }
            }

            return savedAny; // ⬅ ОБЯЗАТЕЛЬНО
        }


        public Task CheckNewCertificatesFastAsync()
        {
            return Task.Run(() =>
            {
                ProcessNewCertificates(checkAllMessages: false);
            });
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
