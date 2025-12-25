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
            Log("=== ImapNewCertificatesWatcher: старт ===");

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

                    client.Authenticate(_settings.MailLogin, _settings.MailPassword);

                    var root = client.GetFolder(_settings.ImapFolder);
                    ProcessFolderRecursive(client, root, checkAllMessages);

                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex}");
            }

            Log("=== ImapNewCertificatesWatcher: завершено ===");
        }

        // =====================================================================
        // CORE LOGIC
        // =====================================================================

        private void ProcessFolderRecursive(
            ImapClient client,
            IMailFolder folder,
            bool checkAllMessages)
        {
            if (folder == null) return;

            Log($"Обработка папки: {folder.FullName}");

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
                uids = checkAllMessages
                    ? folder.Search(SearchQuery.All)
                    : folder.Search(SearchQuery.DeliveredAfter(DateTime.Now.AddDays(-5)));
            }
            catch (Exception ex)
            {
                Log($"Ошибка поиска писем: {ex.Message}");
                return;
            }

            foreach (var uid in uids)
            {
                string uidStr = uid.ToString();
                string folderPath = folder.FullName;

                if (_db.IsMailProcessed(folderPath, uidStr, "NEW"))
                    continue;

                try
                {
                    var message = folder.GetMessage(uid);
                    ProcessMessage(folder, message, uid, uidStr);
                }
                catch (Exception ex)
                {
                    Log($"UID={uidStr}: ошибка обработки письма: {ex.Message}");
                }
            }

            try { folder.Close(); } catch { }

            foreach (var sub in folder.GetSubfolders(false))
                ProcessFolderRecursive(client, sub, checkAllMessages);
        }

        private void ProcessMessage(
            IMailFolder folder,
            MimeMessage message,
            UniqueId uid,
            string uidStr)
        {
            string folderPath = folder.FullName;

            var zipAttachments = message.Attachments
                .Where(a => IsZipAttachment(a))
                .ToList();

            if (!zipAttachments.Any())
                return;

            Log($"UID={uidStr}: найдено ZIP-вложений={zipAttachments.Count}");

            foreach (var attach in zipAttachments)
            {
                string zipPath = SaveAttachmentToTemp(attach);
                if (zipPath == null)
                    continue;

                try
                {
                    ProcessZip(zipPath, folderPath, uidStr);
                }
                finally
                {
                    try { File.Delete(zipPath); } catch { }
                }
            }

            _db.MarkMailProcessed(folderPath, uidStr, "NEW");
            Log($"UID={uidStr}: письмо помечено как обработанное");
        }

        // =====================================================================
        // ZIP → CER → DB
        // =====================================================================

        private void ProcessZip(string zipPath, string folderPath, string uidStr)
        {
            var cerFiles = ZipCerExtractor.ExtractCerFiles(zipPath, Log);

            if (cerFiles.Count == 0)
            {
                Log($"ZIP '{Path.GetFileName(zipPath)}': CER не найден");
                return;
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
                        CertNumber = certInfo.SerialNumber,
                        DateStart = certInfo.NotBefore,
                        DateEnd = certInfo.NotAfter,
                        FolderPath = folderPath,
                        MailUid = uidStr,
                        MessageDate = DateTime.Now
                    };

                    var (_, _, certId) = _db.InsertOrUpdateAndGetId(entry);

                    if (certId <= 0)
                        continue;

                    _db.DeleteArchiveFromDb(certId);
                    _db.SaveArchiveToDbTransactional(certId, cerPath);

                    Log($"CertID={certId}: CER сохранён ({Path.GetFileName(cerPath)})");
                }
                catch (Exception ex)
                {
                    Log($"Ошибка обработки CER '{cerPath}': {ex.Message}");
                }
                finally
                {
                    try { File.Delete(cerPath); } catch { }
                }
            }
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
