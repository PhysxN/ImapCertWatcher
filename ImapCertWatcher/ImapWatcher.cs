using FirebirdSql.Data.FirebirdClient;
using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Authentication;
using System.Text.RegularExpressions;

namespace ImapCertWatcher.Services
{
    public class ImapWatcher
    {
        private readonly AppSettings _settings;
        private readonly DbHelper _db;
        private readonly string _attachmentsBasePath;
        private readonly string _logDirectory;
        private readonly Action<string> _addToMiniLog;

        private static Regex certNumberRegex = new Regex(@"Сертификат №\s*(?<number>[0-9A-F]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex fioRegex = new Regex(@"ФИО:\s*(?<fio>[\p{L}\s\-]+?)(?:\.|\n|$)", RegexOptions.Compiled);
        private static Regex datesRegex = new Regex(@"с\s*(?<ds>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})\s*по\s*(?<de>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex fioRegexAlt1 = new Regex(@"ФИО:\s*(?<fio>[\p{L}\s\-]+)\s*$", RegexOptions.Compiled | RegexOptions.Multiline);
        private static Regex fioRegexAlt2 = new Regex(@"ФИО\s*=\s*(?<fio>[\p{L}\s\-]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex datesRegexAlt1 = new Regex(@"с\s*(?<ds>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}).*?по\s*(?<de>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})", RegexOptions.Compiled | RegexOptions.Singleline);
        private static Regex datesRegexAlt2 = new Regex(@"Срок действия сертификата:\s*с\s*(?<ds>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})\s*по\s*(?<de>\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public ImapWatcher(AppSettings settings, DbHelper db, Action<string> addToMiniLog = null)
        {
            _settings = settings;
            _db = db;
            _addToMiniLog = addToMiniLog;
            _attachmentsBasePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs");
            _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");

            // Создаем папку для архивов, если ее нет
            if (!Directory.Exists(_attachmentsBasePath))
                Directory.CreateDirectory(_attachmentsBasePath);

            // Создаем папку для логов, если ее нет
            if (!Directory.Exists(_logDirectory))
                Directory.CreateDirectory(_logDirectory);

            Log("Инициализация ImapWatcher завершена");
        }



        public (List<CertEntry> processedEntries, int updatedCount, int addedCount) CheckMail(bool checkAllMessages = false)
        {
            var results = new List<CertEntry>();
            int updatedCount = 0;
            int addedCount = 0;

            Log($"Начало проверки почты. Режим: {(checkAllMessages ? "все письма" : "последние 3 дня")}");

            // Проверяем настройки перед подключением
            try
            {
                ValidateMailSettings();
            }
            catch (Exception ex)
            {
                Log($"Ошибка в настройках почты: {ex.Message}");
                return (results, updatedCount, addedCount); // Возвращаем пустой список вместо исключения
            }

            const int maxRetries = 3;
            const int retryDelayMs = 2000; // 2 секунды между попытками

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var client = new ImapClient())
                    {
                        // Увеличиваем таймауты
                        client.Timeout = 120 * 1000; // 2 минуты
                        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                        Log($"Попытка {attempt}/{maxRetries}: подключение к {_settings.MailHost}:{_settings.MailPort} (SSL: {_settings.MailUseSsl})");

                        if (_settings.MailUseSsl)
                        {
                            client.Connect(_settings.MailHost, _settings.MailPort,
                                         MailKit.Security.SecureSocketOptions.SslOnConnect);
                        }
                        else
                        {
                            // Пробуем разные варианты подключения
                            try
                            {
                                client.Connect(_settings.MailHost, _settings.MailPort,
                                             MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);
                                Log("Подключение через StartTLS успешно");
                            }
                            catch (Exception ex)
                            {
                                Log($"StartTLS не поддерживается, пробуем без SSL: {ex.Message}");
                                client.Connect(_settings.MailHost, _settings.MailPort,
                                             MailKit.Security.SecureSocketOptions.None);
                                Log("Подключение без SSL успешно");
                            }
                        }

                        Log("Подключение к почтовому серверу успешно");

                        Log($"Выполняем аутентификацию пользователя: {_settings.MailLogin}");
                        client.Authenticate(_settings.MailLogin, _settings.MailPassword);
                        Log("Аутентификация успешна");

                        Log($"Поиск папки: {_settings.ImapFolder}");
                        var mainFolder = GetFolderRecursive(client, _settings.ImapFolder);
                        if (mainFolder == null)
                        {
                            Log($"Не удалось найти папку: {_settings.ImapFolder}");
                            break;
                        }

                        Log($"Основная папка найдена: {mainFolder.FullName}");

                        var allFolders = GetAllSubfoldersRecursive(mainFolder);
                        allFolders.Insert(0, mainFolder);

                        Log($"Будет проверено папок: {allFolders.Count}");

                        foreach (var folder in allFolders)
                        {
                            try
                            {
                                Log($"Открываем папку: {folder.FullName}");
                                folder.Open(FolderAccess.ReadOnly);

                                Log($"В папке {folder.FullName}: {folder.Count} писем");

                                // Создаем запрос в зависимости от режима
                                SearchQuery query;
                                if (checkAllMessages)
                                {
                                    // Все письма
                                    query = SearchQuery.SubjectContains("Сертификат №")
                                                .And(SearchQuery.FromContains(_settings.FilterRecipient));
                                    Log("Режим: проверка ВСЕХ писем");
                                }
                                else
                                {
                                    // Только письма за последние 3 дня
                                    var threeDaysAgo = DateTime.Now.AddDays(-3);
                                    query = SearchQuery.SubjectContains("Сертификат №")
                                                .And(SearchQuery.FromContains(_settings.FilterRecipient))
                                                .And(SearchQuery.DeliveredAfter(threeDaysAgo));
                                    Log($"Режим: проверка писем за последние 3 дня (с {threeDaysAgo:dd.MM.yyyy})");
                                }

                                Log($"Выполняем поиск писем с критериями: Тема содержит 'Сертификат №', От: {_settings.FilterRecipient}");
                                var uids = folder.Search(query);
                                Log($"В папке {folder.FullName} найдено писем для обработки: {uids.Count}");

                                foreach (var uid in uids)
                                {
                                    try
                                    {
                                        Log($"Получаем письмо UID: {uid}");
                                        var msg = folder.GetMessage(uid);
                                        var subject = msg.Subject ?? "";
                                        var fromAddress = msg.From.Mailboxes.FirstOrDefault()?.Address ?? "";
                                        var body = GetMessageBody(msg);
                                        var messageDate = msg.Date.UtcDateTime.ToLocalTime();

                                        Log($"Обрабатываем письмо из папки {folder.FullName}: {subject}, От: {fromAddress}, Дата: {messageDate:dd.MM.yyyy HH:mm}");

                                        if (!subject.StartsWith("Сертификат №", StringComparison.OrdinalIgnoreCase))
                                        {
                                            Log($"Пропускаем письмо - тема не начинается с 'Сертификат №': {subject}");
                                            continue;
                                        }

                                        if (!fromAddress.Contains(_settings.FilterRecipient))
                                        {
                                            Log($"Пропускаем письмо - отправитель не совпадает: {fromAddress} (ожидается: {_settings.FilterRecipient})");
                                            continue;
                                        }

                                        Log($"Извлекаем данные из письма: номер сертификата, ФИО, даты");
                                        var certNumber = ExtractCertNumber(subject);
                                        var fio = ExtractFio(body);
                                        var dates = ExtractDates(body);

                                        Log($"Извлеченные данные: Сертификат={certNumber}, ФИО={fio}, Дата начала={dates.start}, Дата окончания={dates.end}");

                                        if (string.IsNullOrEmpty(fio))
                                        {
                                            Log($"Не найден ФИО в письме {uid}");
                                            continue;
                                        }

                                        if (dates.start == DateTime.MinValue || dates.end == DateTime.MinValue)
                                        {
                                            Log($"Не найдены даты в письме {uid}");
                                            continue;
                                        }

                                        // ★★★★ ПРОВЕРКА ДУБЛИКАТА ДО ОБРАБОТКИ ВЛОЖЕНИЙ ★★★★
                                        Log($"Проверяем наличие дубликата в БД: {fio}, сертификат: {certNumber}");
                                        bool isDuplicate = _db.CheckDuplicate(fio, certNumber, messageDate);

                                        if (isDuplicate)
                                        {
                                            Log($"Пропускаем письмо (дубликат найден в БД): {fio}, сертификат: {certNumber}");
                                            continue; // Переходим к следующему письму
                                        }

                                        Log($"Дубликат не найден, продолжаем обработку письма");

                                        // Сохраняем вложения (ZIP архивы) - только если это не дубликат
                                        Log("Проверяем вложения письма");
                                        string archivePath = SaveAttachments(msg, fio, certNumber);

                                        var daysLeft = (int)Math.Ceiling((dates.end - DateTime.Now).TotalDays);
                                        if (daysLeft < 0) daysLeft = 0;

                                        Log($"Рассчитано осталось дней: {daysLeft}");

                                        var entry = new CertEntry
                                        {
                                            MailUid = uid.ToString(),
                                            Fio = fio,
                                            DateStart = dates.start,
                                            DateEnd = dates.end,
                                            DaysLeft = daysLeft,
                                            Subject = subject,
                                            Received = msg.Date.UtcDateTime.ToLocalTime(),
                                            CertNumber = certNumber,
                                            FromAddress = fromAddress,
                                            FolderPath = folder.FullName,
                                            ArchivePath = archivePath,
                                            MessageDate = messageDate
                                        };

                                        try
                                        {
                                            Log($"Сохраняем запись в БД: {fio}, сертификат: {certNumber}");
                                            var (wasUpdated, wasAdded) = _db.InsertOrUpdate(entry);

                                            if (wasUpdated)
                                            {
                                                updatedCount++;
                                                Log($"Успешно ОБНОВЛЕНО письмо из папки {folder.FullName}: {fio}, {dates.start:dd.MM.yyyy} - {dates.end:dd.MM.yyyy}, Сертификат: {certNumber}");
                                            }
                                            else if (wasAdded)
                                            {
                                                addedCount++;
                                                Log($"Успешно ДОБАВЛЕНО письмо из папки {folder.FullName}: {fio}, {dates.start:dd.MM.yyyy} - {dates.end:dd.MM.yyyy}, Сертификат: {certNumber}");
                                            }
                                            else
                                            {
                                                Log($"Пропущено письмо (дубликат при сохранении): {fio}, сертификат: {certNumber}");
                                            }

                                            results.Add(entry);
                                        }
                                        catch (Exception ex)
                                        {
                                            Log($"Ошибка сохранения в БД: {ex.Message}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Ошибка обработки письма {uid}: {ex.Message}");
                                    }
                                }

                                folder.Close();
                                Log($"Папка {folder.FullName} обработана");
                            }
                            catch (Exception ex)
                            {
                                Log($"Ошибка при работе с папкой {folder.FullName}: {ex.Message}");
                            }
                        }

                        client.Disconnect(true);
                        Log($"Обработка завершена. Проверено папок: {allFolders.Count}, Обработано писем: {results.Count}, Обновлено: {updatedCount}, Добавлено: {addedCount}");

                        // Если дошли до сюда - подключение успешно, выходим из цикла повторных попыток
                        break;
                    }
                }
                catch (System.Net.Sockets.SocketException socketEx)
                {
                    Log($"Попытка {attempt}/{maxRetries}: Ошибка подключения к почтовому серверу {_settings.MailHost}:{_settings.MailPort}. Детали: {socketEx.Message}");

                    if (attempt < maxRetries)
                    {
                        Log($"Повторная попытка через {retryDelayMs}мс...");
                        System.Threading.Thread.Sleep(retryDelayMs);
                    }
                    else
                    {
                        Log($"Все {maxRetries} попытки подключения завершились ошибкой");
                    }
                }
                catch (System.IO.IOException ioEx)
                {
                    Log($"Попытка {attempt}/{maxRetries}: Ошибка ввода-вывода. Детали: {ioEx.Message}");

                    if (attempt < maxRetries)
                    {
                        Log($"Повторная попытка через {retryDelayMs}мс...");
                        System.Threading.Thread.Sleep(retryDelayMs);
                    }
                }
                catch (TimeoutException timeoutEx)
                {
                    Log($"Попытка {attempt}/{maxRetries}: Таймаут подключения. Детали: {timeoutEx.Message}");

                    if (attempt < maxRetries)
                    {
                        Log($"Повторная попытка через {retryDelayMs}мс...");
                        System.Threading.Thread.Sleep(retryDelayMs);
                    }
                }
                catch (AuthenticationException authEx)
                {
                    Log($"Попытка {attempt}/{maxRetries}: Ошибка аутентификации. Детали: {authEx.Message}");
                    break; // При ошибке аутентификации повторные попытки бессмысленны
                }
                catch (Exception ex)
                {
                    Log($"Попытка {attempt}/{maxRetries}: Общая ошибка работы с почтовым сервером: {ex.Message}");

                    if (attempt < maxRetries)
                    {
                        Log($"Повторная попытка через {retryDelayMs}мс...");
                        System.Threading.Thread.Sleep(retryDelayMs);
                    }
                }
            }

            return (results, updatedCount, addedCount);
        }

        
        private string SaveAttachments(MimeKit.MimeMessage msg, string fio, string certNumber)
        {
            try
            {
                if (!msg.Attachments.Any())
                {
                    Log("В письме нет вложений");
                    return null;
                }

                Log($"Проверяем вложения письма. Количество: {msg.Attachments.Count()}");

                // Создаем безопасное имя папки для ФИО
                var safeFio = MakeValidFileName(fio);
                var certFolder = Path.Combine(_attachmentsBasePath, safeFio);

                if (!Directory.Exists(certFolder))
                {
                    Log($"Создаем папку для сертификатов: {certFolder}");
                    Directory.CreateDirectory(certFolder);
                }

                string savedArchivePath = null;
                int attachmentCount = 0;

                foreach (var attachment in msg.Attachments)
                {
                    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    attachmentCount++;
                    Log($"Обрабатываем вложение {attachmentCount}: {fileName}");

                    // Сохраняем только ZIP архивы
                    if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        var filePath = Path.Combine(certFolder, $"{certNumber}_{fileName}");

                        Log($"Сохраняем ZIP архив: {filePath}");
                        using (var stream = File.Create(filePath))
                        {
                            if (attachment is MimeKit.MessagePart)
                            {
                                var part = (MimeKit.MessagePart)attachment;
                                part.Message.WriteTo(stream);
                            }
                            else
                            {
                                var part = (MimeKit.MimePart)attachment;
                                part.Content.DecodeTo(stream);
                            }
                        }

                        savedArchivePath = filePath;
                        Log($"Сохранен архив: {filePath}");
                    }
                    else
                    {
                        Log($"Пропускаем вложение (не ZIP): {fileName}");
                    }
                }

                Log($"Обработано вложений: {attachmentCount}, сохранено архивов: {(savedArchivePath != null ? 1 : 0)}");
                return savedArchivePath;
            }
            catch (Exception ex)
            {
                Log($"Ошибка при сохранении вложений: {ex.Message}");
                return null;
            }
        }

        private string MakeValidFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(name.Where(ch => !invalidChars.Contains(ch)).ToArray());
        }

        // Метод для распаковки архива
        public bool ExtractArchive(string archivePath, string fio)
        {
            try
            {
                if (string.IsNullOrEmpty(archivePath) || !File.Exists(archivePath))
                {
                    Log($"Архив не найден: {archivePath}");
                    return false;
                }

                Log($"Распаковываем архив: {archivePath} для {fio}");
                var safeFio = MakeValidFileName(fio);
                var extractPath = Path.Combine(_attachmentsBasePath, safeFio, "extracted");

                if (Directory.Exists(extractPath))
                {
                    Log($"Удаляем существующую папку для распаковки: {extractPath}");
                    Directory.Delete(extractPath, true);
                }

                Directory.CreateDirectory(extractPath);
                Log($"Создана папка для распаковки: {extractPath}");

                // Используем полное имя с указанием пространства имен
                System.IO.Compression.ZipFile.ExtractToDirectory(archivePath, extractPath);
                Log($"Архив распакован: {archivePath} -> {extractPath}");

                // Открываем папку с распакованными файлами
                Log($"Открываем папку с распакованными файлами: {extractPath}");
                System.Diagnostics.Process.Start("explorer.exe", extractPath);

                return true;
            }
            catch (Exception ex)
            {
                Log($"Ошибка при распаковке архива: {ex.Message}");
                return false;
            }
        }

        private string ExtractFio(string body)
        {
            Log("Извлекаем ФИО из тела письма");

            var match = fioRegex.Match(body);
            if (match.Success)
            {
                Log($"ФИО найдено (основной шаблон): {match.Groups["fio"].Value.Trim()}");
                return match.Groups["fio"].Value.Trim();
            }

            match = fioRegexAlt1.Match(body);
            if (match.Success)
            {
                Log($"ФИО найдено (альтернативный шаблон 1): {match.Groups["fio"].Value.Trim()}");
                return match.Groups["fio"].Value.Trim();
            }

            match = fioRegexAlt2.Match(body);
            if (match.Success)
            {
                Log($"ФИО найдено (альтернативный шаблон 2): {match.Groups["fio"].Value.Trim()}");
                return match.Groups["fio"].Value.Trim();
            }

            // Дополнительная попытка найти ФИО в формате "ФИО: Имя Отчество Фамилия"
            var fioPatterns = new[]
            {
                @"ФИО[:\s]+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)",
                @"ФИО[:\s]+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][\.]\s*[А-ЯЁ][\.])",
                @"ФИО\s*=\s*([^\.\r\n]+)"
            };

            foreach (var pattern in fioPatterns)
            {
                match = Regex.Match(body, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    Log($"ФИО найдено (дополнительный шаблон): {match.Groups[1].Value.Trim()}");
                    return match.Groups[1].Value.Trim();
                }
            }

            Log("ФИО не найдено в теле письма");
            return null;
        }

        private (DateTime start, DateTime end) ExtractDates(string body)
        {
            Log("Извлекаем даты из тела письма");

            var match = datesRegex.Match(body);
            if (match.Success)
            {
                if (TryParseDate(match.Groups["ds"].Value, out DateTime start) &&
                    TryParseDate(match.Groups["de"].Value, out DateTime end))
                {
                    Log($"Даты найдены (основной шаблон): {start} - {end}");
                    return (start, end);
                }
            }

            match = datesRegexAlt1.Match(body);
            if (match.Success)
            {
                if (TryParseDate(match.Groups["ds"].Value, out DateTime start) &&
                    TryParseDate(match.Groups["de"].Value, out DateTime end))
                {
                    Log($"Даты найдены (альтернативный шаблон 1): {start} - {end}");
                    return (start, end);
                }
            }

            match = datesRegexAlt2.Match(body);
            if (match.Success)
            {
                if (TryParseDate(match.Groups["ds"].Value, out DateTime start) &&
                    TryParseDate(match.Groups["de"].Value, out DateTime end))
                {
                    Log($"Даты найдены (альтернативный шаблон 2): {start} - {end}");
                    return (start, end);
                }
            }

            Log("Даты не найдены в теле письма");
            return (DateTime.MinValue, DateTime.MinValue);
        }

        private bool TryParseDate(string dateString, out DateTime result)
        {
            Log($"Парсим дату: {dateString}");

            var formats = new[]
            {
                "dd.MM.yyyy HH:mm:ss",
                "dd.MM.yyyy H:mm:ss",
                "d.MM.yyyy HH:mm:ss",
                "d.MM.yyyy H:mm:ss"
            };

            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(dateString.Trim(), format,
                    System.Globalization.CultureInfo.GetCultureInfo("ru-RU"),
                    System.Globalization.DateTimeStyles.None, out result))
                {
                    Log($"Дата успешно распарсена: {result}");
                    return true;
                }
            }

            Log($"Не удалось распарсить дату: {dateString}");
            result = DateTime.MinValue;
            return false;
        }

        private IMailFolder GetFolderRecursive(ImapClient client, string folderName)
        {
            try
            {
                var folder = client.GetFolder(folderName);
                if (folder != null)
                    return folder;
            }
            catch (FolderNotFoundException)
            {
                Log($"Папка '{folderName}' не найдена, начинаем рекурсивный поиск...");
            }

            var personal = client.GetFolder(client.PersonalNamespaces[0]);
            return FindFolderRecursive(personal, folderName);
        }

        private IMailFolder FindFolderRecursive(IMailFolder parent, string folderName)
        {
            try
            {
                var subfolders = parent.GetSubfolders(false);

                foreach (var folder in subfolders)
                {
                    Log($"Проверяем папку: {folder.FullName}");

                    if (folder.FullName.Equals(folderName, StringComparison.OrdinalIgnoreCase) ||
                        folder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"Найдена папка: {folder.FullName}");
                        return folder;
                    }

                    try
                    {
                        var foundFolder = FindFolderRecursive(folder, folderName);
                        if (foundFolder != null)
                            return foundFolder;
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка при поиске в папке {folder.Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при получении подпапок для {parent.Name}: {ex.Message}");
            }

            return null;
        }

        private List<IMailFolder> GetAllSubfoldersRecursive(IMailFolder parent)
        {
            var folders = new List<IMailFolder>();

            try
            {
                var subfolders = parent.GetSubfolders(false);
                foreach (var folder in subfolders)
                {
                    folders.Add(folder);
                    folders.AddRange(GetAllSubfoldersRecursive(folder));
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при получении подпапок из {parent.FullName}: {ex.Message}");
            }

            return folders;
        }

        public (List<CertEntry> processedEntries, int updatedCount, int addedCount) ProcessAllExistingEmails()
        {
            return CheckMail(true);
        }

        private string GetMessageBody(MimeKit.MimeMessage msg)
        {
            var textPart = msg.TextBody;
            if (!string.IsNullOrEmpty(textPart))
                return textPart;

            var htmlPart = msg.HtmlBody;
            if (!string.IsNullOrEmpty(htmlPart))
            {
                return System.Text.RegularExpressions.Regex.Replace(htmlPart, "<.*?>", string.Empty);
            }

            return msg.Body?.ToString() ?? "";
        }

        private string ExtractCertNumber(string subject)
        {
            var match = certNumberRegex.Match(subject);
            string result = match.Success ? match.Groups["number"].Value.Trim() : "Неизвестно";
            Log($"Извлечен номер сертификата: {result}");
            return result;
        }

        public static List<string> GetMailFolders(AppSettings settings)
        {
            var folders = new List<string>();
            LogStatic("Начало загрузки списка папок с почтового сервера");

            // Проверяем настройки
            if (string.IsNullOrEmpty(settings.MailHost))
                throw new Exception("Не указан хост почтового сервера");
            if (string.IsNullOrEmpty(settings.MailLogin))
                throw new Exception("Не указан логин почтового ящика");
            if (string.IsNullOrEmpty(settings.MailPassword))
                throw new Exception("Не указан пароль почтового ящика");
            if (settings.MailPort <= 0)
                throw new Exception("Неверный порт почтового сервера");

            using (var client = new ImapClient())
            {
                client.Timeout = 30 * 1000; // 30 секунд
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                try
                {
                    LogStatic($"Подключение к {settings.MailHost}:{settings.MailPort} (SSL: {settings.MailUseSsl})");

                    if (settings.MailUseSsl)
                    {
                        client.Connect(settings.MailHost, settings.MailPort,
                                     MailKit.Security.SecureSocketOptions.SslOnConnect);
                    }
                    else
                    {
                        try
                        {
                            client.Connect(settings.MailHost, settings.MailPort,
                                         MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);
                            LogStatic("Подключение через StartTLS успешно");
                        }
                        catch (Exception ex)
                        {
                            LogStatic($"StartTLS не поддерживается, пробуем без SSL: {ex.Message}");
                            client.Connect(settings.MailHost, settings.MailPort,
                                         MailKit.Security.SecureSocketOptions.None);
                            LogStatic("Подключение без SSL успешно");
                        }
                    }

                    LogStatic("Подключение к почтовому серверу успешно");

                    LogStatic($"Аутентификация пользователя: {settings.MailLogin}");
                    client.Authenticate(settings.MailLogin, settings.MailPassword);
                    LogStatic("Аутентификация успешна");

                    var personal = client.GetFolder(client.PersonalNamespaces[0]);
                    var allFolders = GetAllFoldersRecursiveStatic(personal);

                    foreach (var folder in allFolders)
                    {
                        folders.Add(folder.FullName);
                    }

                    folders.Sort();
                    client.Disconnect(true);
                    LogStatic("Отключение от почтового сервера");
                }
                catch (Exception ex)
                {
                    LogStatic($"Ошибка при загрузке папок: {ex.Message}");
                    throw new Exception($"Не удалось подключиться к почтовому серверу: {ex.Message}", ex);
                }
            }

            LogStatic($"Загружено папок: {folders.Count}");
            return folders;
        }

        private static List<IMailFolder> GetAllFoldersRecursiveStatic(IMailFolder parent)
        {
            var folders = new List<IMailFolder>();

            try
            {
                var subfolders = parent.GetSubfolders(false);
                folders.AddRange(subfolders);

                foreach (var folder in subfolders)
                {
                    folders.AddRange(GetAllFoldersRecursiveStatic(folder));
                }
            }
            catch (Exception ex)
            {
                LogStatic($"Ошибка при получении папок из {parent.Name}: {ex.Message}");
            }

            return folders;
        }

        private void Log(string message)
        {
            try
            {
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                string logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [ImapWatcher] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

                // Также пишем в мини-лог через переданный делегат
                _addToMiniLog?.Invoke($"[ImapWatcher] {message}");

                System.Diagnostics.Debug.WriteLine($"[ImapWatcher] {message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи в лог: {ex.Message}");
            }
        }

        private static void LogStatic(string message)
        {
            try
            {
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                string logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [ImapWatcher] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

                // Для статических методов не вызываем мини-лог, т.к. нет доступа к экземпляру
                System.Diagnostics.Debug.WriteLine($"[ImapWatcher] {message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи в лог: {ex.Message}");
            }
        }

        private void ValidateMailSettings()
        {
            if (string.IsNullOrEmpty(_settings.MailHost))
                throw new Exception("Не указан хост почтового сервера");

            if (string.IsNullOrEmpty(_settings.MailLogin))
                throw new Exception("Не указан логин почтового ящика");

            if (string.IsNullOrEmpty(_settings.MailPassword))
                throw new Exception("Не указан пароль почтового ящика");

            if (_settings.MailPort <= 0)
                throw new Exception("Неверный порт почтового сервера");

            Log("Проверка настроек почты завершена успешно");
        }
    }
}