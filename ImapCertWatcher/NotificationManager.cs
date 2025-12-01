using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ImapCertWatcher.Services
{
    /// <summary>
    /// Отвечает за отправку уведомлений через ObimpCMD + лог + анти-спам.
    /// </summary>
    public class NotificationManager
    {
        private readonly AppSettings _settings;
        private readonly Action<string> _log; // логгер из MainWindow (в session_*.log + мини-лог)

        private readonly string _obimpDir;
        private readonly string _notificationsLogFile;
        private readonly string _sentTodayFile;

        private readonly HashSet<string> _sentToday = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _lock = new object();

        public NotificationManager(AppSettings settings, Action<string> logAction)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _log = logAction ?? (_ => { });

            string appDir = AppDomain.CurrentDomain.BaseDirectory;

            _obimpDir = Path.Combine(appDir, "ObimpCMD");

            string logBaseDirectory = Path.Combine(appDir, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(logBaseDirectory);

            _notificationsLogFile = Path.Combine(logBaseDirectory, "notifications.log");
            _sentTodayFile = Path.Combine(logBaseDirectory, "notifications_sent.txt");

            LoadSentToday();
        }

        /// <summary>
        /// Загружаем список уже отправленных сегодня уведомлений (антиспам).
        /// </summary>
        private void LoadSentToday()
        {
            try
            {
                if (File.Exists(_sentTodayFile))
                {
                    foreach (var line in File.ReadAllLines(_sentTodayFile, Encoding.UTF8))
                    {
                        var key = line.Trim();
                        if (!string.IsNullOrEmpty(key))
                            _sentToday.Add(key);
                    }
                }
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка чтения {_sentTodayFile}: {ex.Message}");
            }
        }

        private void SaveSentToday()
        {
            try
            {
                File.WriteAllLines(_sentTodayFile, _sentToday, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка записи {_sentTodayFile}: {ex.Message}");
            }
        }

        private void LogNotification(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            try
            {
                File.AppendAllText(_notificationsLogFile, line + Environment.NewLine, Encoding.UTF8);
            }
            catch { /* не валим приложение из-за лога */ }

            _log($"[Notify] {message}");
        }

        /// <summary>
        /// Возвращает список аккаунтов для здания
        /// building: "Краснофлотская", "Пионерская" или "ALL" (объединить).
        ///</summary>
        private List<string> GetAccountsForBuilding(string building)
        {
            string raw;

            if (string.Equals(building, "Краснофлотская", StringComparison.OrdinalIgnoreCase))
            {
                raw = _settings.BimoidAccountsKrasnoflotskaya;
            }
            else if (string.Equals(building, "Пионерская", StringComparison.OrdinalIgnoreCase))
            {
                raw = _settings.BimoidAccountsPionerskaya;
            }
            else if (string.Equals(building, "ALL", StringComparison.OrdinalIgnoreCase))
            {
                raw = string.Join(Environment.NewLine,
                    _settings.BimoidAccountsKrasnoflotskaya ?? string.Empty,
                    _settings.BimoidAccountsPionerskaya ?? string.Empty);
            }
            else
            {
                // неизвестное здание — ничего не шлём
                return new List<string>();
            }

            if (string.IsNullOrWhiteSpace(raw))
                return new List<string>();

            return raw
                .Split(new[] { '\r', '\n', ';', ',', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(a => a.Trim())
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// Генерирует ключ для антиспама "один раз в сутки".
        /// type: "expiry" / "new".
        /// </summary>
        private string MakeSpamKey(string type, CertRecord rec)
        {
            // Привязываемся к конкретному сертификату + зданию + дате проверки
            return $"{DateTime.Today:yyyy-MM-dd}|{type}|{rec.Fio}|{rec.Building}|{rec.CertNumber}";
        }

        /// <summary>
        /// true = уже сегодня слали, повторять нельзя.
        /// </summary>
        private bool AlreadySentToday(string type, CertRecord rec)
        {
            var key = MakeSpamKey(type, rec);
            lock (_lock)
            {
                if (_sentToday.Contains(key))
                    return true;

                _sentToday.Add(key);
                SaveSentToday();
                return false;
            }
        }

        /// <summary>
        /// Низкоуровневый запуск demo.bat с указанными аккаунтами и текстом.
        /// </summary>
        private void SendMessageInternal(IEnumerable<string> accounts, string message, string tagForLog)
        {
            try
            {
                if (!Directory.Exists(_obimpDir))
                {
                    _log($"[Notify] ObimpCMD не найден: {_obimpDir}");
                    return;
                }

                string demoBat = Path.Combine(_obimpDir, "demo.bat");
                if (!File.Exists(demoBat))
                {
                    _log($"[Notify] demo.bat не найден: {demoBat}");
                    return;
                }

                var accList = (accounts ?? Enumerable.Empty<string>()).ToList();
                if (accList.Count == 0)
                {
                    _log($"[Notify] Нет аккаунтов для отправки ({tagForLog})");
                    return;
                }

                string accountsFile = Path.Combine(_obimpDir, "accounts.txt");
                string messageFile = Path.Combine(_obimpDir, "message.txt");

                File.WriteAllLines(accountsFile, accList, Encoding.UTF8);
                File.WriteAllText(messageFile, message ?? string.Empty, Encoding.UTF8);

                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c \"demo.bat\"",
                    WorkingDirectory = _obimpDir,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                Process.Start(psi);

                LogNotification($"Отправлено ({tagForLog}) {accList.Count} получателям.");
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка отправки через ObimpCMD: {ex.Message}");
            }
        }

        /// <summary>
        /// Уведомления о скором окончании сертификата.
        /// Антиспам: по одному сообщению в сутки на каждый сертификат.
        /// </summary>
        public void NotifyExpiringCerts(IEnumerable<CertRecord> records)
        {
            try
            {
                int threshold = _settings.NotifyDaysThreshold > 0
                    ? _settings.NotifyDaysThreshold
                    : 10;

                var list = records?
                    .Where(r => !r.IsDeleted && r.DaysLeft >= 0 && r.DaysLeft <= threshold)
                    .ToList() ?? new List<CertRecord>();

                if (list.Count == 0)
                    return;

                foreach (var rec in list)
                {
                    if (AlreadySentToday("expiry", rec))
                        continue; // уже слали сегодня

                    var accounts = GetAccountsForBuilding(rec.Building);
                    if (accounts.Count == 0)
                    {
                        _log($"[Notify] Нет аккаунтов для здания '{rec.Building}', пропуск {rec.Fio}");
                        continue;
                    }

                    string msg = $"У {rec.Fio} осталось {rec.DaysLeft} дней до окончания сертификата.";
                    string tag = $"expiry, {rec.Fio}, {rec.Building}";
                    SendMessageInternal(accounts, msg, tag);
                }

                if (list.Count > 0)
                    LogNotification($"Проверка истекающих сертификатов: найдено {list.Count} записей (порог {threshold} дней).");
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка NotifyExpiringCerts: {ex.Message}");
            }
        }

        /// <summary>
        /// Уведомление о новых пользователях:
        /// принимает уже отфильтрованный список "новых" записей.
        /// Оповещает всех контактов (объединённый список по зданиям).
        /// Антиспам: по одному сообщению в сутки на пользователя+сертификат.
        /// </summary>
        public void NotifyNewUsers(IEnumerable<CertRecord> records)
        {
            try
            {
                var list = records?
                    .Where(r => !r.IsDeleted)
                    .ToList() ?? new List<CertRecord>();

                if (list.Count == 0)
                    return;

                // Берём все контакты (объединяем списки для Краснофлотской и Пионерской)
                var allAccounts = GetAccountsForBuilding("ALL");
                if (allAccounts.Count == 0)
                {
                    _log("[Notify] Нет Bimoid аккаунтов для оповещения о новых пользователях.");
                    return;
                }

                foreach (var rec in list)
                {
                    if (AlreadySentToday("new", rec))
                        continue; // уже оповещали сегодня

                    string msg =
                        $"Добавлен новый пользователь: {rec.Fio}. " +
                        $"Дата сертификата: {rec.DateStart:dd.MM.yyyy} - {rec.DateEnd:dd.MM.yyyy}.";

                    string tag = $"new, {rec.Fio}, {rec.Building}";
                    SendMessageInternal(allAccounts, msg, tag);
                }

                if (list.Count > 0)
                    LogNotification($"Оповещены о новых пользователях: {list.Count} записей.");
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка NotifyNewUsers: {ex.Message}");
            }
        }

        /// <summary>
        /// Тестовое сообщение для кнопки "Тестовое сообщение".
        /// </summary>
        public void SendTestMessage()
        {
            try
            {
                var allAccounts = GetAccountsForBuilding("ALL");
                if (allAccounts.Count == 0)
                {
                    _log("[Notify] Тест: нет аккаунтов для отправки.");
                    return;
                }

                string msg = "Тестовое сообщение от ImapCertWatcher (Bimoid/ObimpCMD).";
                SendMessageInternal(allAccounts, msg, "test");
            }
            catch (Exception ex)
            {
                _log($"[Notify] Ошибка SendTestMessage: {ex.Message}");
            }
        }
    }
}
