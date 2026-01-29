using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;

namespace ImapCertWatcher.Services
{
    public class NotificationManager
    {
        private readonly ServerSettings _settings;
        private readonly Action<string> _log;
        private readonly string _obimpDir;
        private readonly string _stateFile;
        private readonly object _sync = new object();

        // ===== ВНУТРЕННЕЕ СОСТОЯНИЕ АНТИ-СПАМА =====
        private class NotificationState
        {
            // ключ: "expiring:CertId"
            public Dictionary<string, DateTime> ExpiringLastSent { get; set; }
            // ключ: "newuser:CertId"
            public Dictionary<string, DateTime> NewUserLastSent { get; set; }

            public NotificationState()
            {
                ExpiringLastSent = new Dictionary<string, DateTime>();
                NewUserLastSent = new Dictionary<string, DateTime>();
            }
        }

        private NotificationState _state = new NotificationState();

        public NotificationManager(ServerSettings settings, Action<string> log)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _log = log ?? (_ => { });

            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            _obimpDir = Path.Combine(baseDir, "ObimpCMD");

            var logDir = Path.Combine(baseDir, "LOG");
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            _stateFile = Path.Combine(logDir, "notifications_state.txt");
            LoadState();
        }

        // ===== ПУБЛИЧНЫЕ МЕТОДЫ =====

        /// <summary>
        /// Уведомления об истекающих сертификатах (с анти-спамом "не чаще раза в сутки").
        /// </summary>
        public void NotifyExpiringCerts(IEnumerable<CertRecord> records)
        {
            if (records == null) return;

            lock (_sync)
            {
                int threshold = _settings.NotifyDaysThreshold > 0
                    ? _settings.NotifyDaysThreshold
                    : 10;

                DateTime today = DateTime.Today;

                // Отбираем кандидатов
                var expiring = records
                    .Where(r => r != null
                                && !r.IsDeleted
                                && r.DaysLeft > 0
                                && r.DaysLeft <= threshold
                                && !string.IsNullOrWhiteSpace(r.Fio))
                    .ToList();

                if (expiring.Count == 0)
                    return;

                // Группируем по зданию
                var byBuilding = expiring
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.Building) ? "UNKNOWN" : r.Building)
                    .ToList();

                int totalSent = 0;

                foreach (var group in byBuilding)
                {
                    string building = group.Key;

                    var accounts = GetAccountsForBuilding(building);
                    if (accounts.Count == 0)
                    {
                        _log($"[Notify] Для здания '{building}' не настроены bimoid-аккаунты.");
                        continue;
                    }

                    // Фильтруем уже уведомлённые сегодня
                    var toNotify = new List<CertRecord>();

                    foreach (var rec in group)
                    {
                        string key = $"expiring:{rec.Id}";
                        if (_state.ExpiringLastSent.TryGetValue(key, out var last) &&
                            last.Date == today)
                            continue;

                        toNotify.Add(rec);
                    }

                    if (toNotify.Count == 0)
                        continue;

                    // Формируем ОДНО сообщение
                    var sb = new StringBuilder();
                    sb.AppendLine($"Истекают сертификаты (порог {threshold} дней):");
                    sb.AppendLine();

                    foreach (var rec in toNotify.OrderBy(r => r.DaysLeft))
                    {
                        sb.AppendLine(
                            $"• {rec.Fio} — осталось {rec.DaysLeft} дн. (до {rec.DateEnd:dd.MM.yyyy})");
                    }

                    string message = sb.ToString();

                    if (SendToAccounts(accounts, message))
                    {
                        foreach (var rec in toNotify)
                        {
                            _state.ExpiringLastSent[$"expiring:{rec.Id}"] = today;
                        }

                        totalSent += toNotify.Count;
                    }
                }

                if (totalSent > 0)
                {
                    _log($"[Notify] Отправлены уведомления об истекающих сертификатах: {totalSent} шт.");
                    SaveState();
                }
            }
        }

        /// <summary>
        /// Уведомления о НОВЫХ пользователях (анти-спам: не более одного раза в сутки на сертификат).
        /// </summary>
        public void NotifyNewUsers(IEnumerable<CertRecord> newRecords)
        {
            if (newRecords == null) return;

            lock (_sync)
            {
                DateTime today = DateTime.Today;

                var allAccounts = GetAllAccounts();
                if (allAccounts.Count == 0)
                {
                    _log("[Notify] Нет bimoid-аккаунтов для уведомлений о новых пользователях.");
                    return;
                }

                // ★ собираем новых пользователей
                var newUsers = new List<CertRecord>();

                foreach (var rec in newRecords)
                {
                    if (rec == null || string.IsNullOrWhiteSpace(rec.Fio))
                        continue;

                    string key = $"newuser:{rec.Id}";
                    if (_state.NewUserLastSent.TryGetValue(key, out var last) &&
                        last.Date == today)
                        continue;

                    newUsers.Add(rec);
                }

                if (newUsers.Count == 0)
                    return;

                // ★ формируем ОДНО сообщение
                var sb = new StringBuilder();
                sb.AppendLine("Добавлены новые пользователи:");
                sb.AppendLine();

                foreach (var rec in newUsers)
                {
                    sb.AppendLine($"• {rec.Fio} (с {rec.DateStart:dd.MM.yyyy} по {rec.DateEnd:dd.MM.yyyy})");
                }

                string message = sb.ToString();

                if (SendToAccounts(allAccounts, message))
                {
                    foreach (var rec in newUsers)
                    {
                        _state.NewUserLastSent[$"newuser:{rec.Id}"] = today;
                    }

                    _log($"[Notify] Отправлено уведомление о {newUsers.Count} новых пользователях.");
                    SaveState();
                }
            }
        }

        /// <summary>
        /// Тестовое обычное сообщение (проверка канала).
        /// </summary>
        public void SendTestMessage()
        {
            var allAccounts = GetAllAccounts();
            if (allAccounts.Count == 0)
            {
                _log("[Notify] Тест: нет аккаунтов для отправки.");
                return;
            }

            string msg = "Тестовое сообщение от ImapCertWatcher (проверка канала связи).";
            if (SendToAccounts(allAccounts, msg))
            {
                _log(string.Format("[Notify] Тестовое сообщение отправлено ({0} получателей).", allAccounts.Count));
            }
        }

        /// <summary>
        /// Тестовое уведомление о "новом пользователе" (для последнего/произвольного).
        /// </summary>
        public void SendTestNewUserNotification(CertRecord rec)
        {
            if (rec == null)
            {
                _log("[Notify] Тест нового пользователя: запись не передана.");
                return;
            }

            var allAccounts = GetAllAccounts();
            if (allAccounts.Count == 0)
            {
                _log("[Notify] Тест нового пользователя: нет аккаунтов.");
                return;
            }

            string msg = string.Format(
                "[ТЕСТ] Добавлен новый пользователь: {0}. Сертификат действует с {1:dd.MM.yyyy} по {2:dd.MM.yyyy}.",
                rec.Fio, rec.DateStart, rec.DateEnd);

            if (SendToAccounts(allAccounts, msg))
            {
                _log("[Notify] Тестовое уведомление о новом пользователе отправлено.");
            }
        }

        // ===== ВНУТРЕННИЕ ХЕЛПЕРЫ =====

        private List<string> GetAccountsForBuilding(string building)
        {
            string raw = null;

            if (string.Equals(building, "Краснофлотская", StringComparison.OrdinalIgnoreCase))
                raw = _settings.BimoidAccountsKrasnoflotskaya;
            else if (string.Equals(building, "Пионерская", StringComparison.OrdinalIgnoreCase))
                raw = _settings.BimoidAccountsPionerskaya;
            else
                return new List<string>();

            if (string.IsNullOrWhiteSpace(raw))
                return new List<string>();

            return raw
                .Split(new[] { '\r', '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(a => a.Trim())
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> GetAllAccounts()
        {
            var all = new List<string>();

            void AddFrom(string raw)
            {
                if (string.IsNullOrWhiteSpace(raw)) return;
                var parts = raw
                    .Split(new[] { '\r', '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(a => a.Trim());
                all.AddRange(parts);
            }

            AddFrom(_settings.BimoidAccountsKrasnoflotskaya);
            AddFrom(_settings.BimoidAccountsPionerskaya);

            return all
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private bool SendForBuilding(string building, string message)
        {
            var accounts = GetAccountsForBuilding(building);
            if (accounts.Count == 0)
            {
                _log(string.Format("[Notify] Для здания '{0}' не настроены bimoid-аккаунты.", building));
                return false;
            }

            return SendToAccounts(accounts, message);
        }

        private bool SendToAccounts(List<string> accounts, string message)
        {
            if (accounts == null || accounts.Count == 0)
                return false;

            try
            {
                if (!Directory.Exists(_obimpDir))
                {
                    _log(string.Format("[Notify] ObimpCMD не найден: {0}", _obimpDir));
                    return false;
                }

                string demoBat = Path.Combine(_obimpDir, "demo.bat");
                if (!File.Exists(demoBat))
                {
                    _log(string.Format("[Notify] demo.bat не найден: {0}", demoBat));
                    return false;
                }

                string accountsFile = Path.Combine(_obimpDir, "accounts.txt");
                string messageFile = Path.Combine(_obimpDir, "message.txt");

                File.WriteAllLines(accountsFile, accounts, Encoding.UTF8);
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

                _log(string.Format("[Notify] Отправлено через ObimpCMD ({0} получателей).", accounts.Count));
                return true;
            }
            catch (Exception ex)
            {
                _log(string.Format("[Notify] Ошибка отправки через ObimpCMD: {0}", ex.Message));
                return false;
            }
        }

        // ===== ЗАГРУЗКА / СОХРАНЕНИЕ СОСТОЯНИЯ (ПРОСТОЙ ТЕКСТ) =====

        private void LoadState()
        {
            try
            {
                if (!File.Exists(_stateFile))
                {
                    _state = new NotificationState();
                    return;
                }

                var state = new NotificationState();
                var lines = File.ReadAllLines(_stateFile, Encoding.UTF8);

                foreach (var rawLine in lines)
                {
                    var line = rawLine.Trim();
                    if (string.IsNullOrEmpty(line)) continue;

                    // Формат:
                    // E|key|2025-12-01T10:00:00.0000000
                    // N|key|2025-12-01T10:00:00.0000000
                    var parts = line.Split('|');
                    if (parts.Length != 3) continue;

                    string type = parts[0];
                    string key = parts[1];
                    string dtStr = parts[2];

                    if (!DateTime.TryParse(
                            dtStr,
                            null,
                            System.Globalization.DateTimeStyles.RoundtripKind,
                            out DateTime dt))
                        continue;

                    if (type == "E")
                    {
                        state.ExpiringLastSent[key] = dt;
                    }
                    else if (type == "N")
                    {
                        state.NewUserLastSent[key] = dt;
                    }
                }

                _state = state;
            }
            catch (Exception ex)
            {
                _log(string.Format("[Notify] Ошибка загрузки состояния уведомлений: {0}", ex.Message));
                _state = new NotificationState();
            }
        }

        private void SaveState()
        {
            try
            {
                var sb = new StringBuilder();

                foreach (var kv in _state.ExpiringLastSent)
                {
                    sb.Append("E|")
                      .Append(kv.Key)
                      .Append("|")
                      .Append(kv.Value.ToString("o"))
                      .AppendLine();
                }

                foreach (var kv in _state.NewUserLastSent)
                {
                    sb.Append("N|")
                      .Append(kv.Key)
                      .Append("|")
                      .Append(kv.Value.ToString("o"))
                      .AppendLine();
                }

                File.WriteAllText(_stateFile, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                _log(string.Format("[Notify] Ошибка сохранения состояния уведомлений: {0}", ex.Message));
            }
        }
    }
}
