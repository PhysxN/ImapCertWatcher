using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ImapCertWatcher.Services
{
    public class NotificationManager
    {
        private readonly ServerSettings _settings;
        private readonly Action<string> _log;
        private readonly string _stateFile;
        private readonly object _sync = new object();

        private class NotificationState
        {
            public Dictionary<string, DateTime> ExpiringLastSent { get; set; }
            public Dictionary<string, DateTime> NewUserLastSent { get; set; }

            public NotificationState()
            {
                ExpiringLastSent = new Dictionary<string, DateTime>();
                NewUserLastSent = new Dictionary<string, DateTime>();
            }
        }

        private class BroadcastResultItemFile
        {
            public string recipient { get; set; }
            public string status { get; set; }
            public string message { get; set; }

            public BroadcastResultItemFile()
            {
                recipient = string.Empty;
                status = string.Empty;
                message = string.Empty;
            }
        }

        private class BroadcastResultFile
        {
            public string startedAt { get; set; }
            public string finishedAt { get; set; }
            public bool isSuccess { get; set; }
            public string summaryMessage { get; set; }
            public List<BroadcastResultItemFile> items { get; set; }

            public BroadcastResultFile()
            {
                startedAt = string.Empty;
                finishedAt = string.Empty;
                summaryMessage = string.Empty;
                items = new List<BroadcastResultItemFile>();
            }
        }

        private NotificationState _state = new NotificationState();

        public NotificationManager(ServerSettings settings, Action<string> log)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _log = log ?? (_ => { });

            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var logDir = Path.Combine(baseDir, "LOG");

            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            _stateFile = Path.Combine(logDir, "notifications_state.txt");
            LoadState();
        }

        public void NotifyExpiringCerts(IEnumerable<CertRecord> records)
        {
            if (records == null) return;

            if (!CanSendNow("истекающие сертификаты"))
                return;

            lock (_sync)
            {
                int threshold = _settings.NotifyDaysThreshold > 0
                    ? _settings.NotifyDaysThreshold
                    : 10;

                DateTime today = DateTime.Today;

                var expiring = records
                    .Where(r => r != null
                                && !r.IsDeleted
                                && r.DaysLeft >= 0
                                && r.DaysLeft <= threshold
                                && !string.IsNullOrWhiteSpace(r.Fio))
                    .ToList();

                if (expiring.Count == 0)
                    return;

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

        public void NotifyNewUsers(IEnumerable<CertRecord> changedRecords)
        {
            if (changedRecords == null) return;

            if (!CanSendNow("изменения по сертификатам"))
                return;

            lock (_sync)
            {
                DateTime today = DateTime.Today;

                var allAccounts = GetAllAccounts();
                if (allAccounts.Count == 0)
                {
                    _log("[Notify] Нет bimoid-аккаунтов для уведомлений об изменениях сертификатов.");
                    return;
                }

                var changed = new List<CertRecord>();

                foreach (var rec in changedRecords)
                {
                    if (rec == null
                        || rec.IsDeleted
                        || string.IsNullOrWhiteSpace(rec.Fio))
                        continue;

                    string key = $"newuser:{rec.Id}";
                    if (_state.NewUserLastSent.TryGetValue(key, out var last) &&
                        last.Date == today)
                        continue;

                    changed.Add(rec);
                }

                if (changed.Count == 0)
                    return;

                var sb = new StringBuilder();
                sb.AppendLine("Изменения по сертификатам:");
                sb.AppendLine();

                foreach (var rec in changed)
                {
                    sb.AppendLine(
                        $"• {rec.Fio} — сертификат {rec.CertNumber}, действует с {rec.DateStart:dd.MM.yyyy} по {rec.DateEnd:dd.MM.yyyy}");
                }

                string message = sb.ToString();

                if (SendToAccounts(allAccounts, message))
                {
                    foreach (var rec in changed)
                    {
                        _state.NewUserLastSent[$"newuser:{rec.Id}"] = today;
                    }

                    _log($"[Notify] Отправлено уведомление об изменениях сертификатов: {changed.Count} шт.");
                    SaveState();
                }
            }
        }

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

        public void SendTestNewUserNotification(CertRecord rec)
        {
            if (rec == null)
            {
                _log("[Notify] Тест изменений сертификата: запись не передана.");
                return;
            }

            var allAccounts = GetAllAccounts();
            if (allAccounts.Count == 0)
            {
                _log("[Notify] Тест изменений сертификата: нет аккаунтов.");
                return;
            }

            string msg = string.Format(
                "[ТЕСТ] Изменения по сертификату: {0}. Сертификат {1}, действует с {2:dd.MM.yyyy} по {3:dd.MM.yyyy}.",
                rec.Fio, rec.CertNumber, rec.DateStart, rec.DateEnd);

            if (SendToAccounts(allAccounts, msg))
            {
                _log("[Notify] Тестовое уведомление об изменениях сертификата отправлено.");
            }
        }

        private List<string> GetAccountsForBuilding(string building)
        {
            building = building?.Trim();
            string raw = null;

            if (string.Equals(building, "Краснофлотская", StringComparison.OrdinalIgnoreCase))
                raw = _settings.BimoidAccountsKrasnoflotskaya;
            else if (string.Equals(building, "Пионерская", StringComparison.OrdinalIgnoreCase))
                raw = _settings.BimoidAccountsPionerskaya;
            else
                return new List<string>();

            return SplitAccounts(raw);
        }

        private List<string> GetAllAccounts()
        {
            var all = new List<string>();

            all.AddRange(SplitAccounts(_settings.BimoidAccountsKrasnoflotskaya));
            all.AddRange(SplitAccounts(_settings.BimoidAccountsPionerskaya));

            return all
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> SplitAccounts(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return new List<string>();

            return raw
                .Split(new[] { '\r', '\n', ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(a => a.Trim())
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private bool CanSendNow(string kindForLog)
        {
            if (!_settings.NotifyOnlyInWorkHours)
                return true;

            DateTime now = DateTime.Now;
            DayOfWeek day = now.DayOfWeek;

            bool isWorkDay =
                day != DayOfWeek.Saturday &&
                day != DayOfWeek.Sunday;

            if (!isWorkDay)
            {
                _log($"[Notify] Уведомление '{kindForLog}' не отправлено: сейчас нерабочий день ({now:dd.MM.yyyy HH:mm:ss}).");
                return false;
            }

            TimeSpan start = new TimeSpan(7, 30, 0);
            TimeSpan end = new TimeSpan(17, 0, 0);
            TimeSpan current = now.TimeOfDay;

            if (current < start || current > end)
            {
                _log($"[Notify] Уведомление '{kindForLog}' не отправлено: сейчас вне рабочего времени ({now:dd.MM.yyyy HH:mm:ss}), разрешено Пн-Пт с 07:30 до 17:00.");
                return false;
            }

            return true;
        }

        private bool SendToAccounts(List<string> accounts, string message)
        {
            if (accounts == null || accounts.Count == 0)
                return false;

            try
            {
                accounts = accounts
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .Select(a => a.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (accounts.Count == 0)
                    return false;

                string senderExePath = ResolvePath(
                    _settings.BimoidSenderExePath,
                    @"BimoidBroadcastSender\BimoidBroadcastSender.exe");

                string jobDirectory = ResolvePath(
                    _settings.BimoidJobDirectory,
                    "BimoidJobs");

                if (!File.Exists(senderExePath))
                {
                    _log($"[Notify] BimoidBroadcastSender.exe не найден: {senderExePath}");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(_settings.BimoidServer))
                {
                    _log("[Notify] Не задан BimoidServer.");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(_settings.BimoidLogin))
                {
                    _log("[Notify] Не задан BimoidLogin.");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(_settings.BimoidPassword))
                {
                    _log("[Notify] Не задан BimoidPassword.");
                    return false;
                }

                if (_settings.BimoidPort <= 0 || _settings.BimoidPort > 65535)
                {
                    _log("[Notify] Указан некорректный BimoidPort.");
                    return false;
                }

                Directory.CreateDirectory(jobDirectory);
                CleanupOldJobFiles(jobDirectory);

                string jobId = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                string jobFilePath = Path.Combine(jobDirectory, "job_" + jobId + ".json");
                string resultFilePath = Path.Combine(jobDirectory, "result_" + jobId + ".json");

                var jobObject = new
                {
                    server = _settings.BimoidServer.Trim(),
                    port = _settings.BimoidPort,
                    login = (_settings.BimoidLogin ?? string.Empty).Trim(),
                    password = _settings.BimoidPassword ?? string.Empty,
                    recipients = accounts,
                    messageText = message ?? string.Empty,
                    delayBetweenMessagesMs = _settings.BimoidDelayBetweenMessagesMs < 0 ? 0 : _settings.BimoidDelayBetweenMessagesMs,
                    resultFilePath = resultFilePath
                };

                string jobJson = JsonConvert.SerializeObject(jobObject, Formatting.Indented);
                File.WriteAllText(jobFilePath, jobJson, Encoding.UTF8);

                var psi = new ProcessStartInfo
                {
                    FileName = senderExePath,
                    Arguments = "\"" + jobFilePath + "\"",
                    WorkingDirectory = Path.GetDirectoryName(senderExePath),
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                int timeoutMs = 90000 + (accounts.Count * Math.Max(_settings.BimoidDelayBetweenMessagesMs, 0));

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        _log("[Notify] Не удалось запустить BimoidBroadcastSender.");
                        return false;
                    }

                    if (!proc.WaitForExit(timeoutMs))
                    {
                        _log("[Notify] BimoidBroadcastSender завис (таймаут ожидания).");
                        try { proc.Kill(); } catch { }
                        return false;
                    }

                    if (!File.Exists(resultFilePath))
                    {
                        _log($"[Notify] Не найден result json: {resultFilePath}");
                        return false;
                    }

                    BroadcastResultFile result;
                    try
                    {
                        string resultJson = File.ReadAllText(resultFilePath, Encoding.UTF8);
                        result = JsonConvert.DeserializeObject<BroadcastResultFile>(resultJson) ?? new BroadcastResultFile();
                    }
                    catch (Exception ex)
                    {
                        _log("[Notify] Не удалось прочитать result json: " + ex.Message);
                        return false;
                    }

                    if (!string.IsNullOrWhiteSpace(result.summaryMessage))
                    {
                        _log("[Notify] " + result.summaryMessage);
                    }

                    if (result.items != null)
                    {
                        foreach (var item in result.items
                            .Where(i => !string.Equals(i.status, "SUCCESS", StringComparison.OrdinalIgnoreCase))
                            .Take(20))
                        {
                            _log($"[Notify] Получатель='{item.recipient}', статус='{item.status}', сообщение='{item.message}'");
                        }
                    }

                    if (proc.ExitCode != 0)
                    {
                        _log("[Notify] BimoidBroadcastSender завершился с кодом: " + proc.ExitCode);
                    }

                    if (!result.isSuccess)
                        return false;
                }

                _log(string.Format("[Notify] Отправлено через BimoidBroadcastSender ({0} получателей).", accounts.Count));
                return true;
            }
            catch (Exception ex)
            {
                _log(string.Format("[Notify] Ошибка отправки через BimoidBroadcastSender: {0}", ex.Message));
                return false;
            }
        }

        private string ResolvePath(string configuredPath, string defaultRelativePath)
        {
            string path = string.IsNullOrWhiteSpace(configuredPath)
                ? defaultRelativePath
                : configuredPath.Trim();

            if (!Path.IsPathRooted(path))
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
            }

            return Path.GetFullPath(path);
        }

        private void CleanupOldJobFiles(string directory)
        {
            try
            {
                if (!Directory.Exists(directory))
                    return;

                DateTime border = DateTime.Now.AddDays(-30);

                foreach (var file in Directory.GetFiles(directory, "*.json"))
                {
                    try
                    {
                        DateTime lastWrite = File.GetLastWriteTime(file);
                        if (lastWrite < border)
                            File.Delete(file);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

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