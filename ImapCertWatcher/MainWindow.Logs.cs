using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private async Task LoadLogs()
        {
            try
            {
                if (_api != null)
                {
                    try
                    {
                        var serverLog = await _api.GetServerLog();

                        if (!string.IsNullOrWhiteSpace(serverLog))
                        {
                            txtLogs.Text = serverLog;
                            logStatusText.Text = "Серверный лог (текущая сессия)";
                            txtLogs.ScrollToEnd();

                            AddToMiniLog("Загружен серверный лог");
                            return;
                        }

                        AddToMiniLog("Серверный лог пустой, загружаем локальный лог клиента");
                    }
                    catch (Exception ex)
                    {
                        AddToMiniLog("Ошибка получения серверного лога: " + ex.Message);
                    }
                }

                LoadLocalClientLogs();
            }
            catch (Exception ex)
            {
                txtLogs.Text = $"Ошибка загрузки логов: {ex.Message}";
                logStatusText.Text = "Ошибка загрузки";
                AddToMiniLog($"Ошибка загрузки логов: {ex.Message}");
            }
        }

        private void LoadLocalClientLogs()
        {
            var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");

            if (!Directory.Exists(logBaseDirectory))
            {
                UpdateLogsUI(
                    "Папка локальных логов клиента не существует",
                    "Локальные логи клиента не найдены");
                return;
            }

            var dateFolders = Directory.GetDirectories(logBaseDirectory)
                .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(
                    Path.GetFileName(dir),
                    @"^\d{4}-\d{2}-\d{2}$"))
                .Select(dir => new DirectoryInfo(dir))
                .OrderByDescending(d => d.Name)
                .ToList();

            if (!dateFolders.Any())
            {
                UpdateLogsUI(
                    "Папки локальных логов клиента не найдены",
                    "Нет локальных логов клиента");
                return;
            }

            FileInfo latestSessionLog = null;

            foreach (var folder in dateFolders)
            {
                var sessionLogs = folder.GetFiles("session_*.log")
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                if (sessionLogs.Any())
                {
                    latestSessionLog = sessionLogs.First();
                    break;
                }
            }

            if (latestSessionLog == null)
            {
                UpdateLogsUI(
                    "Файлы локальных логов клиента не найдены",
                    "Нет локальных логов клиента");
                return;
            }

            var lines = File.ReadAllLines(latestSessionLog.FullName);
            var lastLines = lines.Skip(Math.Max(0, lines.Length - 1000));

            txtLogs.Text = string.Join(Environment.NewLine, lastLines);
            logStatusText.Text = $"Локальный лог клиента ({latestSessionLog.LastWriteTime:dd.MM.yyyy HH:mm})";
            txtLogs.ScrollToEnd();

            AddToMiniLog("Загружен локальный лог клиента");
        }
        private void UpdateLogsUI(string logText, string statusTextValue)
        {
            txtLogs.Text = logText;
            logStatusText.Text = statusTextValue;
        }

        private void CleanOldLogs()
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                    return;

                var cutoffDate = DateTime.Now.AddMonths(-2);

                var dateFolders = Directory.GetDirectories(logBaseDirectory)
                    .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(
                        Path.GetFileName(dir),
                        @"^\d{4}-\d{2}-\d{2}$"));

                foreach (var dateFolder in dateFolders)
                {
                    var folderName = Path.GetFileName(dateFolder);

                    if (DateTime.TryParseExact(
                        folderName,
                        "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out DateTime folderDate))
                    {
                        if (folderDate < cutoffDate.Date)
                        {
                            try
                            {
                                Directory.Delete(dateFolder, true);
                            }
                            catch
                            {
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private async void LogsTab_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is TabControl tabControl &&
                tabControl.SelectedItem is TabItem selectedTab &&
                selectedTab.Header?.ToString() == "Логи")
            {
                await LoadLogs();
                AddToMiniLog("Автообновление логов при переходе на вкладку");
            }
        }

        private async void BtnRefreshLogs_Click(object sender, RoutedEventArgs e)
        {
            await LoadLogs();
            AddToMiniLog("Логи обновлены вручную");
        }

        private async void BtnClearLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                {
                    MessageBox.Show(
                        "Папка логов не существует",
                        "Информация",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var cutoffDate = DateTime.Now.AddMonths(-2);
                int deletedCount = 0;

                var dateFolders = Directory.GetDirectories(logBaseDirectory)
                    .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(
                        Path.GetFileName(dir),
                        @"^\d{4}-\d{2}-\d{2}$"));

                foreach (var dateFolder in dateFolders)
                {
                    var folderName = Path.GetFileName(dateFolder);

                    if (DateTime.TryParseExact(
                        folderName,
                        "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out DateTime folderDate))
                    {
                        if (folderDate < cutoffDate.Date)
                        {
                            try
                            {
                                Directory.Delete(dateFolder, true);
                                deletedCount++;
                                AddToMiniLog($"Удалена папка логов: {folderName}");
                            }
                            catch (Exception ex)
                            {
                                AddToMiniLog($"Ошибка удаления папки {folderName}: {ex.Message}");
                            }
                        }
                    }
                }

                MessageBox.Show(
                    $"Удалено папок с логами: {deletedCount}",
                    "Очистка логов",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                await LoadLogs();
                AddToMiniLog($"Очищено логов: {deletedCount} папок");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при очистке логов: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                AddToMiniLog($"Ошибка очистки логов: {ex.Message}");
            }
        }

        private void BtnOpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                    Directory.CreateDirectory(logBaseDirectory);

                Process.Start(new ProcessStartInfo("explorer.exe", logBaseDirectory)
                {
                    UseShellExecute = true
                });

                AddToMiniLog("Открыта локальная папка логов клиента");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при открытии папки логов: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                AddToMiniLog($"Ошибка открытия папки логов: {ex.Message}");
            }
        }
    }
}