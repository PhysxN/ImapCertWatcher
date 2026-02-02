using DocumentFormat.OpenXml.Office2010.Excel;
using ImapCertWatcher;
using ImapCertWatcher.Client;
using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.UI;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
// Альтернатива Excel без использования Office Interop
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Forms = System.Windows.Forms;
using Newtonsoft.Json;
using System.IO.Compression;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window
    {
        private ClientSettings _clientSettings;
        private ServerSettings _serverSettings;
        private NotificationManager _notificationManager;
        private HashSet<int> _knownCertIds = new HashSet<int>();
        private ServerApiClient _api;
        private ObservableCollection<CertRecord> _items = new ObservableCollection<CertRecord>();
        private ObservableCollection<CertRecord> _allItems = new ObservableCollection<CertRecord>();
        private DispatcherTimer _refreshTimer;
        private ObservableCollection<string> _miniLogMessages = new ObservableCollection<string>();
        private const int MAX_MINI_LOG_LINES = 3; 
        private string _currentBuildingFilter = "Все";
        private bool _showDeleted = false;
        private string _searchText = "";
        private bool _isDarkTheme = false;
        private const string AUTOSTART_REG_PATH = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AUTOSTART_VALUE_NAME = "ImapCertWatcher";
        
        private DispatcherTimer _serverMonitorTimer;
        private DispatcherTimer _connectingAnimationTimer;
        private int _connectingDots = 0;
        private bool _reconnectInProgress = false;
        private bool _isLoadingFromServer = false;

        enum ServerConnectionState
        {
            Connecting,
            Online,
            Busy,
            Offline
        }
        private int _reconnectDelaySeconds = 2;
        private const int MAX_RECONNECT_DELAY = 30;
        private ServerConnectionState _serverState = ServerConnectionState.Offline;
        private string _lastTimerValue = "";
        private void ApplyAutoStartSetting()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(AUTOSTART_REG_PATH, true))
                {
                    if (key == null) return;

                    string exePath = Assembly.GetExecutingAssembly().Location;
                    exePath = "\"" + exePath + "\"";

                    if (_clientSettings.AutoStart)
                    {
                        key.SetValue(AUTOSTART_VALUE_NAME, exePath);
                        Log("Автозапуск включен");
                    }
                    else
                    {
                        key.DeleteValue(AUTOSTART_VALUE_NAME, false);
                        Log("Автозапуск выключен");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка настройки автозапуска: {ex.Message}");
            }
        }

        // Список доступных зданий
        public ObservableCollection<string> AvailableBuildings { get; } =
                new ObservableCollection<string>
                {
                    "", "Краснофлотская", "Пионерская"
                };

        

        // Флаг для предотвращения множественного сохранения
        private bool _isSavingBuilding = false;

        // Событие для обновления прогресса splash screen
        public event Action<string, double> ProgressUpdated;

        public MainWindow()
        {
            
            try
            {
                InitializeComponent();

                                // Инициализация мини-лога
                _miniLogMessages = new ObservableCollection<string>();
                AddToMiniLog("Приложение запущено");

                // Очистка старых логов при запуске (в фоне)
                Task.Run(() => CleanOldLogs());

                // Загружаем настройки
                var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
                _clientSettings = SettingsLoader.LoadClient(settingsPath);
                _serverSettings = SettingsLoader.LoadServer(settingsPath);

                this.DataContext = _clientSettings;

                ApplyAutoStartSetting();

                _notificationManager = new NotificationManager(_serverSettings, Log);



                // Загрузка темы
                LoadThemeSettings();

                // Настройка DataGrid и обработчиков
                dgCerts.CanUserAddRows = false;
                dgCerts.ItemsSource = _items;
                dgCerts.SelectionChanged += (s, e) =>
                {
                    bool hasSelection = dgCerts.SelectedItem != null;

                    btnDeleteSelected.IsEnabled = hasSelection;
                    btnRestoreSelected.IsEnabled = hasSelection;
                };

                // стартовое состояние
                btnDeleteSelected.IsEnabled = false;
                btnRestoreSelected.IsEnabled = false;

                dgCerts.CellEditEnding += DgCerts_CellEditEnding;
                dgCerts.BeginningEdit += DgCerts_BeginningEdit;
                dgCerts.PreparingCellForEdit += DgCerts_PreparingCellForEdit;

                this.Loaded += MainWindow_Loaded;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации приложения: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }
        }

        private bool IsWorkingTime()
        {
            var now = DateTime.Now;

            // Суббота/воскресенье — не рабочие
            if (now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday)
                return false;

            var t = now.TimeOfDay;
            var start = new TimeSpan(8, 0, 0);   // 08:00
            var end = new TimeSpan(16, 30, 0); // 16:30

            return t >= start && t <= end;
        }

        


        private void SetServerState(ServerConnectionState state)
        {
            if (_serverState == state && state != ServerConnectionState.Connecting)
    return;
            _serverState = state;

            Dispatcher.Invoke(() =>
            {
                if (state != ServerConnectionState.Busy)
                {
                    pbServerProgress.Visibility = Visibility.Collapsed;
                }
                if (state == ServerConnectionState.Connecting)
                    StartConnectingAnimation();
                else
                    StopConnectingAnimation();

                switch (state)
                {

                    case ServerConnectionState.Connecting:
                        txtServerStatus.Text = "🟡 Подключение к серверу...";
                        txtServerStatus.Foreground = Brushes.Gold;
                        break;

                    case ServerConnectionState.Online:
                        txtServerStatus.Text = "🟢 Сервер готов";
                        txtServerStatus.Foreground = Brushes.LimeGreen;
                        break;

                    case ServerConnectionState.Busy:
                        txtServerStatus.Text = "🟠 Сервер работает";
                        txtServerStatus.Foreground = Brushes.Orange;
                        break;

                    case ServerConnectionState.Offline:
                        txtServerStatus.Text = "🔴 Сервер OFFLINE";
                        txtServerStatus.Foreground = Brushes.Red;
                        break;
                }
            });
        }

        class ServerApiClient
        {
            private ClientSettings _settings;

            public ServerApiClient(ClientSettings settings)
            {
                _settings = settings;
            }

            public async Task<List<CertRecord>> GetCertificates()
            {
                var response = await TcpCommandClient.SendAsync(
                    _settings.ServerIp,
                    _settings.ServerPort,
                    "GET_CERTS");

                if (string.IsNullOrWhiteSpace(response))
                    return new List<CertRecord>();

                if (response.StartsWith("CERTS "))
                    response = response.Substring(6);

                // Убираем мусор перевода строки
                response = response.Trim();

                return JsonConvert.DeserializeObject<List<CertRecord>>(response);
            }

            public async Task UpdateBuilding(int id, string building)
            {
                await TcpCommandClient.SendAsync(
                    _settings.ServerIp,
                    _settings.ServerPort,
                    $"SET_BUILDING|{id}|{building}");
            }

            public async Task MarkDeleted(int id, bool deleted)
            {
                await TcpCommandClient.SendAsync(
                    _settings.ServerIp,
                    _settings.ServerPort,
                    $"MARK_DELETED|{id}|{deleted}");
            }
        }
        private void StartServerMonitor()
        {
            _serverMonitorTimer = new DispatcherTimer();
            _serverMonitorTimer.Interval = TimeSpan.FromSeconds(2);
            _serverMonitorTimer.Tick += async (s, e) =>
            {
                await UpdateServerMonitor();
            };

            _serverMonitorTimer.Start();
        }

        private async Task UpdateServerMonitor()
        {
            try
            {
                if (_serverState == ServerConnectionState.Offline)
                {
                    SetServerState(ServerConnectionState.Connecting);
                }

                var status = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    "STATUS");

                var progress = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    "GET_PROGRESS");

                var timer = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    "GET_TIMER");

                var log = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    "GET_LOG");

                // ✅ только если реально получили ответ
                if (!string.IsNullOrWhiteSpace(status))
                {
                    UpdateStatusUI(status);
                    UpdateProgressUI(progress);
                    UpdateTimerUI(timer);
                    UpdateServerLog(log);
                }

                // reset reconnect backoff on success
                _reconnectDelaySeconds = 2;
            }
            catch (Exception ex)
            {
                Log("Server monitor error: " + ex.Message);

                SetServerState(ServerConnectionState.Offline);

                Dispatcher.Invoke(() =>
                {
                    pbServerProgress.Visibility = Visibility.Collapsed;
                });

                StartReconnectBackoff();
            }
        }

        private void UpdateServerLog(string log)
        {
            if (!log.StartsWith("LOG"))
                return;

            var content = log.Replace("LOG ", "");

            AddToMiniLog("[SERVER] " + content);
        }

        private void UpdateStatusUI(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return;

            status = status.ToUpperInvariant();

            // DEBUG (очень полезно)
            AddToMiniLog("[STATUS RAW] " + status);

            if (status.Contains("BUSY"))
            {
                SetServerState(ServerConnectionState.Busy);
            }
            else if (status.Contains("IDLE") || status.Contains("READY") || status.Contains("OK"))
            {
                SetServerState(ServerConnectionState.Online);
            }
            else
            {
                SetServerState(ServerConnectionState.Offline);
            }
        }

        private void UpdateProgressUI(string progress)
        {
            if (!progress.StartsWith("PROGRESS"))
                return;

            var data = progress.Replace("PROGRESS ", "").Split('|');

            if (!int.TryParse(data[0], out int percent))
                return;

            Dispatcher.Invoke(() =>
            {
                // показываем только если сервер реально BUSY
                if (_serverState == ServerConnectionState.Busy)
                {
                    pbServerProgress.IsIndeterminate = false;
                    pbServerProgress.Visibility = Visibility.Visible;
                    pbServerProgress.Value = percent;
                }
                else
                {
                    pbServerProgress.Visibility = Visibility.Collapsed;
                }
            });
        }

        private void UpdateTimerUI(string timer)
        {
            if (string.IsNullOrWhiteSpace(timer))
                return;
            if (timer == _lastTimerValue)
                return;

            _lastTimerValue = timer;

            Dispatcher.Invoke(() =>
            {
                if (timer.Contains("READY"))
                {
                    txtNextCheck.Text = "Следующая проверка: сейчас";
                }
                else if (timer.StartsWith("TIMER"))
                {
                    var min = timer.Replace("TIMER ", "").Trim();
                    txtNextCheck.Text = $"Следующая проверка через {min} мин";
                }
            });
        }





        private void StartConnectingAnimation()
        {
            if (_connectingAnimationTimer == null)
                _connectingAnimationTimer = new DispatcherTimer();

            _connectingAnimationTimer.Interval = TimeSpan.FromMilliseconds(400);

            _connectingAnimationTimer.Tick -= ConnectingAnimationTick;
            _connectingAnimationTimer.Tick += ConnectingAnimationTick;

            if (!_connectingAnimationTimer.IsEnabled)
                _connectingAnimationTimer.Start();
        }

        private void ConnectingAnimationTick(object sender, EventArgs e)
        {
            if (_serverState != ServerConnectionState.Connecting)
                return;

            _connectingDots = (_connectingDots + 1) % 4;

            txtServerStatus.Text =
                "🟡 Подключение к серверу" + new string('.', _connectingDots);
        }

        private void StopConnectingAnimation()
        {
            if (_connectingAnimationTimer != null)
            {
                _connectingAnimationTimer.Stop();
                _connectingAnimationTimer.Tick -= ConnectingAnimationTick;
                _connectingDots = 0;
            }
        }

        private void OnProgressUpdated(string message, double progress)
        {
            ProgressUpdated?.Invoke(message, progress);
        }

        public event Action DataLoaded;

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Очистка папки Certs старше 48 часов (в фоне)
            _ = CleanupCertsFolderAsync();

            this.Loaded -= MainWindow_Loaded;

            System.Diagnostics.Debug.WriteLine("MainWindow_Loaded начат");

            OnProgressUpdated("Инициализация базы данных и почтовых модулей...", 10);

            try
            {
                // ТЯЖЁЛАЯ ЧАСТЬ – в фоновом потоке
                _api = new ServerApiClient(_clientSettings);

                await LoadFromServer();

                System.Diagnostics.Debug.WriteLine("Фоновая инициализация завершена");

                System.Diagnostics.Debug.WriteLine($"Загружено записей: {_allItems.Count}");

                // Запоминаем уже известные сертификаты (для определения "новых владельцев")
                _knownCertIds = new HashSet<int>(_allItems.Select(r => r.Id));

                OnProgressUpdated("Применение фильтров...", 85);
                ApplySearchFilter();

                OnProgressUpdated("Настройка интерфейса...", 90);
                AutoFitDataGridColumns();

                
                // Таймер для обновления дней
                InitializeRefreshTimer();

                OnProgressUpdated("Загрузка логов...", 95);
                LoadLogs();

                // ★ ВЫЗЫВАЕМ СОБЫТИЕ ПОЛНОЙ ЗАГРУЗКИ ДАННЫХ ★
                OnProgressUpdated("Готово", 100);

                System.Diagnostics.Debug.WriteLine("Вызываю DataLoaded...");

                // ★ УБИРАЕМ ЗАДЕРЖКУ - ЗАКРЫВАЕМ СРАЗУ ★
                DataLoaded?.Invoke();                
                StartServerMonitor();
                // Инициализируем менеджер уведомлений после получения настроек
                statusText.Text = "Готово";
                AddToMiniLog($"Инициализация завершена. Загружено записей: {_allItems.Count}");

                System.Diagnostics.Debug.WriteLine("MainWindow_Loaded завершен");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в MainWindow_Loaded: {ex.Message}");

                OnProgressUpdated($"Ошибка: {ex.Message}", -1);
                statusText.Text = "Ошибка инициализации";
                AddToMiniLog($"Ошибка инициализации: {ex.Message}");
                MessageBox.Show($"Ошибка инициализации: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                // Даже при ошибке уведомляем о завершении
                DataLoaded?.Invoke();
                StartServerMonitor();
            }
        }

        /// <summary>
        /// Инициализация таймера для авто-обновления дней
        /// </summary>
        private void InitializeRefreshTimer()
        {
            try
            {
                _refreshTimer = new DispatcherTimer();
                _refreshTimer.Interval = TimeSpan.FromMinutes(10); // Обновляем каждые 10 минут
                _refreshTimer.Tick += RefreshDaysLeft;
                _refreshTimer.Start();

                AddToMiniLog("Таймер обновления дней запущен (интервал: 10 минут)");
            }
            catch (Exception ex)
            {
                Log($"Ошибка инициализации таймера обновления: {ex.Message}");
            }
        }

        /// <summary>
        /// Обновление отображения оставшихся дней для видимых записей в DataGrid
        /// </summary>
        private void RefreshDaysLeft(object sender, EventArgs e)
        {
            try
            {
                // Оптимизация: обновляем только видимые элементы в DataGrid
                // Это значительно снижает нагрузку при большом количестве записей
                
                if (dgCerts != null && dgCerts.ItemsSource is IEnumerable<CertRecord> visibleItems)
                {
                    // Обновляем только видимые элементы
                    foreach (var item in visibleItems)
                    {
                        if (item != null)
                        {
                            item.RefreshDaysLeft();
                        }
                    }
                }

                // Также обновляем отфильтрованные элементы в памяти (но не все сразу)
                // Это гарантирует, что данные будут актуальны для поиска/фильтрации
                if (_items.Count > 0)
                {
                    // Обновляем только первые 100 элементов за раз, чтобы не зависнуть
                    int updateCount = Math.Min(100, _items.Count);
                    for (int i = 0; i < updateCount; i++)
                    {
                        if (_items[i] != null)
                        {
                            _items[i].RefreshDaysLeft();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при обновлении дней: {ex.Message}");
            }
        }

        /// <summary>
        /// Ручное обновление дней (можно вызвать из кода)
        /// </summary>
        private void ManualRefreshDaysLeft()
        {
            try
            {
                RefreshDaysLeft(null, EventArgs.Empty);
                AddToMiniLog("Ручное обновление дней выполнено");
            }
            catch (Exception ex)
            {
                Log($"Ошибка при ручном обновлении дней: {ex.Message}");
            }
        }

        private void AddToMiniLog(string message)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    // Добавляем новое сообщение
                    _miniLogMessages.Insert(0, $"{DateTime.Now:HH:mm:ss} - {message}");

                    // Ограничиваем количество строк
                    while (_miniLogMessages.Count > MAX_MINI_LOG_LINES)
                    {
                        _miniLogMessages.RemoveAt(_miniLogMessages.Count - 1);
                    }

                    // Обновляем оба мини-лога
                    UpdateMiniLogDisplay();
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в AddToMiniLog: {ex.Message}");
            }
        }

        private void UpdateMiniLogDisplay()
        {
            try
            {
                string logText = string.Join(Environment.NewLine, _miniLogMessages);

                // Обновляем мини-лог на основной вкладке
                if (txtMiniLog != null)
                {
                    txtMiniLog.Text = logText;
                }

                // Обновляем мини-лог на вкладке настроек
                if (txtMiniLogSettings != null)
                {
                    txtMiniLogSettings.Text = logText;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка обновления мини-лога: {ex.Message}");
            }
        }

        private void LoadThemeSettings()
        {
            try
            {
                var themeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "theme.txt");
                if (File.Exists(themeFile))
                {
                    var theme = File.ReadAllText(themeFile).Trim();
                    _isDarkTheme = theme.Equals("dark", StringComparison.OrdinalIgnoreCase);
                    chkDarkTheme.IsChecked = _isDarkTheme;
                    ApplyTheme(_isDarkTheme);
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка загрузки настроек темы: {ex.Message}");
            }
        }

        private void SaveThemeSettings()
        {
            try
            {
                var themeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "theme.txt");
                File.WriteAllText(themeFile, _isDarkTheme ? "dark" : "light");
            }
            catch (Exception ex)
            {
                Log($"Ошибка сохранения настроек темы: {ex.Message}");
            }
        }

        private void ApplyTheme(bool dark)
        {
            var res = Application.Current.Resources;

            if (dark)
            {
                res["WindowBackgroundBrush"] = res["DarkWindowBackground"];
                res["ControlBackgroundBrush"] = res["DarkControlBackground"];
                res["TextColorBrush"] = res["DarkTextColor"];
                res["BorderColorBrush"] = res["DarkBorderColor"];
                res["AccentColorBrush"] = res["DarkAccentColor"];
                res["HoverColorBrush"] = res["DarkHoverColor"];
            }
            else
            {
                res["WindowBackgroundBrush"] = res["LightWindowBackground"];
                res["ControlBackgroundBrush"] = res["LightControlBackground"];
                res["TextColorBrush"] = res["LightTextColor"];
                res["BorderColorBrush"] = res["LightBorderColor"];
                res["AccentColorBrush"] = res["LightAccentColor"];
                res["HoverColorBrush"] = res["LightHoverColor"];
            }
        }

        private void ChkDarkTheme_Changed(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = chkDarkTheme.IsChecked == true;

            ApplyTheme(_isDarkTheme);
            SaveThemeSettings();

            Log($"[UI] Тема изменена: {(_isDarkTheme ? "Тёмная" : "Светлая")}");
        }

        // ========== ЭКСПОРТ В EXCEL С ИСПОЛЬЗОВАНИЕМ CSV ==========

        private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            await ExportToExcel(false);
        }

        private async void ExportSelectedMenuItem_Click(object sender, RoutedEventArgs e)
        {
            await ExportToExcel(true);
        }

        private void DgCerts_Loaded(object sender, RoutedEventArgs e)
        {
            AutoFitDataGridColumns();
        }

        private void AutoFitDataGridColumns()
        {
            try
            {
                if (dgCerts == null || dgCerts.Columns == null) return;

                // Ждем немного чтобы DataGrid полностью отрисовался
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        foreach (var column in dgCerts.Columns)
                        {
                            // Для каждого типа столбцов свой подход
                            if (column is DataGridTextColumn textColumn)
                            {
                                AutoFitTextColumn(textColumn);
                            }
                            else if (column is DataGridTemplateColumn templateColumn)
                            {
                                AutoFitTemplateColumn(templateColumn);
                            }
                            else if (column is DataGridCheckBoxColumn checkBoxColumn)
                            {
                                AutoFitCheckBoxColumn(checkBoxColumn);
                            }
                        }

                        // Принудительно обновляем layout
                        dgCerts.UpdateLayout();
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка автоподбора столбцов: {ex.Message}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                Log($"Ошибка в AutoFitDataGridColumns: {ex.Message}");
            }
        }

        private void AutoFitTextColumn(DataGridTextColumn column)
        {
            try
            {
                // Сбрасываем ширину чтобы применить авто-размер
                column.Width = DataGridLength.Auto;

                // Принудительно обновляем
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора TextColumn: {ex.Message}");
            }
        }

        private void AutoFitTemplateColumn(DataGridTemplateColumn column)
        {
            try
            {
                // Для TemplateColumn используем авто-размер
                column.Width = DataGridLength.Auto;
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора TemplateColumn: {ex.Message}");
            }
        }

        private void AutoFitCheckBoxColumn(DataGridCheckBoxColumn column)
        {
            try
            {
                // Для CheckBoxColumn фиксированная ширина
                column.Width = new DataGridLength(80); // Фиксированная ширина для чекбокса
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора CheckBoxColumn: {ex.Message}");
            }
        }

        private async Task ExportToExcel(bool exportSelectedOnly)
        {
            try
            {
                exportProgress.Visibility = Visibility.Visible;
                statusText.Text = "Подготовка экспорта...";
                AddToMiniLog("Начат экспорт в Excel");

                var dataToExport = exportSelectedOnly && dgCerts.SelectedItem != null
                    ? new ObservableCollection<CertRecord> { (CertRecord)dgCerts.SelectedItem }
                    : _items;

                if (!dataToExport.Any())
                {
                    MessageBox.Show("Нет данных для экспорта", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    AddToMiniLog("Экспорт: нет данных");
                    return;
                }

                // Диалог выбора файла - теперь только Excel
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Сертификаты_ЭЦП_{DateTime.Now:yyyy-MM-dd_HH-mm}" +
                              (exportSelectedOnly ? "_выделенный" : "") + ".xlsx"
                };

                if (saveFileDialog.ShowDialog() != true)
                {
                    return; // Пользователь отменил
                }

                string filePath = saveFileDialog.FileName;

                // Генерация файла в фоновом потоке
                await Task.Run(() =>
                {
                    GenerateExcelFile(dataToExport, filePath);
                });

                // Сообщение об успехе
                var openResult = MessageBox.Show("Excel файл успешно экспортирован. Открыть файл?",
                                               "Экспорт завершен",
                                               MessageBoxButton.YesNo,
                                               MessageBoxImage.Question);

                if (openResult == MessageBoxResult.Yes)
                {
                    try
                    {
                        Process.Start(filePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть файл: {ex.Message}", "Ошибка",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }

                statusText.Text = exportSelectedOnly
                    ? "Выделенная запись экспортирована в Excel"
                    : $"Экспортировано записей в Excel: {dataToExport.Count}";

                AddToMiniLog($"Excel экспорт завершен: {dataToExport.Count} записей");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                Log($"Ошибка экспорта: {ex.Message}");
                AddToMiniLog($"Ошибка экспорта: {ex.Message}");
            }
            finally
            {
                exportProgress.Visibility = Visibility.Collapsed;
            }
        }

        private void GenerateCsvFile(ObservableCollection<CertRecord> data, string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                // Заголовки
                writer.WriteLine("ФИО;Номер сертификата;Дата начала;Дата окончания;Осталось дней;Здание;Примечание;Статус");

                // Данные
                foreach (var record in data)
                {
                    var status = record.IsDeleted ? "Удален" : "Активен";
                    var fio = EscapeCsvField(record.Fio ?? "");
                    var certNumber = EscapeCsvField(record.CertNumber ?? "");
                    var building = EscapeCsvField(record.Building ?? "");
                    var note = EscapeCsvField(record.Note ?? "");

                    writer.WriteLine($"{fio};{certNumber};{record.DateStart:dd.MM.yyyy};{record.DateEnd:dd.MM.yyyy};{record.DaysLeft};{building};{note};{status}");
                }
            }
        }

        private void GenerateExcelFile(ObservableCollection<CertRecord> data, string filePath)
        {
            try
            {
                // Создаем новую книгу Excel
                var workbook = new ClosedXML.Excel.XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Сертификаты ЭЦП");

                // Стили для оформления
                var headerStyle = workbook.Style;
                headerStyle.Font.Bold = true;
                headerStyle.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                headerStyle.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                headerStyle.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                var normalStyle = workbook.Style;
                normalStyle.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                // Заголовок отчета
                worksheet.Cell(1, 1).Value = "Отчет по сертификатам ЭЦП";
                worksheet.Range(1, 1, 1, 8).Merge();
                worksheet.Cell(1, 1).Style.Font.Bold = true;
                worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                worksheet.Cell(1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                // Подзаголовок с датой
                worksheet.Cell(2, 1).Value = $"Сформирован: {DateTime.Now:dd.MM.yyyy HH:mm}";
                worksheet.Range(2, 1, 2, 8).Merge();
                worksheet.Cell(2, 1).Style.Font.Italic = true;
                worksheet.Cell(2, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                // Заголовки столбцов
                string[] headers = {
                    "ФИО",
                    "Номер сертификата",
                    "Дата начала",
                    "Дата окончания",
                    "Осталось дней",
                    "Здание",
                    "Примечание",
                    "Статус"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(4, i + 1).Value = headers[i];
                    worksheet.Cell(4, i + 1).Style = headerStyle;
                }

                // Данные
                int row = 5;
                foreach (var record in data)
                {
                    worksheet.Cell(row, 1).Value = record.Fio ?? "";
                    worksheet.Cell(row, 2).Value = record.CertNumber ?? "";
                    worksheet.Cell(row, 3).Value = record.DateStart;
                    worksheet.Cell(row, 3).Style.NumberFormat.Format = "dd.mm.yyyy";
                    worksheet.Cell(row, 4).Value = record.DateEnd;
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "dd.mm.yyyy";
                    worksheet.Cell(row, 5).Value = record.DaysLeft;
                    worksheet.Cell(row, 6).Value = record.Building ?? "";
                    worksheet.Cell(row, 7).Value = record.Note ?? "";
                    worksheet.Cell(row, 8).Value = record.IsDeleted ? "Удален" : "Активен";

                    // Цветовое кодирование по сроку действия
                    if (!record.IsDeleted)
                    {
                        if (record.DaysLeft <= 10)
                        {
                            worksheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightCoral;
                        }
                        else if (record.DaysLeft <= 30)
                        {
                            worksheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightYellow;
                        }
                        else
                        {
                            worksheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGreen;
                        }
                    }
                    else
                    {
                        // Для удаленных записей - серый цвет
                        worksheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                    }

                    // Применяем стиль границ ко всем ячейкам строки
                    for (int col = 1; col <= headers.Length; col++)
                    {
                        worksheet.Cell(row, col).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    }

                    row++;
                }

                // Настраиваем ширину столбцов
                worksheet.Columns().AdjustToContents();

                // Добавляем автофильтр
                worksheet.Range(4, 1, row - 1, headers.Length).SetAutoFilter();

                // Замораживаем заголовки
                worksheet.SheetView.FreezeRows(4);

                // Сохраняем файл
                workbook.SaveAs(filePath);

                AddToMiniLog($"Создан Excel файл: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                AddToMiniLog($"Ошибка создания Excel: {ex.Message}");
                throw new Exception($"Ошибка при создании Excel файла: {ex.Message}", ex);
            }
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(";") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private void TryGenerateExcelWithClosedXML(ObservableCollection<CertRecord> data, string filePath)
        {
            try
            {
                GenerateExcelFile(data, filePath);
            }
            catch (Exception ex)
            {
                // Если не удалось создать XLSX, пробуем создать CSV как запасной вариант
                string csvPath = Path.ChangeExtension(filePath, ".csv");
                GenerateCsvFile(data, csvPath);
                throw new Exception($"Не удалось создать Excel файл: {ex.Message}. Создан CSV файл: {Path.GetFileName(csvPath)}");
            }
        }

        
        private async Task LoadFromServer()
        {
            // защита от повторных кликов
            if (_isLoadingFromServer)
                return;

            _isLoadingFromServer = true;

            try
            {
                if (_api == null)
                {
                    AddToMiniLog("API ещё не инициализирован");
                    return;
                }

                statusText.Text = "Загрузка данных с сервера...";
                SetServerState(ServerConnectionState.Busy);

                // ===== ЗАГРУЗКА С СЕРВЕРА (НЕ UI ПОТОК) =====

                AddToMiniLog("Запрос GET_CERTS отправлен");

                var list = await _api.GetCertificates();
                AddToMiniLog("CLIENT received records: " + list.Count);

                AddToMiniLog("Ответ получен");

                if (list == null)
                {
                    AddToMiniLog("ОШИБКА: list == null");
                    list = new List<CertRecord>();
                }

                // ===== ОБНОВЛЕНИЕ UI =====

                await Dispatcher.InvokeAsync(() =>
                {
                    if (_allItems == null)
                        _allItems = new ObservableCollection<CertRecord>();

                    _allItems.Clear();

                    foreach (var item in list)
                    {
                        if (item != null)
                            _allItems.Add(item);
                    }

                    ApplySearchFilter();

                    statusText.Text = $"Загружено записей: {_allItems.Count}";
                });

                AddToMiniLog($"Получено с сервера: {list.Count} записей");
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка загрузки: " + ex.Message);

                Dispatcher.Invoke(() =>
                {
                    statusText.Text = "Ошибка загрузки с сервера";
                });
            }
            finally
            {
                _isLoadingFromServer = false;

                if (_serverState != ServerConnectionState.Offline)
                    SetServerState(ServerConnectionState.Online);
            }
        }


        // ★ МЕТОД ДЛЯ РАСПАКОВКИ АРХИВА В ПАПКУ ★
        private bool ExtractArchiveToFolder(string archivePath, string extractToFolder)
        {
            try
            {
                // Используем System.IO.Compression для распаковки ZIP
                System.IO.Compression.ZipFile.ExtractToDirectory(archivePath, extractToFolder);
                return true;
            }
            catch (Exception ex)
            {
                Log($"Ошибка распаковки архива {archivePath}: {ex.Message}");
                return false;
            }
        }

        // ★ МЕТОД ДЛЯ СОЗДАНИЯ ВАЛИДНОГО ИМЕНИ ПАПКИ ★
        private string MakeValidFolderName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "Unknown";

            var invalidChars = Path.GetInvalidFileNameChars();
            var validName = new string(name.Where(ch => !invalidChars.Contains(ch)).ToArray());

            // Убираем пробелы в начале и конце, заменяем множественные пробелы на один
            validName = System.Text.RegularExpressions.Regex.Replace(validName.Trim(), @"\s+", " ");

            // Ограничиваем длину имени папки
            if (validName.Length > 50)
                validName = validName.Substring(0, 50);

            return validName;
        }

        private void ApplySearchFilter()
        {
            try
            {
                _items.Clear();

                IEnumerable<CertRecord> query = _allItems;

                // ===== ФИЛЬТР УДАЛЕННЫХ =====
                if (!_showDeleted)
                {
                    query = query.Where(x =>
                        !x.IsDeleted &&
                        x.DaysLeft >= 0
                    );
                }

                // ===== ФИЛЬТР ЗДАНИЯ =====
                if (_currentBuildingFilter != "Все")
                {
                    query = query.Where(x =>
                        string.Equals(x.Building, _currentBuildingFilter,
                        StringComparison.OrdinalIgnoreCase));
                }

                // ===== ПОИСК ПО ФИО =====
                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    var terms = _searchText.ToLower().Split(new[] { ' ' },
                        StringSplitOptions.RemoveEmptyEntries);

                    query = query.Where(item =>
                    {
                        var fio = (item.Fio ?? "").ToLower();
                        return terms.All(t => fio.Contains(t));
                    });
                }

                // ===== ЗАПОЛНЕНИЕ UI =====
                foreach (var item in query)
                    _items.Add(item);

                searchStatusText.Text =
                    $"Показано: {_items.Count} из {_allItems.Count}";

                Dispatcher.BeginInvoke(new Action(AutoFitDataGridColumns),
                    DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                Log("Ошибка фильтра: " + ex.Message);
            }
        }

        private void TxtSearchFio_TextChanged(object sender, TextChangedEventArgs e)
        {
            _searchText = txtSearchFio.Text.Trim();
            ApplySearchFilter();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearchFio.Text = "";
            _searchText = "";
            ApplySearchFilter();
            AddToMiniLog("Поиск очищен");
        }

        private void ShowProgressBar(bool show)
        {
            Dispatcher.Invoke(() =>
            {
                var visibility = show ? Visibility.Visible : Visibility.Collapsed;

                // Основная вкладка
                progressMailCheck.Visibility = visibility;

                // Вкладка настроек
                progressMailCheckSettings.Visibility = visibility;

                // Вкладка логов
                progressMailCheckLogs.Visibility = visibility;
            });
        }
        
        

        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FAST_CHECK");
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FULL_CHECK");
        }

        private async Task SendServerCommand(string cmd)
        {
            try
            {
                statusText.Text = "Связь с сервером...";
                Cursor = Cursors.Wait;

                var response = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    cmd);

                statusText.Text = "Ответ сервера: " + response;
                AddToMiniLog("SERVER: " + response);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка подключения к серверу:\n{ex.Message}",
                    "TCP ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                statusText.Text = "Сервер недоступен";
                AddToMiniLog("SERVER ERROR: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        private async void BuildingFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAllBuildings.IsChecked == true)
                _currentBuildingFilter = "Все";
            else if (rbKrasnoflotskaya.IsChecked == true)
                _currentBuildingFilter = "Краснофлотская";
            else if (rbPionerskaya.IsChecked == true)
                _currentBuildingFilter = "Пионерская";

            await LoadFromServer();
            AddToMiniLog($"Фильтр: {_currentBuildingFilter}");
        }

        private async void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            await LoadFromServer();
            AddToMiniLog(_showDeleted ? "Показаны удаленные" : "Скрыты удаленные");
        }


        private async void BuildingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSavingBuilding)
                return;

            if (!(sender is ComboBox comboBox))
                return;

            if (!(comboBox.DataContext is CertRecord record))
                return;

            // ===== ЗАЩИТА ОТ ИНИЦИАЛИЗАЦИИ =====

            if (e.AddedItems.Count == 0)
                return;

            string newValue = e.AddedItems[0] as string;

            // игнорируем пустые значения
            if (string.IsNullOrWhiteSpace(newValue))
                return;

            string oldValue = e.RemovedItems.Count > 0
                ? e.RemovedItems[0] as string
                : record.Building;

            // если реально не изменилось
            if (string.Equals(newValue, oldValue, StringComparison.Ordinal))
                return;

            _isSavingBuilding = true;

            try
            {
                await _api.UpdateBuilding(record.Id, newValue);

                AddToMiniLog($"Здание изменено: {record.Fio} → {newValue}");
                statusText.Text = "Здание сохранено";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения здания:\n" + ex.Message);

                // rollback UI если сервер не принял
                record.Building = oldValue;
            }
            finally
            {
                _isSavingBuilding = false;
            }
        }



        private void DgCerts_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            
        }

        private void DgCerts_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            // Логика без изменений
        }

        private void DgCerts_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit && e.Column.Header.ToString() == "Здание")
            {
                var comboBox = e.EditingElement as ComboBox;
                if (comboBox?.DataContext is CertRecord record)
                {
                    // Не сохраняем здесь, т.к. уже сохранили в SelectionChanged
                }
            }
        }

        private async void SaveBuilding(CertRecord record)
        {
            try
            {
                if (!string.IsNullOrEmpty(record.Building))
                {
                    await _api.UpdateBuilding(record.Id, record.Building);
                    await LoadFromServer();
                    statusText.Text = "Здание сохранено";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                try
                {
                    var result = MessageBox.Show(
                    $"Пометить сертификат:\n\n{record.Fio}\n\nкак удалённый?",
                    "Подтверждение",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                    if (result != MessageBoxResult.Yes)
                        return;
                    Log($"Пометка на удаление: {record.Fio} (ID: {record.Id})");
                    await _api.MarkDeleted(record.Id, true);

                    // локальное обновление UI
                    record.IsDeleted = true;

                    // обновляем представление
                    ApplySearchFilter();

                    statusText.Text = "Запись помечена на удаление";
                    AddToMiniLog($"Удален: {record.Fio}");
                    statusText.Text = "Запись помечена на удаление";
                    AddToMiniLog($"Удален: {record.Fio}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при пометке на удаление: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка удаления: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для пометки на удаление", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
                

        private void DgCerts_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is CertRecord record)
            {
                // Устанавливаем только фон, цвет текста будет управляться стилями
                if (record.IsDeleted)
                {
                    // Для удаленных записей - серый фон
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 80, 80, 80)  // Темно-серый для темной темы
                        : System.Windows.Media.Color.FromArgb(255, 220, 220, 220)); // Светло-серый для светлой темы
                }
                else if (record.DaysLeft <= 10)
                {
                    // Красный фон для срочных сертификатов
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 120, 50, 50)   // Темно-красный для темной темы
                        : System.Windows.Media.Color.FromArgb(255, 255, 200, 200)); // Светло-красный для светлой темы
                }
                else if (record.DaysLeft <= 30)
                {
                    // Желтый фон для предупреждения
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 120, 120, 50)   // Темно-желтый для темной темы
                        : System.Windows.Media.Color.FromArgb(255, 255, 255, 200)); // Светло-желтый для светлой темы
                }
                else
                {
                    // Зеленый фон для нормальных сертификатов
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 50, 80, 50)     // Темно-зеленый для темной темы
                        : System.Windows.Media.Color.FromArgb(255, 200, 255, 200)); // Светло-зеленый для светлой темы
                }

                // Принудительно устанавливаем цвет текста для всей строки
                e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                    ? Colors.White : Colors.Black);
            }
        }

        private async void StartReconnectBackoff()
        {
            if (_reconnectInProgress)
                return;

            _reconnectInProgress = true;


            SetServerState(ServerConnectionState.Offline);


            await Task.Delay(_reconnectDelaySeconds * 1000);

            _reconnectDelaySeconds =
                Math.Min(_reconnectDelaySeconds * 2, MAX_RECONNECT_DELAY);

            _reconnectInProgress = false;
            SetServerState(ServerConnectionState.Connecting);
        }
        private void DgCerts_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                var menuItems = dgCerts.ContextMenu.Items;
                var archiveMenuItem = (MenuItem)menuItems[0];
                var exportMenuItem = (MenuItem)menuItems[2];

                if (!string.IsNullOrEmpty(record.ArchivePath) && File.Exists(record.ArchivePath))
                {
                    archiveMenuItem.IsEnabled = false;
                    archiveMenuItem.ToolTip = "Архив уже прикреплен";
                }
                else
                {
                    archiveMenuItem.IsEnabled = true;
                    archiveMenuItem.ToolTip = "Добавить архив с подписью";
                }

                exportMenuItem.IsEnabled = dgCerts.SelectedItem != null;
            }
        }

        

        private string MakeValidFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(name.Where(ch => !invalidChars.Contains(ch)).ToArray());
        }

        private void LoadLogs()
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtLogs.Text = "Папка логов не существует";
                        logStatusText.Text = "Логи не найдены";
                    });
                    return;
                }

                // Ищем папку с сегодняшней датой
                var todayFolder = Path.Combine(logBaseDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(todayFolder))
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtLogs.Text = "Логи за сегодня не найдены";
                        logStatusText.Text = "Нет логов за сегодня";
                    });
                    return;
                }

                // Ищем самый свежий файл сессии за сегодня
                var sessionLogs = Directory.GetFiles(todayFolder, "session_*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .ToList();

                if (!sessionLogs.Any())
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtLogs.Text = "Файлы сессий не найдены";
                        logStatusText.Text = "Нет логов сессий";
                    });
                    return;
                }

                // Берем самый свежий файл сессии (текущей сессии)
                var latestSessionLog = sessionLogs.First();
                var logContent = File.ReadAllText(latestSessionLog.FullName);

                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = logContent;
                    logStatusText.Text = $"Логи текущей сессии ({latestSessionLog.CreationTime:dd.MM.yyyy HH:mm})";
                    txtLogs.ScrollToEnd();
                });

                AddToMiniLog($"Загружены логи текущей сессии");
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = $"Ошибка загрузки логов: {ex.Message}";
                    logStatusText.Text = "Ошибка загрузки";
                });
                AddToMiniLog($"Ошибка загрузки логов: {ex.Message}");
            }
        }

        private void CleanOldLogs()
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                    return;

                var cutoffDate = DateTime.Now.AddMonths(-1);
                int deletedCount = 0;

                var dateFolders = Directory.GetDirectories(logBaseDirectory)
                    .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(dir), @"^\d{4}-\d{2}-\d{2}$"));

                foreach (var dateFolder in dateFolders)
                {
                    var folderName = Path.GetFileName(dateFolder);
                    if (DateTime.TryParseExact(folderName, "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out DateTime folderDate))
                    {
                        if (folderDate < cutoffDate.Date)
                        {
                            try
                            {
                                Directory.Delete(dateFolder, true);
                                deletedCount++;
                            }
                            catch
                            {
                                // Игнорируем ошибки
                            }
                        }
                    }
                }
            }
            catch
            {
                // Игнорируем все ошибки при очистке
            }
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is TabControl tabControl)
            {
                // Проверяем, что переключились именно на вкладку "Логи"
                if (tabControl.SelectedItem is TabItem selectedTab && selectedTab.Header?.ToString() == "Логи")
                {
                    // Автоматически обновляем логи при переходе на вкладку
                    LoadLogs();
                    AddToMiniLog("Автообновление логов при переходе на вкладку");
                }

                // Можно добавить обработку других вкладок при необходимости
                else if (tabControl.SelectedItem is TabItem settingsTab && settingsTab.Header?.ToString() == "Настройки")
                {
                    // Например, обновить список папок при переходе на настройки
                    // Или выполнить другие действия
                }
            }
        }

        private void LogsTab_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Проверяем, что переключились именно на вкладку "Логи"
            if (e.Source is TabControl tabControl &&
                tabControl.SelectedItem is TabItem selectedTab &&
                selectedTab.Header?.ToString() == "Логи")
            {
                // Автоматически обновляем логи при переходе на вкладку
                LoadLogs();
                AddToMiniLog("Автообновление логов при переходе на вкладку");
            }
        }

        private void BtnRefreshLogs_Click(object sender, RoutedEventArgs e)
        {
            LoadLogs();
            AddToMiniLog("Логи обновлены вручную");
        }

        private void BtnClearLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logBaseDirectory))
                {
                    MessageBox.Show("Папка логов не существует", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var cutoffDate = DateTime.Now.AddMonths(-2);
                int deletedCount = 0;

                // Ищем все папки с датами
                var dateFolders = Directory.GetDirectories(logBaseDirectory)
                    .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(dir), @"^\d{4}-\d{2}-\d{2}$"));

                foreach (var dateFolder in dateFolders)
                {
                    var folderName = Path.GetFileName(dateFolder);
                    if (DateTime.TryParseExact(folderName, "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out DateTime folderDate))
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

                MessageBox.Show($"Удалено папок с логами: {deletedCount}", "Очистка логов",
                              MessageBoxButton.OK, MessageBoxImage.Information);

                LoadLogs();
                AddToMiniLog($"Очищено логов: {deletedCount} папок");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке логов: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
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

                // Открываем основную папку LOG (пользователь сам увидит папки с датами)
                System.Diagnostics.Process.Start("explorer.exe", logBaseDirectory);
                AddToMiniLog("Открыта папка логов");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии папки логов: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                AddToMiniLog($"Ошибка открытия папки логов: {ex.Message}");
            }
        }

        

        private void Log(string message)
        {
            try
            {
                string logFile = ImapCertWatcher.Utils.LogSession.SessionLogFile;

                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [MainWindow] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

                AddToMiniLog($"[MainWindow] {message}");
                System.Diagnostics.Debug.WriteLine($"[MainWindow] {message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи в лог: {ex.Message}");
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
        }


        protected override void OnClosed(EventArgs e)
        {
            
            _serverMonitorTimer?.Stop();
            base.OnClosed(e);

            
            _refreshTimer?.Stop();
            if (_refreshTimer != null)
                _refreshTimer.Tick -= RefreshDaysLeft;

            
        }

        // ================= TEMP STUBS =================

        private async void BtnRestoreSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                try
                {
                    await _api.MarkDeleted(record.Id, false);

                    // локальное обновление UI
                    record.IsDeleted = false;

                    // обновляем таблицу
                    ApplySearchFilter();

                    statusText.Text = "Запись восстановлена";
                    AddToMiniLog("Восстановлено: " + record.Fio);

                    statusText.Text = "Запись восстановлена";
                    AddToMiniLog("Восстановлено: " + record.Fio);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка восстановления:\n" + ex.Message);
                }
            }
        }


        private void BtnLoadCer_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Ручная загрузка сертификата временно отключена.",
                "Временно недоступно",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private async void AddArchiveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "ZIP (*.zip)|*.zip";

                if (dlg.ShowDialog() != true)
                    return;

                byte[] bytes = File.ReadAllBytes(dlg.FileName);
                string base64 = Convert.ToBase64String(bytes);
                string fileName = Path.GetFileName(dlg.FileName);

                await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    $"ADD_ARCHIVE|{record.Id}|{fileName}|{base64}");

                AddToMiniLog("Архив добавлен");
            }
        }

        private async void BtnOpenArchive_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var record = dgCerts.SelectedItem as CertRecord;
                if (record == null)
                    return;

                int certId = record.Id;
                string fio = record.Fio;

                // TEMP\ImapCertWatcher\ФИО
                string safeName = MakeSafeFolderName(fio);

                string certsRoot = GetCertsRoot();

                string certFolder = Path.Combine(
                    certsRoot,
                    safeName);

                // =====================================================
                // ✅ ПРОВЕРКА ЛОКАЛЬНОГО КЭША (ПЕРЕД TCP)
                // =====================================================

                if (Directory.Exists(certFolder))
                {
                    var files = Directory.GetFiles(certFolder);

                    if (files.Length > 0)
                    {
                        // Уже загружено ранее — просто открываем папку
                        Process.Start("explorer.exe", certFolder);
                        return;
                    }
                }

                // =====================================================
                // 🔽 ЕСЛИ НЕТ КЭША — ИДЁМ НА СЕРВЕР
                // =====================================================

                var response = await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    $"GET_ARCHIVE|{certId}");

                if (string.IsNullOrWhiteSpace(response) || !response.StartsWith("ARCHIVE|"))
                {
                    MessageBox.Show("Архив не найден");
                    return;
                }

                // ARCHIVE filename|base64
                var payload = response.Substring("ARCHIVE|".Length);
                var parts = payload.Split(new[] { '|' }, 2);

                string fileName = parts[0];
                byte[] fileData = Convert.FromBase64String(parts[1]);

                // создаем папку
                Directory.CreateDirectory(certFolder);

                // сохраняем файл
                string filePath = Path.Combine(certFolder, fileName);

                File.WriteAllBytes(filePath, fileData);

                // ===== РЕШЕНИЕ ПО ТИПУ ФАЙЛА =====

                if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    ZipFile.ExtractToDirectory(filePath, certFolder);

                    File.Delete(filePath);

                    Process.Start("explorer.exe", certFolder);
                }
                else
                {
                    // CER — просто открываем папку
                    Process.Start("explorer.exe", certFolder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия сертификата:\n" + ex.Message);
            }
        }

        private string GetCertsRoot()
        {
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Certs");
        }

        private async Task CleanupCertsFolderAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    string root = GetCertsRoot();

                    Directory.CreateDirectory(root);

                    foreach (var dir in Directory.GetDirectories(root))
                    {
                        try
                        {
                            var info = new DirectoryInfo(dir);
                            var age = DateTime.Now - info.LastWriteTime;

                            if (age > TimeSpan.FromHours(48))
                            {
                                Directory.Delete(dir, true);
                            }
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                }
            });
        }

        private string MakeSafeFolderName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            return name.Trim();
        }

        private async void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is CertRecord record)
            {
                await TcpCommandClient.SendAsync(
                    _clientSettings.ServerIp,
                    _clientSettings.ServerPort,
                    $"UPDATE_NOTE|{record.Id}|{tb.Text}");

                AddToMiniLog("Примечание сохранено");
            }
        }

    }
}