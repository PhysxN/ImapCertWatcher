//using DocumentFormat.OpenXml.Office2010.Excel;
using ImapCertWatcher;
using ImapCertWatcher.Client;
using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.UI;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
// Альтернатива Excel без использования Office Interop
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window
    {
        private ClientSettings _clientSettings;  
        private HashSet<int> _knownCertIds = new HashSet<int>();
        private ServerApiClient _api;
        private ObservableCollection<CertRecord> _items = new ObservableCollection<CertRecord>();
        private ObservableCollection<CertRecord> _allItems = new ObservableCollection<CertRecord>();
        private ObservableCollection<TokenRecord> _tokens = new ObservableCollection<TokenRecord>();
        private ObservableCollection<TokenRecord> _freeTokens = new ObservableCollection<TokenRecord>();
        public ObservableCollection<TokenRecord> Tokens => _tokens;
        private ObservableCollection<TokenListItem> _filteredTokens
    = new ObservableCollection<TokenListItem>();
        public ObservableCollection<TokenListItem> FilteredTokens => _filteredTokens;

        public ObservableCollection<TokenRecord> FreeTokens => _freeTokens;
        private DispatcherTimer _refreshTimer;
        private ObservableCollection<string> _miniLogMessages = new ObservableCollection<string>();
        private const int MAX_MINI_LOG_LINES = 3; 
        private string _currentBuildingFilter = "Все";
        private bool _showDeleted = false;
        private string _searchText = "";
        private bool _isDarkTheme = false;        
        private bool _certsLoadedOnce = false;
        private bool _showBusyTokens = false;
        private bool _isAssigningToken = false;


        private DispatcherTimer _serverMonitorTimer;
        private DispatcherTimer _connectingAnimationTimer;
        private int _connectingDots = 0;
        private bool _reconnectInProgress = false;
        private bool _isLoadingFromServer = false;
        private bool _isApplyingFromServer = false;
        private int _monitorBusyFlag = 0;
        private int _reconnectBusy = 0;

        enum ServerConnectionState
        {
            Connecting,
            Online,
            Busy,
            Offline
        }
        private int _reconnectDelaySeconds = 2;
        private const int MAX_RECONNECT_DELAY = 30;
        private ServerConnectionState _serverState = ServerConnectionState.Connecting;
        private static readonly TimeSpan PollBusy = TimeSpan.FromMilliseconds(500);
        private static readonly TimeSpan PollOnline = TimeSpan.FromSeconds(3);
        private static readonly TimeSpan PollConnecting = TimeSpan.FromSeconds(1);



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
                SetServerState(ServerConnectionState.Offline);

                // Инициализация мини-лога
                _miniLogMessages = new ObservableCollection<string>();
                AddToMiniLog("Приложение запущено");

                // Очистка старых логов при запуске (в фоне)
                Task.Run(() => CleanOldLogs());

                // Загружаем настройки
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;

                var clientPath = Path.Combine(baseDir, "client.settings.txt");

                _clientSettings = SettingsLoader.LoadClient(clientPath);


                this.DataContext = _clientSettings;
                

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

        private void SetServerState(ServerConnectionState state)
        {
            if (_serverState == state && state != ServerConnectionState.Connecting)
                return;

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.InvokeAsync(() => SetServerState(state));
                return;
            }

            _serverState = state;

            if (state != ServerConnectionState.Busy)
                pbServerProgress.Visibility = Visibility.Collapsed;

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

                    txtNextCheck.Visibility = Visibility.Visible;
                    break;

                case ServerConnectionState.Busy:
                    txtServerStatus.Text = "🟠 Сервер работает";
                    txtServerStatus.Foreground = Brushes.Orange;

                    txtNextCheck.Visibility = Visibility.Collapsed;
                    break;

                case ServerConnectionState.Offline:
                    txtServerStatus.Text = "🔴 Сервер OFFLINE";
                    txtServerStatus.Foreground = Brushes.Red;

                    txtNextCheck.Text = "Следующая проверка: --:--";
                    txtNextCheck.Visibility = Visibility.Visible;
                    break;
            }
            AdjustPollingInterval();
        }

        
        private async void ServerMonitorTimer_Tick(object sender, EventArgs e)
        {
            await UpdateServerMonitor();
        }

        private void StartServerMonitor()
        {
            if (_serverMonitorTimer != null)
                return;

            _serverMonitorTimer = new DispatcherTimer();
            _serverMonitorTimer.Interval = PollConnecting;
            _serverMonitorTimer.Tick += ServerMonitorTimer_Tick;

            _serverMonitorTimer.Start();
            _ = UpdateServerMonitor();
        }

        private void AdjustPollingInterval()
        {
            if (_serverMonitorTimer == null)
                return;

            switch (_serverState)
            {
                case ServerConnectionState.Busy:
                    _serverMonitorTimer.Interval = PollBusy;
                    break;

                case ServerConnectionState.Online:
                    _serverMonitorTimer.Interval = PollOnline;
                    break;

                case ServerConnectionState.Connecting:
                    _serverMonitorTimer.Interval = PollConnecting;
                    break;

                case ServerConnectionState.Offline:
                    _serverMonitorTimer.Interval =
                        TimeSpan.FromSeconds(_reconnectDelaySeconds);
                    break;
            }
        }

        private async Task UpdateServerMonitor()
        {
            // ✅ Защита от null API
            if (_api == null)
                return;

            if (Interlocked.Exchange(ref _monitorBusyFlag, 1) == 1)
                return;

            try
            {
                var response = (await _api.SendCommand("STATUS"))?.Trim();
                

                // 🔴 Любая некорректность ответа → уходим в backoff
                if (string.IsNullOrWhiteSpace(response) ||
                    !response.StartsWith("STATE|"))
                {
                    await StartReconnectBackoff();
                    return;
                }

                // ✅ Ответ корректный — сервер жив
                ParseServerState(response);
            }
            catch
            {
                // 🔴 TCP ошибка
                await StartReconnectBackoff();
            }
            finally
            {
                Interlocked.Exchange(ref _monitorBusyFlag, 0);
            }
        }

        private void ParseServerState(string stateLine)
        {
            try
            {
                var parts = stateLine.Split('|');

                bool isBusy = false;
                int progress = 0;
                int timerMinutes = 0;
                string stage = "";

                foreach (var part in parts)
                {
                    if (part.StartsWith("BUSY="))
                        isBusy = part.Substring(5) == "1";

                    else if (part.StartsWith("PROGRESS="))
                        int.TryParse(part.Substring(9), out progress);

                    else if (part.StartsWith("TIMER="))
                        int.TryParse(part.Substring(6), out timerMinutes);

                    else if (part.StartsWith("STAGE="))
                        stage = part.Substring(6);
                }

                _reconnectDelaySeconds = 2;

                SetServerState(isBusy
                    ? ServerConnectionState.Busy
                    : ServerConnectionState.Online);

                if (isBusy)
                {
                    pbServerProgress.Visibility = Visibility.Visible;
                    pbServerProgress.Value = progress;
                }
                else
                {
                    pbServerProgress.Visibility = Visibility.Collapsed;
                }

                if (!isBusy)
                {
                    txtNextCheck.Text = timerMinutes <= 0
                        ? "Следующая проверка: сейчас"
                        : $"Следующая проверка через {timerMinutes} мин";
                }
            }
            catch
            {
                SetServerState(ServerConnectionState.Offline);
            }
        }

        private void StartConnectingAnimation()
        {
            if (_connectingAnimationTimer == null)
            {
                _connectingAnimationTimer = new DispatcherTimer();
                _connectingAnimationTimer.Interval = TimeSpan.FromMilliseconds(400);
                _connectingAnimationTimer.Tick += ConnectingAnimationTick;
            }

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

        private async Task ReportProgressStep(string text, double percent, int minDelayMs = 250)
        {
            OnProgressUpdated(text, percent);

            // даём UI отрисоваться
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);

            // минимальная пауза, чтобы пользователь увидел этап
            await Task.Delay(minDelayMs);
        }
        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _ = CleanupCertsFolderAsync();

            this.Loaded -= MainWindow_Loaded;

            System.Diagnostics.Debug.WriteLine("MainWindow_Loaded начат");

            try
            {
                await ReportProgressStep("Подготовка среды...", 5);

                await ReportProgressStep("Инициализация API...", 10);
                _api = new ServerApiClient(_clientSettings);

                await ReportProgressStep("Подключение к серверу...", 20);

                // ✅ Загружаем данные сразу при старте
                await LoadFromServer();
                _certsLoadedOnce = true; // Помечаем, что сертификаты уже загружены

                System.Diagnostics.Debug.WriteLine("Данные получены");

                await ReportProgressStep("Анализ полученных данных...", 35);
                _knownCertIds = new HashSet<int>(_allItems.Select(r => r.Id));

                await ReportProgressStep("Формирование списков токенов...", 50);
                RebuildAvailableTokens();

                await ReportProgressStep("Применение фильтров...", 60);
                ApplySearchFilter();

                await ReportProgressStep("Настройка таблицы...", 70);
                AutoFitDataGridColumns();

                await ReportProgressStep("Запуск таймеров...", 80);
                InitializeRefreshTimer();

                await ReportProgressStep("Загрузка логов...", 90);
                LoadLogs();

                await ReportProgressStep("Запуск мониторинга сервера...", 95);
                StartServerMonitor();

                await ReportProgressStep("Готово", 100, 400);

                DataLoaded?.Invoke();

                statusText.Text = _serverState == ServerConnectionState.Online
                    ? "Готово"
                    : "Ожидание сервера...";

                AddToMiniLog($"Инициализация завершена. Загружено записей: {_allItems.Count}");

                System.Diagnostics.Debug.WriteLine("MainWindow_Loaded завершен");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в MainWindow_Loaded: {ex.Message}");

                OnProgressUpdated($"Ошибка: {ex.Message}", -1);

                statusText.Text = "Ошибка инициализации";
                AddToMiniLog($"Ошибка инициализации: {ex.Message}");

                MessageBox.Show($"Ошибка инициализации: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

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
                foreach (var item in _items)
                {
                    item?.RefreshDaysLeft();
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
            void action()
            {
                _miniLogMessages.Insert(0, $"{DateTime.Now:HH:mm:ss} - {message}");

                while (_miniLogMessages.Count > MAX_MINI_LOG_LINES)
                    _miniLogMessages.RemoveAt(_miniLogMessages.Count - 1);

                UpdateMiniLogDisplay();
            }

            if (Dispatcher.CheckAccess())
                action();
            else
                Dispatcher.InvokeAsync(action);
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
                    chkDarkTheme.Checked -= ChkDarkTheme_Changed;
                    chkDarkTheme.Unchecked -= ChkDarkTheme_Changed;

                    chkDarkTheme.IsChecked = _isDarkTheme;
                    ApplyTheme(_isDarkTheme);

                    chkDarkTheme.Checked += ChkDarkTheme_Changed;
                    chkDarkTheme.Unchecked += ChkDarkTheme_Changed;
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
            var dictionaries = Application.Current.Resources.MergedDictionaries;

            // Удаляем старую тему (она всегда первая)
            if (dictionaries.Count > 0)
                dictionaries.RemoveAt(0);

            var themeDict = new ResourceDictionary
            {
                Source = new Uri(
                    dark
                        ? "Themes/DarkTheme.xaml"
                        : "Themes/LightTheme.xaml",
                    UriKind.Relative)
            };

            dictionaries.Insert(0, themeDict);
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
                    TryGenerateExcelWithClosedXML(dataToExport, filePath);
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
                        Process.Start(new ProcessStartInfo(filePath)
                        {
                            UseShellExecute = true
                        });
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
                    var status = record.IsRevoked ? "Аннулирован" : record.IsDeleted ? "Удален" : "Активен";
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
                worksheet.Range(1, 1, 1, 9).Merge();
                worksheet.Cell(1, 1).Style.Font.Bold = true;
                worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                worksheet.Cell(1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                // Подзаголовок с датой
                worksheet.Cell(2, 1).Value = $"Сформирован: {DateTime.Now:dd.MM.yyyy HH:mm}";
                worksheet.Range(2, 1, 2, 9).Merge();
                worksheet.Cell(2, 1).Style.Font.Italic = true;
                worksheet.Cell(2, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                // Заголовки столбцов
                string[] headers = {
                    "ФИО",
                    "Номер сертификата",
                    "Токен",
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
                    worksheet.Cell(row, 3).Value = record.TokenSn ?? "";
                    worksheet.Cell(row, 4).Value = record.DateStart;
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 5).Value = record.DateEnd;
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 6).Value = record.DaysLeft;
                    worksheet.Cell(row, 7).Value = record.Building ?? "";
                    worksheet.Cell(row, 8).Value = record.Note ?? "";
                    worksheet.Cell(row, 9).Value = record.IsRevoked
                        ? "Аннулирован"
                        : record.IsDeleted ? "Удален" : "Активен";

                    // Цветовое кодирование по сроку действия
                    if (!record.IsDeleted)
                    {
                        if (record.DaysLeft <= 10)
                        {
                            worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightCoral;
                        }
                        else if (record.DaysLeft <= 30)
                        {
                            worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightYellow;
                        }
                        else
                        {
                            worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGreen;
                        }
                    }
                    else
                    {
                        // Для удаленных записей - серый цвет
                        worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
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
            if (_isLoadingFromServer)
                return;

            _isLoadingFromServer = true;
            SetServerState(ServerConnectionState.Connecting);

            try
            {
                if (_api == null)
                {
                    AddToMiniLog("API ещё не инициализирован");
                    SetServerState(ServerConnectionState.Offline);
                    statusText.Text = "Ошибка: API не инициализирован";
                    return;
                }

                statusText.Text = "Загрузка данных с сервера...";
                AddToMiniLog("Запрос GET_CERTS отправлен");

                var certTask = _api.GetCertificates();
                var tokensTask = _api.GetTokens();
                var freeTask = _api.GetFreeTokens();

                await Task.WhenAll(certTask, tokensTask, freeTask);

                // ✅ После WhenAll используем .Result (задачи уже завершены)
                var list = certTask.Result ?? new List<CertRecord>();
                var tokens = tokensTask.Result ?? new List<TokenRecord>();
                var freeTokens = freeTask.Result ?? new List<TokenRecord>();

                await Dispatcher.InvokeAsync(() =>
                {
                    _isApplyingFromServer = true;

                    try
                    {
                        _allItems.Clear();
                        foreach (var item in list)
                            _allItems.Add(item);

                        _tokens.Clear();
                        foreach (var t in tokens)
                            _tokens.Add(t);

                        _freeTokens.Clear();
                        foreach (var t in freeTokens)
                            _freeTokens.Add(t);

                        RebuildAvailableTokens();
                        ApplySearchFilter();

                        statusText.Text = $"Загружено записей: {_allItems.Count}";

                        // ✅ Сервер считаем доступным только после успешного заполнения коллекций
                        SetServerState(ServerConnectionState.Online);
                    }
                    finally
                    {
                        _isApplyingFromServer = false;
                    }
                });

                AddToMiniLog($"Получено с сервера: {list.Count} записей");
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка загрузки: " + ex.Message);
                SetServerState(ServerConnectionState.Offline);
                statusText.Text = "Ошибка загрузки с сервера";
            }
            finally
            {
                _isLoadingFromServer = false;
            }
        }



        

        private void ApplySearchFilter()
        {
            try
            {
                if (_items == null || _allItems == null || searchStatusText == null)
                    return;

                _items.Clear();

                IEnumerable<CertRecord> query =
                    _allItems.Where(x => x != null);

                if (!_showDeleted)
                {
                    query = query.Where(x =>
                        !x.IsDeleted &&
                        x.DaysLeft >= 0);
                }

                if (_currentBuildingFilter != "Все")
                {
                    query = query.Where(x =>
                        string.Equals(x.Building, _currentBuildingFilter,
                        StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    var terms = _searchText.ToLower().Split(
                        new[] { ' ' },
                        StringSplitOptions.RemoveEmptyEntries);

                    query = query.Where(item =>
                    {
                        var fio = (item.Fio ?? "").ToLower();
                        return terms.All(t => fio.Contains(t));
                    });
                }

                foreach (var item in query)
                    _items.Add(item);

                searchStatusText.Text =
                    $"Показано: {_items.Count} из {_allItems.Count}";

                Dispatcher.BeginInvoke(
                    new Action(AutoFitDataGridColumns),
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


        private async void BtnAddToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }
            var sn = Prompt.ShowDialog(
                "Введите серийный номер токена",
                "Добавление токена");

            if (string.IsNullOrWhiteSpace(sn))
                return;

            sn = sn.Trim().ToUpper();

            var response = await _api.AddToken(sn);

            if (string.IsNullOrEmpty(response))
            {
                MessageBox.Show("Нет ответа от сервера");
                return;
            }

            if (response?.StartsWith("ERROR|TOKEN_ALREADY_EXISTS") == true)
            {
                MessageBox.Show(
                    "Токен с таким серийным номером уже существует.",
                    "Дубликат",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!response.StartsWith("OK"))
            {
                MessageBox.Show(
                    "Ошибка добавления токена.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            await ReloadTokensOnly();
        }

        private async void BtnUnassignToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }
            if (!(dgTokens.SelectedItem is TokenListItem item) ||
                item.IsHeader ||
                item.Token == null)
                return;

            var token = item.Token;

            if (token.IsFree)
            {
                MessageBox.Show("Токен уже свободен.");
                return;
            }

            var result = MessageBox.Show(
                $"Освободить токен {token.Sn}?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            var response = await _api.UnassignToken(token.Id);

            if (string.IsNullOrEmpty(response))
            {
                MessageBox.Show("Нет ответа от сервера");
                return;
            }

            if (response.StartsWith("OK"))
            {
                await ReloadTokensOnly();
                AddToMiniLog($"Токен {token.Sn} освобожден");
            }
            else
            {
                MessageBox.Show("Ошибка освобождения токена:\n" + response);
            }
        }

        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                var win = new Window
                {
                    Title = caption,
                    Width = 400,
                    Height = 150,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    ResizeMode = ResizeMode.NoResize,
                    Owner = Application.Current.MainWindow
                };

                var panel = new StackPanel { Margin = new Thickness(10) };

                var tb = new TextBox();
                panel.Children.Add(new TextBlock { Text = text });
                panel.Children.Add(tb);

                var btn = new Button
                {
                    Content = "OK",
                    IsDefault = true,
                    Margin = new Thickness(0, 10, 0, 0),
                    HorizontalAlignment = HorizontalAlignment.Right
                };

                btn.Click += (_, __) => win.DialogResult = true;
                panel.Children.Add(btn);

                win.Content = panel;
                win.Loaded += (_, __) => tb.Focus();

                return win.ShowDialog() == true ? tb.Text : null;
            }
        }

        private async void BtnDeleteToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }
            if (dgTokens.SelectedItem is TokenListItem item &&
                !item.IsHeader &&
                item.Token != null)
            {
                var token = item.Token;

                var result = MessageBox.Show(
                    $"Удалить токен {token.Sn}?",
                    "Подтверждение",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                    return;

                var response = await _api.DeleteToken(token.Id);

                if (string.IsNullOrEmpty(response))
                {
                    MessageBox.Show("Нет ответа от сервера");
                    return;
                }

                if (!response.StartsWith("OK"))
                {
                    MessageBox.Show($"Ошибка удаления токена:\n{response}");
                    return;
                }

                AddToMiniLog($"Токен {token.Sn} удалён");
                await ReloadTokensOnly();
            }
        }



        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FAST_CHECK");
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
            var r = MessageBox.Show(
                    "Полная проверка может занять длительное время.\n\nЗапустить?",
                    "Подтверждение",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            await SendServerCommand("FULL_CHECK");
        }

        private async Task SendServerCommand(string cmd)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }
            try
            {
                statusText.Text = "Связь с сервером...";
                Mouse.OverrideCursor = Cursors.Wait;

                var response = await _api.SendCommand(cmd);

                // ✅ Защита от null/пустого ответа
                if (string.IsNullOrEmpty(response))
                {
                    statusText.Text = "Нет ответа от сервера";
                    AddToMiniLog("SERVER: <пустой ответ>");
                    MessageBox.Show("Сервер вернул пустой ответ",
                        "Предупреждение",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                statusText.Text = "Ответ сервера: " + response;
                AddToMiniLog("SERVER: " + response);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения:\n{ex.Message}");
                statusText.Text = "Сервер недоступен";
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void BuildingFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAllBuildings.IsChecked == true)
                _currentBuildingFilter = "Все";
            else if (rbKrasnoflotskaya.IsChecked == true)
                _currentBuildingFilter = "Краснофлотская";
            else if (rbPionerskaya.IsChecked == true)
                _currentBuildingFilter = "Пионерская";

            ApplySearchFilter(); // 🔥 ТОЛЬКО ФИЛЬТР
            AddToMiniLog($"Фильтр: {_currentBuildingFilter}");
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            ApplySearchFilter(); // 🔥
            AddToMiniLog(_showDeleted ? "Показаны удаленные" : "Скрыты удаленные");
        }


        private async void BuildingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Защита от null для API
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_isApplyingFromServer)
                return;
            if (_isSavingBuilding)
                return;

            if (!(sender is ComboBox comboBox))
                return;

            if (!(comboBox.DataContext is CertRecord record))
                return;

            // ===== ЗАЩИТА ОТ ИНИЦИАЛИЗАЦИИ =====

            if (e.AddedItems.Count == 0)
                return;

            // Защита от null в AddedItems[0]
            var newValue = e.AddedItems[0] as string;
            if (newValue == null)
                return;

            // игнорируем пустые значения
            if (string.IsNullOrWhiteSpace(newValue))
                return;

            // Защита от null в RemovedItems
            string oldValue = e.RemovedItems.Count > 0
                ? e.RemovedItems[0] as string ?? record.Building
                : record.Building;

            // если реально не изменилось
            if (string.Equals(newValue, oldValue, StringComparison.Ordinal))
                return;

            _isSavingBuilding = true;

            try
            {
                var resp = await _api.UpdateBuilding(record.Id, newValue);

                if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                {
                    record.Building = oldValue;
                    CollectionViewSource.GetDefaultView(comboBox.ItemsSource)?.Refresh();
                    MessageBox.Show("Ошибка сохранения здания");
                    return;
                }

                AddToMiniLog($"Здание изменено: {record.Fio} → {newValue}");
                statusText.Text = "Здание сохранено";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения здания:\n" + ex.Message);

                // rollback UI если сервер не принял
                record.Building = oldValue;
                CollectionViewSource.GetDefaultView(comboBox.ItemsSource)?.Refresh();
            }
            finally
            {
                _isSavingBuilding = false;
            }
        }

        private void RebuildAvailableTokens()
        {
            foreach (var cert in _allItems)
            {
                if (cert.AvailableTokens == null)
                    cert.AvailableTokens = new ObservableCollection<TokenRecord>();

                var target = cert.AvailableTokens;

                target.Clear();

                foreach (var t in _tokens)
                {
                    if (t != null && (t.IsFree || t.OwnerCertId == cert.Id))
                        target.Add(t);
                }
            }
        }

        private async void TokenComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_api == null)
                return;

            if (_isApplyingFromServer)
                return;

            if (_isAssigningToken)
                return;

            if (e.AddedItems.Count == 0)
                return;

            if (!(sender is ComboBox cb)) return;

            // ✅ Защита от фантомных срабатываний - реагируем только когда список действительно открыт
            if (!cb.IsKeyboardFocusWithin && !cb.IsDropDownOpen)
                return;

            if (!(cb.DataContext is CertRecord cert)) return;
            if (!(cb.SelectedItem is TokenRecord token)) return;

            if (cert.TokenId == token.Id)
                return;

            int? oldTokenId = cert.TokenId;
            string oldTokenSn = cert.TokenSn;

            _isAssigningToken = true;

            try
            {
                var response = await _api.AssignToken(cert.Id, token.Id);

                if (string.IsNullOrWhiteSpace(response) || !response.StartsWith("OK"))
                {
                    MessageBox.Show("Ошибка назначения токена:\n" + (response ?? "Нет ответа от сервера"));

                    cert.TokenId = oldTokenId;
                    cert.TokenSn = oldTokenSn;

                    // ✅ Сначала фиксируем изменения в DataGrid
                    dgCerts.CommitEdit(DataGridEditingUnit.Cell, true);
                    dgCerts.CommitEdit(DataGridEditingUnit.Row, true);

                    // ✅ Потом обновляем представление
                    CollectionViewSource.GetDefaultView(cb.ItemsSource)?.Refresh();

                    RebuildAvailableTokens();
                    return;
                }

                cert.TokenId = token.Id;
                cert.TokenSn = token.Sn;

                // ✅ Сначала фиксируем изменения в DataGrid
                dgCerts.CommitEdit(DataGridEditingUnit.Cell, true);
                dgCerts.CommitEdit(DataGridEditingUnit.Row, true);

                // ✅ Потом обновляем представление
                CollectionViewSource.GetDefaultView(cb.ItemsSource)?.Refresh();

                AddToMiniLog($"Назначен токен {token.Sn} → {cert.Fio}");

                RebuildAvailableTokens();
            }
            catch (Exception ex)
            {
                cert.TokenId = oldTokenId;
                cert.TokenSn = oldTokenSn;

                // ✅ Сначала фиксируем изменения в DataGrid
                dgCerts.CommitEdit(DataGridEditingUnit.Cell, true);
                dgCerts.CommitEdit(DataGridEditingUnit.Row, true);

                // ✅ Потом обновляем представление
                CollectionViewSource.GetDefaultView(cb.ItemsSource)?.Refresh();

                RebuildAvailableTokens();

                MessageBox.Show("Ошибка назначения токена:\n" + ex.Message);
            }
            finally
            {
                _isAssigningToken = false;
            }
        }

        private async void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tc = e.Source as TabControl;
            if (tc == null)
                return;

            var tab = tc.SelectedItem as TabItem;
            if (tab == null)
                return;

            var header = tab.Header != null ? tab.Header.ToString() : null;

            switch (header)
            {
                case "Токены":                    
                    break;

                case "Сертификаты":
                    RebuildAvailableTokens();
                    break;
            }
        }

        private async Task ReloadTokensOnly()
        {
            try
            {
                if (_api == null)
                    return;

                var tokens = await _api.GetTokens();
                var freeTokens = await _api.GetFreeTokens();

                await Dispatcher.InvokeAsync(() =>
                {
                    _tokens.Clear();
                    // ✅ Защита от null для списка токенов
                    foreach (var t in tokens ?? Enumerable.Empty<TokenRecord>())
                        _tokens.Add(t);

                    _freeTokens.Clear();
                    // ✅ Защита от null для списка свободных токенов
                    foreach (var t in freeTokens ?? Enumerable.Empty<TokenRecord>())
                        _freeTokens.Add(t);

                    RebuildAvailableTokens();
                    ApplyTokenFilter();
                });
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка обновления токенов: " + ex.Message);
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

        
        private async void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                if (_api == null)
                {
                    MessageBox.Show("Сервер не подключен");
                    return;
                }
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
                    var resp = await _api.MarkDeleted(record.Id, true);

                    if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка удаления");
                        return;
                    }

                    record.IsDeleted = true;

                    // обновляем представление
                    ApplySearchFilter();

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

        private async void BtnResetRevokes_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }
            var r1 = MessageBox.Show(
                "Эта операция удалит ВСЮ информацию об аннулировании сертификатов.\n\nПродолжить?",
                "ВНИМАНИЕ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r1 != MessageBoxResult.Yes)
                return;

            var r2 = MessageBox.Show(
                "Вы ТОЧНО уверены?\n\nОперация необратима!",
                "ПОДТВЕРЖДЕНИЕ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Error);

            if (r2 != MessageBoxResult.Yes)
                return;

            var resp = await _api.ResetRevokes();

            if (string.IsNullOrEmpty(resp))
            {
                MessageBox.Show("Нет ответа от сервера");
                return;
            }

            AddToMiniLog("RESET_REVOKES: " + resp);
            await LoadFromServer();

            statusText.Text = "Аннулирования сброшены";
        }


        private void DgCerts_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is CertRecord record)
            {
                // ===== ПРИОРИТЕТ: АННУЛИРОВАН =====
                if (record.IsRevoked)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 90, 60, 60)
                        : System.Windows.Media.Color.FromArgb(255, 255, 180, 180));

                    e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                        ? Colors.White : Colors.Black);

                    return;
                }

                // ===== УДАЛЕН =====
                if (record.IsDeleted)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 80, 80, 80)
                        : System.Windows.Media.Color.FromArgb(255, 220, 220, 220));
                }
                else if (record.DaysLeft <= 10)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 120, 50, 50)
                        : System.Windows.Media.Color.FromArgb(255, 255, 200, 200));
                }
                else if (record.DaysLeft <= 30)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 120, 120, 50)
                        : System.Windows.Media.Color.FromArgb(255, 255, 255, 200));
                }
                else
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? System.Windows.Media.Color.FromArgb(255, 50, 80, 50)
                        : System.Windows.Media.Color.FromArgb(255, 200, 255, 200));
                }

                e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                    ? Colors.White : Colors.Black);
            }
        }


        private async Task StartReconnectBackoff()
        {
            if (Interlocked.Exchange(ref _reconnectBusy, 1) == 1)
                return;

            // Меняем на Offline только если еще не в этом состоянии
            if (_serverState != ServerConnectionState.Offline)
                SetServerState(ServerConnectionState.Offline);

            await Task.Delay(_reconnectDelaySeconds * 1000);

            if (_serverState == ServerConnectionState.Offline)
                _reconnectDelaySeconds =
                    Math.Min(_reconnectDelaySeconds * 2, MAX_RECONNECT_DELAY);

            Interlocked.Exchange(ref _reconnectBusy, 0);

            // Переходим в Connecting только если все еще в Offline 
            // (не были переведены в другое состояние за время ожидания)
            if (_serverState == ServerConnectionState.Offline && !_reconnectInProgress)
                SetServerState(ServerConnectionState.Connecting);
        }
        private void DgCerts_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                if (!string.IsNullOrEmpty(record.ArchivePath) &&
                    File.Exists(record.ArchivePath))
                {
                    miAddArchive.IsEnabled = false;
                    miAddArchive.ToolTip = "Архив уже прикреплен";
                }
                else
                {
                    miAddArchive.IsEnabled = true;
                    miAddArchive.ToolTip = "Добавить архив с подписью";
                }

                miExportSelected.IsEnabled = true;
            }
        }

        private void LoadLogs()
        {
            try
            {
                var logBaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");

                if (!Directory.Exists(logBaseDirectory))
                {
                    UpdateLogsUI("Папка логов не существует", "Логи не найдены");
                    return;
                }

                var todayFolder = Path.Combine(logBaseDirectory, DateTime.Now.ToString("yyyy-MM-dd"));

                if (!Directory.Exists(todayFolder))
                {
                    UpdateLogsUI("Логи за сегодня не найдены", "Нет логов за сегодня");
                    return;
                }

                var sessionLogs = Directory.GetFiles(todayFolder, "session_*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .ToList();

                if (!sessionLogs.Any())
                {
                    UpdateLogsUI("Файлы сессий не найдены", "Нет логов сессий");
                    return;
                }

                var latestSessionLog = sessionLogs.First();

                var lines = File.ReadAllLines(latestSessionLog.FullName);

                var lastLines = lines
                    .Skip(Math.Max(0, lines.Length - 1000));

                txtLogs.Text = string.Join("\n", lastLines);
                logStatusText.Text = $"Логи текущей сессии ({latestSessionLog.CreationTime:dd.MM.yyyy HH:mm})";
                txtLogs.ScrollToEnd();

                AddToMiniLog($"Загружены логи текущей сессии");
            }
            catch (Exception ex)
            {
                txtLogs.Text = $"Ошибка загрузки логов: {ex.Message}";
                logStatusText.Text = "Ошибка загрузки";
                AddToMiniLog($"Ошибка загрузки логов: {ex.Message}");
            }
        }

        // Вспомогательный метод для обновления UI
        private void UpdateLogsUI(string logText, string statusText)
        {
            txtLogs.Text = logText;
            logStatusText.Text = statusText;
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

        private async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(sender is TabControl tc))
                return;

            if (!(tc.SelectedItem is TabItem tab))
                return;

            var header = tab.Header?.ToString();

            if (header == "Сертификаты")
            {
                // ✅ Загружаем только если еще не загружали
                if (!_certsLoadedOnce)
                {
                    _certsLoadedOnce = true;
                    await LoadFromServer();
                }
            }
            else if (header == "Токены")
            {
                await ReloadTokensOnly(); // только токены
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
                Process.Start(new ProcessStartInfo("explorer.exe", logBaseDirectory)
                {
                    UseShellExecute = true
                });
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

            if (_serverMonitorTimer != null)
            {
                _serverMonitorTimer.Stop();
                _serverMonitorTimer.Tick -= ServerMonitorTimer_Tick;
            }
            base.OnClosed(e);

            
            _refreshTimer?.Stop();
            if (_refreshTimer != null)
                _refreshTimer.Tick -= RefreshDaysLeft;

            
        }

        // ================= TEMP STUBS =================

        private void ChkShowBusyTokens_Changed(object sender, RoutedEventArgs e)
        {
            _showBusyTokens = chkShowBusyTokens.IsChecked == true;
            ApplyTokenFilter();
        }

        private void ApplyTokenFilter()
        {
            if (_tokens == null)
                return;

            _filteredTokens.Clear();

            var free = _tokens
            .Where(t => t != null && t.IsFree)
                .OrderBy(t => t.Sn)
                .ToList();

            var busy = _tokens
            .Where(t => t != null && !t.IsFree)
                .OrderBy(t => t.OwnerFio ?? "")
                .ToList();

            if (!_showBusyTokens)
                busy.Clear();

            if (free.Any())
            {
                _filteredTokens.Add(new TokenListItem
                {
                    IsHeader = true,
                    HeaderText = "Свободные"
                });

                foreach (var t in free)
                    _filteredTokens.Add(new TokenListItem { Token = t });
            }

            if (busy.Any())
            {
                _filteredTokens.Add(new TokenListItem
                {
                    IsHeader = true,
                    HeaderText = "Занятые"
                });

                foreach (var t in busy)
                    _filteredTokens.Add(new TokenListItem { Token = t });
            }

            tokensStatusText.Text = $"Показано токенов: {free.Count + busy.Count}";
            Debug.WriteLine(_filteredTokens.Count);
        }

        private async void BtnRestoreSelected_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (dgCerts.SelectedItem is CertRecord record)
            {
                try
                {
                    var resp = await _api.MarkDeleted(record.Id, false);

                    if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка восстановления:\n" + (resp ?? "Нет ответа сервера"));
                        return;
                    }

                    // ✅ локально обновляем UI только после OK
                    record.IsDeleted = false;
                    record.IsRevoked = false;
                    record.RevokeDate = null;

                    ApplySearchFilter();

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

                try
                {
                    byte[] bytes = File.ReadAllBytes(dlg.FileName);
                    string base64 = Convert.ToBase64String(bytes);
                    string fileName = Path.GetFileName(dlg.FileName);

                    var response = await _api.AddArchive(record.Id, fileName, base64);

                    if (string.IsNullOrEmpty(response))
                    {
                        MessageBox.Show("Нет ответа от сервера");
                        return;
                    }

                    if (!response.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка добавления архива:\n" + response);
                        return;
                    }

                    AddToMiniLog($"Архив {fileName} добавлен к записи {record.Fio}");
                    statusText.Text = "Архив успешно добавлен";
                }
                catch (FileNotFoundException)
                {
                    MessageBox.Show("Файл не найден. Возможно, он был удален или перемещен.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Нет прав доступа к файлу.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка ввода-вывода при чтении файла:\n{ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Неожиданная ошибка при чтении файла:\n{ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var clientPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "client.settings.txt");

                SettingsSaver.SaveClient(clientPath, _clientSettings);

                AddToMiniLog("Клиентские настройки сохранены");

                MessageBox.Show("Настройки сохранены.",
                    "Сохранено",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                _api = new ServerApiClient(_clientSettings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения:\n" + ex.Message);
            }
        }

        private async void BtnOpenArchive_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (Mouse.OverrideCursor == Cursors.Wait)
                return;

            try
            {
                var record = dgCerts.SelectedItem as CertRecord;
                if (record == null)
                    return;

                int certId = record.Id;
                string fio = record.Fio;

                string safeName = MakeSafeFolderName(fio);
                string certFolder = Path.Combine(GetCertsRoot(), safeName);

                bool hasFiles = false;

                try
                {
                    hasFiles = Directory.Exists(certFolder) &&
                               Directory.EnumerateFiles(certFolder, "*", SearchOption.AllDirectories).Any();
                }
                catch { }

                if (hasFiles)
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", certFolder)
                    {
                        UseShellExecute = true
                    });
                    return;
                }

                Mouse.OverrideCursor = Cursors.Wait;
                statusText.Text = "Загрузка архива...";

                var response = await _api.SendCommand($"GET_ARCHIVE|{certId}");

                if (string.IsNullOrWhiteSpace(response) || !response.StartsWith("ARCHIVE|"))
                {
                    Mouse.OverrideCursor = null;
                    MessageBox.Show("Архив не найден");
                    return;
                }

                var payload = response.Substring("ARCHIVE|".Length);
                var parts = payload.Split(new[] { '|' }, 2);

                if (parts.Length < 2)
                {
                    MessageBox.Show("Некорректный формат архива");
                    return;
                }

                string fileName = parts[0];
                byte[] fileData;
                try
                {
                    fileData = await Task.Run(() => Convert.FromBase64String(parts[1]));
                }
                catch
                {
                    MessageBox.Show("Повреждённые данные архива");
                    return;
                }

                string tempZip = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".zip");

                try
                {
                    await Task.Run(() => File.WriteAllBytes(tempZip, fileData));

                    statusText.Text = "Распаковка архива...";

                    if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        string targetFolder = certFolder;
                        string finalFolder = targetFolder;

                        await Task.Run(() =>
                        {
                            if (Directory.Exists(finalFolder))
                            {
                                try
                                {
                                    Directory.Delete(finalFolder, true);
                                }
                                catch
                                {
                                    finalFolder = finalFolder + "_" + DateTime.Now.Ticks;
                                    Directory.CreateDirectory(finalFolder);
                                }
                            }

                            Directory.CreateDirectory(finalFolder);
                            ZipFile.ExtractToDirectory(tempZip, finalFolder);
                        });

                        certFolder = finalFolder;
                    }
                    else
                    {
                        // Для одиночных файлов тоже используем тот же подход
                        if (Directory.Exists(certFolder))
                        {
                            try
                            {
                                Directory.Delete(certFolder, true);
                            }
                            catch
                            {
                                certFolder = certFolder + "_" + DateTime.Now.Ticks;
                            }
                        }

                        Directory.CreateDirectory(certFolder);
                        string filePath = Path.Combine(certFolder, fileName);
                        await Task.Run(() => File.WriteAllBytes(filePath, fileData));
                    }

                    Process.Start(new ProcessStartInfo("explorer.exe", certFolder)
                    {
                        UseShellExecute = true
                    });

                    AddToMiniLog($"Архив для {fio} загружен и распакован");
                    statusText.Text = "Архив успешно загружен";
                }
                finally
                {
                    if (File.Exists(tempZip))
                    {
                        try { File.Delete(tempZip); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия сертификата:\n" + ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                statusText.Text = "Готово";
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
            if (_api == null)
                return;
            if (sender is TextBox tb && tb.DataContext is CertRecord record)
            {
                try
                {
                    string note = tb.Text ?? "";
                    string base64Note = Convert.ToBase64String(Encoding.UTF8.GetBytes(note));

                    var resp = await _api.UpdateNote(record.Id, base64Note);

                    if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка сохранения примечания");
                        return;
                    }

                    AddToMiniLog("Примечание сохранено");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка сохранения примечания:\n" + ex.Message);
                }
            }
        }

    }
}