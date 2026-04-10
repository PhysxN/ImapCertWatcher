using ImapCertWatcher.Client;
using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using ImapCertWatcher.ViewModels;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Controls;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ClientSettings _clientSettings;
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
        private bool _certsLoadedOnce = false;        
        private bool _isAssigningToken = false;
        private bool _isOfflineUiLocked;
        private bool _shouldOpenCertificatesTabWhenConnectionRestored = true;
        private TokenService _tokenService;
        private TokensViewModel _tokensVm;
        private bool _isUpdatingThemeToggleState = false;
        private int _themeToggleRequestVersion = 0;
        private bool _isThemeVisualStateApplying = false;


        private DispatcherTimer _serverMonitorTimer;
        private DispatcherTimer _connectingAnimationTimer;
        private int _connectingDots = 0;
        private bool _isLoadingFromServer = false;
        private bool _isApplyingFromServer = false;
        private int _monitorBusyFlag = 0;
        private int _reconnectBusy = 0;
        private bool _reloadAfterServerWork = false;
        private bool _serverWasBusy = false;


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

        public TokensViewModel TokensVm
        {
            get => _tokensVm;
            private set
            {
                _tokensVm = value;
                OnPropertyChanged(nameof(TokensVm));
            }
        }

        // Флаг для предотвращения множественного сохранения
        private bool _isSavingBuilding = false;

        // Событие для обновления прогресса splash screen
        public event Action<string, double> ProgressUpdated;

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
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

                UpdateDeleteRestoreButtonState();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации приложения: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }
        }


        public event Action DataLoaded;


        private void AddToMiniLog(string message)
        {
            void action()
            {
                _miniLogMessages.Add($"{DateTime.Now:HH:mm:ss} - {message}");

                while (_miniLogMessages.Count > MAX_MINI_LOG_LINES)
                    _miniLogMessages.RemoveAt(0);

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

                if (txtMiniLog != null)
                {
                    txtMiniLog.Text = logText;
                    txtMiniLog.CaretIndex = txtMiniLog.Text.Length;
                    txtMiniLog.ScrollToEnd();
                }

                if (txtMiniLogSettings != null)
                {
                    txtMiniLogSettings.Text = logText;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка обновления мини-лога: " + ex.Message);
            }
        }

        private void LoadThemeSettings()
        {
            try
            {
                _isDarkTheme = _clientSettings?.DarkTheme == true;
                ApplyTheme(_isDarkTheme);
                SyncThemeToggleSwitchState();
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
                if (_clientSettings == null)
                    return;

                _clientSettings.DarkTheme = _isDarkTheme;

                var clientPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "client.settings.txt");

                SettingsSaver.SaveClient(clientPath, _clientSettings);
            }
            catch (Exception ex)
            {
                Log($"Ошибка сохранения настроек темы: {ex.Message}");
            }
        }

        private void ApplyTheme(bool dark)
        {
            var dictionaries = Application.Current.Resources.MergedDictionaries;

            for (int i = dictionaries.Count - 1; i >= 0; i--)
            {
                var source = dictionaries[i].Source?.OriginalString ?? "";

                if (source.EndsWith("Themes/LightTheme.xaml", StringComparison.OrdinalIgnoreCase) ||
                    source.EndsWith("Themes/DarkTheme.xaml", StringComparison.OrdinalIgnoreCase))
                {
                    dictionaries.RemoveAt(i);
                }
            }

            var themeDict = new ResourceDictionary
            {
                Source = new Uri(
                    dark
                        ? "Themes/DarkTheme.xaml"
                        : "Themes/LightTheme.xaml",
                    UriKind.Relative)
            };

            dictionaries.Insert(0, themeDict);

            SyncThemeToggleSwitchState();
        }

        private void SyncThemeToggleSwitchState()
        {
            if (themeToggleSwitch == null)
                return;

            _isUpdatingThemeToggleState = true;
            try
            {
                themeToggleSwitch.IsChecked = _isDarkTheme;
                themeToggleSwitch.ToolTip = _isDarkTheme
                    ? "Переключить на светлую тему"
                    : "Переключить на тёмную тему";
            }
            finally
            {
                _isUpdatingThemeToggleState = false;
            }

            if (themeToggleSwitch.IsLoaded)
                ApplyThemeToggleVisualState(_isDarkTheme);
        }

        private void ThemeToggleSwitch_Loaded(object sender, RoutedEventArgs e)
        {
            ApplyThemeToggleVisualState(_isDarkTheme);
        }

        private void ApplyThemeToggleVisualState(bool isDark)
        {
            if (themeToggleSwitch == null)
                return;

            themeToggleSwitch.ApplyTemplate();

            var lightTrack = themeToggleSwitch.Template.FindName("LightTrack", themeToggleSwitch) as Border;
            var darkTrack = themeToggleSwitch.Template.FindName("DarkTrack", themeToggleSwitch) as Border;
            var switchThumb = themeToggleSwitch.Template.FindName("SwitchThumb", themeToggleSwitch) as Border;

            if (lightTrack == null || darkTrack == null || switchThumb == null)
                return;

            _isThemeVisualStateApplying = true;
            try
            {
                lightTrack.BeginAnimation(UIElement.OpacityProperty, null);
                darkTrack.BeginAnimation(UIElement.OpacityProperty, null);
                switchThumb.BeginAnimation(FrameworkElement.MarginProperty, null);

                lightTrack.Opacity = isDark ? 0 : 1;
                darkTrack.Opacity = isDark ? 1 : 0;
                switchThumb.Margin = isDark
                    ? new Thickness(38, 0, 0, 0)
                    : new Thickness(4, 0, 0, 0);
            }
            finally
            {
                _isThemeVisualStateApplying = false;
            }
        }

        private async Task AnimateWindowOpacityAsync(double to, int durationMs)
        {
            var tcs = new TaskCompletionSource<bool>();

            var animation = new DoubleAnimation
            {
                To = to,
                Duration = TimeSpan.FromMilliseconds(durationMs),
                FillBehavior = FillBehavior.HoldEnd
            };

            animation.Completed += (s, e) => tcs.TrySetResult(true);

            BeginAnimation(Window.OpacityProperty, animation);

            await tcs.Task;
        }

        private async void ThemeToggleSwitch_Checked(object sender, RoutedEventArgs e)
        {
            await ApplyThemeFromSwitchAsync(true);
        }

        private async void ThemeToggleSwitch_Unchecked(object sender, RoutedEventArgs e)
        {
            await ApplyThemeFromSwitchAsync(false);
        }

        private async Task ApplyThemeFromSwitchAsync(bool dark)
        {
            if (_isUpdatingThemeToggleState || _isThemeVisualStateApplying)
                return;

            try
            {
                if (_isDarkTheme == dark)
                    return;

                int requestVersion = ++_themeToggleRequestVersion;

                await Task.Delay(25);

                if (requestVersion != _themeToggleRequestVersion)
                    return;

                await AnimateWindowOpacityAsync(0.94, 35);

                _isDarkTheme = dark;

                if (_clientSettings != null)
                    _clientSettings.DarkTheme = _isDarkTheme;

                ApplyTheme(_isDarkTheme);
                SyncThemeToggleSwitchState();

                if (dgCerts != null)
                    dgCerts.Items.Refresh();

                SaveThemeSettings();

                await AnimateWindowOpacityAsync(1.0, 45);

                Log($"[UI] Тема изменена: {(_isDarkTheme ? "Тёмная" : "Светлая")}");
            }
            catch (Exception ex)
            {
                Opacity = 1.0;

                MessageBox.Show(
                    "Ошибка переключения темы:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                SyncThemeToggleSwitchState();
            }
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

        private async Task RecreateApiAndTokenServicesAsync()
        {
            if (_tokenService != null)
                _tokenService.TokensChanged -= RebuildAvailableTokens;

            _api = new ServerApiClient(_clientSettings);

            _tokenService = new TokenService(_api);
            TokensVm = new TokensViewModel(_tokenService);

            _tokenService.TokensChanged += RebuildAvailableTokens;

            try
            {
                await _tokenService.Reload();
            }
            catch (Exception ex)
            {
                AddToMiniLog("Не удалось обновить токены после смены настроек: " + ex.Message);
            }

            RebuildAvailableTokens();
        }

        private void SelectCertificatesTab()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(SelectCertificatesTab);
                return;
            }

            if (tabCertificates == null || mainTabControl == null)
                return;

            if (tabCertificates.Visibility != Visibility.Visible || !tabCertificates.IsEnabled)
                return;

            _isSwitchingTabInternally = true;
            try
            {
                mainTabControl.SelectedItem = tabCertificates;
            }
            finally
            {
                _isSwitchingTabInternally = false;
            }
        }

        private void ApplyOfflineUiState(bool isOffline)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => ApplyOfflineUiState(isOffline));
                return;
            }

            if (_isOfflineUiLocked == isOffline)
                return;

            _isOfflineUiLocked = isOffline;

            if (isOffline)
            {
                _shouldOpenCertificatesTabWhenConnectionRestored = true;

                _isSwitchingTabInternally = true;
                try
                {
                    mainTabControl.SelectedItem = tabSettings;
                }
                finally
                {
                    _isSwitchingTabInternally = false;
                }

                tabCertificates.Visibility = Visibility.Collapsed;
                tabTokens.Visibility = Visibility.Collapsed;
                tabLogs.Visibility = Visibility.Collapsed;
                tabSettings.Visibility = Visibility.Visible;

                tabCertificates.IsEnabled = false;
                tabTokens.IsEnabled = false;
                tabLogs.IsEnabled = false;
                tabSettings.IsEnabled = true;

                txtServerStatus.Text = "🔴 Сервер OFFLINE";
                txtServerStatus.Foreground = Brushes.Red;
                txtNextCheck.Text = "Следующая проверка: --:--";
                txtNextCheck.Visibility = Visibility.Visible;

                if (statusText != null)
                    statusText.Text = "Нет связи с сервером. Доступны только настройки.";

                if (tokensStatusText != null)
                    tokensStatusText.Text = "Нет связи с сервером.";

                if (logStatusText != null)
                    logStatusText.Text = "Нет связи с сервером.";
            }
            else
            {
                tabCertificates.Visibility = Visibility.Visible;
                tabTokens.Visibility = Visibility.Visible;
                tabLogs.Visibility = Visibility.Visible;
                tabSettings.Visibility = Visibility.Visible;

                tabCertificates.IsEnabled = true;
                tabTokens.IsEnabled = true;
                tabLogs.IsEnabled = true;
                tabSettings.IsEnabled = true;

                if (_shouldOpenCertificatesTabWhenConnectionRestored)
                {
                    SelectCertificatesTab();
                    _shouldOpenCertificatesTabWhenConnectionRestored = false;
                }

                if (statusText != null)
                    statusText.Text = "Готов";

                if (tokensStatusText != null)
                    tokensStatusText.Text = "Токены загружены";

                if (logStatusText != null)
                    logStatusText.Text = "Логи загружены";
            }

            mainTabControl.UpdateLayout();
        }

        private async void BtnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var clientPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "client.settings.txt");

                SettingsSaver.SaveClient(clientPath, _clientSettings);

                await RecreateApiAndTokenServicesAsync();
                var loaded = await LoadFromServer(showErrorDialog: false);

                AddToMiniLog("Клиентские настройки сохранены");

                if (loaded)
                {
                    MessageBox.Show("Настройки сохранены.",
                        "Сохранено",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Настройки сохранены, но подключиться к серверу не удалось.\nПроверьте IP и порт.",
                        "Сохранено с предупреждением",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения:\n" + ex.Message);
            }
        }
    }
}