using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Diagnostics;
using System.Security.Authentication;
using System.Windows.Controls.Primitives;

// Альтернатива Excel без использования Office Interop
using System.Data;
using System.Data.OleDb;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window
    {
        private AppSettings _settings;
        private DbHelper _db;
        private ImapNewCertificatesWatcher _newWatcher;
        private ImapRevocationsWatcher _revWatcher;
        private ObservableCollection<CertRecord> _items = new ObservableCollection<CertRecord>();
        private ObservableCollection<CertRecord> _allItems = new ObservableCollection<CertRecord>();
        private Timer _timer;
        private DispatcherTimer _refreshTimer;
        private ObservableCollection<string> _availableFolders = new ObservableCollection<string>();
        private ObservableCollection<string> _miniLogMessages = new ObservableCollection<string>();
        private const int MAX_MINI_LOG_LINES = 3;

        private string _currentBuildingFilter = "Все";
        private bool _showDeleted = false;
        private string _searchText = "";
        private bool _isDarkTheme = false;

        

        // Список доступных зданий
        private readonly ObservableCollection<string> _availableBuildings = new ObservableCollection<string>
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
                _settings = SettingsLoader.Load(settingsPath);

                this.DataContext = _settings;

                pwdMailPassword.Password = _settings.MailPassword;
                pwdFbPassword.Password = _settings.FbPassword;

                cmbImapFolder.ItemsSource = _availableFolders;
                txtInterval.Text = _settings.CheckIntervalHours.ToString();

                // Загрузка темы
                LoadThemeSettings();

                // Настройка DataGrid и обработчиков
                dgCerts.CanUserAddRows = false;
                dgCerts.ItemsSource = _items;

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

        

        private void OnProgressUpdated(string message, double progress)
        {
            ProgressUpdated?.Invoke(message, progress);
        }

        public event Action DataLoaded;

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Loaded -= MainWindow_Loaded;

            System.Diagnostics.Debug.WriteLine("MainWindow_Loaded начат");

            OnProgressUpdated("Инициализация базы данных и почтовых модулей...", 10);

            try
            {
                // ТЯЖЁЛАЯ ЧАСТЬ – в фоновом потоке
                var initResult = await Task.Run(() =>
                {
                    OnProgressUpdated("Создание подключения к БД...", 20);
                    var db = new DbHelper(_settings, AddToMiniLog);

                    OnProgressUpdated("Загрузка данных из БД...", 40);
                    var list = db.LoadAll(_showDeleted, _currentBuildingFilter);

                    OnProgressUpdated("Инициализация почтовых модулей...", 60);
                    var newWatcher = new ImapNewCertificatesWatcher(_settings, db, AddToMiniLog);
                    var revWatcher = new ImapRevocationsWatcher(_settings, db, AddToMiniLog);

                    OnProgressUpdated("Завершение инициализации...", 80);
                    return (db, list, newWatcher, revWatcher);
                });

                System.Diagnostics.Debug.WriteLine("Фоновая инициализация завершена");

                // ЛЁГКАЯ ЧАСТЬ – уже в UI-потоке
                _db = initResult.db;
                _newWatcher = initResult.newWatcher;
                _revWatcher = initResult.revWatcher;

                // Заполняем коллекции для грида
                _allItems.Clear();
                foreach (var rec in initResult.list)
                {
                    _allItems.Add(rec);
                }

                System.Diagnostics.Debug.WriteLine($"Загружено записей: {_allItems.Count}");

                OnProgressUpdated("Применение фильтров...", 85);
                ApplySearchFilter();

                OnProgressUpdated("Настройка интерфейса...", 90);
                AutoFitDataGridColumns();

                // Инициализация таймера проверки почты
                _timer?.Dispose();
                if (_settings.CheckIntervalHours > 0)
                {
                    var intervalMs = _settings.CheckIntervalHours * 60 * 60 * 1000;
                    _timer = new Timer(async _ => await DoCheckAsync(false),
                        null,
                        TimeSpan.FromMilliseconds(intervalMs),
                        TimeSpan.FromMilliseconds(intervalMs));
                }

                // Таймер для обновления дней
                InitializeRefreshTimer();

                OnProgressUpdated("Загрузка логов...", 95);
                LoadLogs();

                // ★ ВЫЗЫВАЕМ СОБЫТИЕ ПОЛНОЙ ЗАГРУЗКИ ДАННЫХ ★
                OnProgressUpdated("Готово", 100);

                System.Diagnostics.Debug.WriteLine("Вызываю DataLoaded...");

                // ★ УБИРАЕМ ЗАДЕРЖКУ - ЗАКРЫВАЕМ СРАЗУ ★
                DataLoaded?.Invoke();

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
        /// Обновление отображения оставшихся дней для всех записей
        /// </summary>
        private void RefreshDaysLeft(object sender, EventArgs e)
        {
            try
            {
                // Принудительно обновляем отображение DaysLeft для всех записей
                foreach (var item in _allItems)
                {
                    item.RefreshDaysLeft();
                }

                // Также обновляем отображение для отфильтрованных записей
                foreach (var item in _items)
                {
                    item.RefreshDaysLeft();
                }

                // Логируем только при отладке (можно закомментировать в продакшене)
                // AddToMiniLog("Авто-обновление: пересчитаны оставшиеся дни");
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

        private void ApplyTheme(bool isDark)
        {
            _isDarkTheme = isDark;

            var resources = this.Resources;

            if (isDark)
            {
                // Применяем темную тему
                resources["WindowBackgroundBrush"] = resources["DarkWindowBackground"];
                resources["ControlBackgroundBrush"] = resources["DarkControlBackground"];
                resources["TextColorBrush"] = resources["DarkTextColor"];
                resources["BorderColorBrush"] = resources["DarkBorderColor"];
                resources["AccentColorBrush"] = resources["DarkAccentColor"];
                resources["HoverColorBrush"] = resources["DarkHoverColor"];

                pwdMailPassword.Foreground = Brushes.White;
                pwdFbPassword.Foreground = Brushes.White;
            }
            else
            {
                // Применяем светлую тему
                resources["WindowBackgroundBrush"] = resources["LightWindowBackground"];
                resources["ControlBackgroundBrush"] = resources["LightControlBackground"];
                resources["TextColorBrush"] = resources["LightTextColor"];
                resources["BorderColorBrush"] = resources["LightBorderColor"];
                resources["AccentColorBrush"] = resources["LightAccentColor"];
                resources["HoverColorBrush"] = resources["LightHoverColor"];

                pwdMailPassword.Foreground = Brushes.Black;
                pwdFbPassword.Foreground = Brushes.Black;
            }

            try
            {
                resources[SystemColors.WindowBrushKey] = resources["ControlBackgroundBrush"];
                resources[SystemColors.ControlTextBrushKey] = resources["TextColorBrush"];
                resources[SystemColors.HighlightBrushKey] = resources["AccentColorBrush"];
                resources[SystemColors.HighlightTextBrushKey] = new SolidColorBrush(Colors.White);
            }
            catch { }

            // Принудительно обновляем визуал
            this.InvalidateVisual();

            // Обновляем строки DataGrid для применения новых цветов
            var tempItems = _items.ToList();
            _items.Clear();
            foreach (var item in tempItems)
                _items.Add(item);
        }

        private void ChkDarkTheme_Changed(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = chkDarkTheme.IsChecked == true;
            ApplyTheme(_isDarkTheme);
            SaveThemeSettings();
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

        private void LoadFromDb()
        {
            try
            {
                if (_db == null) return;

                var sw = Stopwatch.StartNew();
                var list = _db.LoadAll(_showDeleted, _currentBuildingFilter);
                sw.Stop();

                Log($"LoadFromDb: {list.Count} записей, showDeleted={_showDeleted}, " +
                    $"filter={_currentBuildingFilter}, время={sw.ElapsedMilliseconds} мс");

                Dispatcher.Invoke(() =>
                {
                    _allItems.Clear();
                    foreach (var e in list)
                    {
                        // Больше НЕ делаем _db.HasArchiveInDb(e.Id) по каждой строке
                        _allItems.Add(e);
                    }

                    ApplySearchFilter();
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка загрузки из БД: " + ex.Message);
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
                if (string.IsNullOrWhiteSpace(_searchText))
                {
                    _items.Clear();
                    foreach (var item in _allItems)
                    {
                        _items.Add(item);
                    }
                    searchStatusText.Text = $"Всего записей: {_items.Count}";
                }
                else
                {
                    var searchTerms = _searchText.ToLower().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    _items.Clear();
                    foreach (var item in _allItems)
                    {
                        bool matchesAllTerms = true;
                        string itemFio = (item.Fio ?? "").ToLower();

                        foreach (var term in searchTerms)
                        {
                            if (!itemFio.Contains(term))
                            {
                                matchesAllTerms = false;
                                break;
                            }
                        }

                        if (matchesAllTerms)
                        {
                            _items.Add(item);
                        }
                    }
                    searchStatusText.Text = $"Найдено: {_items.Count} из {_allItems.Count}";
                }

                // Автоподбор столбцов после применения фильтра
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    AutoFitDataGridColumns();
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                Log($"Ошибка применения фильтра поиска: {ex.Message}");
                searchStatusText.Text = "Ошибка поиска";
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

        private async Task DoCheckAsync(bool checkAllMessages)
        {
            if (_newWatcher == null || _revWatcher == null)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка: компоненты не инициализированы");
                return;
            }

            // Показываем прогресс-бар на всех вкладках
            ShowProgressBar(true);

            Dispatcher.Invoke(() =>
            {
                statusText.Text = checkAllMessages ? "Обработка всех писем..." : "Проверка почты...";
                AddToMiniLog(checkAllMessages ? "Начата обработка всех писем" : "Начата проверка почты");
            });

            int addedCount = 0;
            int updatedCount = 0;
            int revokedCount = 0;

            try
            {
                // Первый модуль: новые сертификаты (деструктуризация tuple)
                var (processedEntries, updatedDelta, addedDelta) =
                    await Task.Run(() => _newWatcher.ProcessNewCertificates(checkAllMessages));

                // Аккумулируем счётчики
                addedCount += addedDelta;
                updatedCount += updatedDelta;

                // Второй модуль: аннулирования — ожидаем, что метод возвращает int (число аннулирований)
                int revokedDelta = await Task.Run(() => _revWatcher.ProcessRevocations(checkAllMessages));
                revokedCount += revokedDelta;

                // Всегда обновляем данные из БД, даже если не было новых писем
                var newList = _db?.LoadAll(_showDeleted, _currentBuildingFilter);
                if (newList != null)
                {
                    Dispatcher.Invoke(() =>
                    {
                        _allItems.Clear();
                        foreach (var e in newList) _allItems.Add(e);
                        ApplySearchFilter();

                        // Автоподбор столбцов
                        AutoFitDataGridColumns();

                        if (updatedCount > 0 || addedCount > 0 || revokedCount > 0)
                        {
                            string message = "";
                            var parts = new List<string>();
                            if (addedCount > 0) parts.Add($"Добавлено: {addedCount}");
                            if (updatedCount > 0) parts.Add($"Обновлено: {updatedCount}");
                            if (revokedCount > 0) parts.Add($"Аннулировано: {revokedCount}");
                            message = string.Join(", ", parts) + $" ({DateTime.Now})";

                            statusText.Text = message;
                            AddToMiniLog(message);
                        }
                        else
                        {
                            statusText.Text = checkAllMessages
                                ? $"Обработка завершена: изменений нет ({DateTime.Now})"
                                : $"Проверка завершена: новых писем нет ({DateTime.Now})";
                            AddToMiniLog(checkAllMessages
                                ? "Обработка завершена: изменений нет"
                                : "Проверка завершена: новых писем нет");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    statusText.Text = "Ошибка при проверке почты";
                    AddToMiniLog($"Ошибка проверки: {ex.Message}");
                    Log($"Ошибка при проверке почты: {ex.Message}");
                });
            }
            finally
            {
                // Скрываем прогресс-бар на всех вкладках
                ShowProgressBar(false);
            }
        }

        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            Log("Ручная проверка почты запущена");
            await DoCheckAsync(false);
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
            Log("Обработка всех писем запущена");
            await DoCheckAsync(true);
        }

        private void BuildingFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAllBuildings.IsChecked == true)
                _currentBuildingFilter = "Все";
            else if (rbKrasnoflotskaya.IsChecked == true)
                _currentBuildingFilter = "Краснофлотская";
            else if (rbPionerskaya.IsChecked == true)
                _currentBuildingFilter = "Пионерская";

            LoadFromDb();
            AddToMiniLog($"Фильтр: {_currentBuildingFilter}");
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            LoadFromDb();
            AddToMiniLog(_showDeleted ? "Показаны удаленные" : "Скрыты удаленные");
        }

        private void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is CertRecord record)
            {
                try
                {
                    _db.UpdateNote(record.Id, textBox.Text);
                    statusText.Text = "Примечание сохранено";
                    AddToMiniLog($"Обновлено примечание: {record.Fio}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения примечания: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка примечания: {ex.Message}");
                }
            }
        }

        private void BuildingComboBox_DropDownOpened(object sender, EventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                if (comboBox.ItemsSource == null)
                {
                    comboBox.ItemsSource = _availableBuildings;
                }
            }
        }

        private void BuildingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSavingBuilding) return;

            if (sender is ComboBox comboBox && comboBox.DataContext is CertRecord record)
            {
                if (e.AddedItems.Count > 0)
                {
                    string newValue = e.AddedItems[0] as string;
                    if (newValue != record.Building)
                    {
                        _isSavingBuilding = true;
                        try
                        {
                            SaveBuilding(record);
                        }
                        finally
                        {
                            _isSavingBuilding = false;
                        }
                    }
                }
            }
        }

        private void BuildingComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isSavingBuilding) return;

            if (sender is ComboBox comboBox && comboBox.DataContext is CertRecord record)
            {
                _isSavingBuilding = true;
                try
                {
                    SaveBuilding(record);
                }
                finally
                {
                    _isSavingBuilding = false;
                }
            }
        }

        private void DgCerts_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column.Header.ToString() == "Здание")
            {
                var comboBox = e.EditingElement as ComboBox;
                if (comboBox != null && e.Row.DataContext is CertRecord record)
                {
                    comboBox.ItemsSource = _availableBuildings;
                    comboBox.DisplayMemberPath = ".";
                    comboBox.SelectedItem = record.Building;
                }
            }
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

        private void SaveBuilding(CertRecord record)
        {
            try
            {
                if (record.Building != null)
                {
                    _db.UpdateBuilding(record.Id, record.Building);
                    statusText.Text = $"Здание сохранено: {record.Building}";
                    AddToMiniLog($"Обновлено здание: {record.Fio} -> {record.Building}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения здания: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                AddToMiniLog($"Ошибка здания: {ex.Message}");
            }
        }

        private void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                try
                {
                    Log($"Пометка на удаление: {record.Fio} (ID: {record.Id})");
                    _db.MarkAsDeleted(record.Id, true);
                    LoadFromDb();
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

        private void BtnRestoreSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                try
                {
                    _db.MarkAsDeleted(record.Id, false);
                    LoadFromDb();
                    statusText.Text = "Запись восстановлена";
                    AddToMiniLog($"Восстановлен: {record.Fio}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при восстановлении: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка восстановления: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для восстановления", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnOpenArchive_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is CertRecord record)
            {
                try
                {
                    AddToMiniLog($"Попытка открыть архив для: {record.Fio}");

                    // ★ ПРОВЕРЯЕМ АРХИВ В БД В ПЕРВУЮ ОЧЕРЕДЬ ★
                    if (_db.HasArchiveInDb(record.Id))
                    {
                        AddToMiniLog($"Найден архив в БД для ID: {record.Id}");
                        var (fileData, fileName) = _db.GetArchiveFromDb(record.Id);
                        if (fileData != null && fileData.Length > 0)
                        {
                            AddToMiniLog($"Загружен архив из БД: {fileName} ({fileData.Length} байт)");

                            // Создаем папку Certs в папке программы, если её нет
                            string certsBaseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs");
                            if (!Directory.Exists(certsBaseDir))
                                Directory.CreateDirectory(certsBaseDir);

                            // Создаем папку для конкретного пользователя
                            string userFolder = Path.Combine(certsBaseDir, MakeValidFolderName(record.Fio));
                            if (!Directory.Exists(userFolder))
                                Directory.CreateDirectory(userFolder);

                            // Сохраняем временный архивный файл
                            string tempArchivePath = Path.Combine(userFolder, fileName ?? "archive.zip");
                            File.WriteAllBytes(tempArchivePath, fileData);

                            AddToMiniLog($"Архив сохранен во временный файл: {tempArchivePath}");

                            // Распаковываем архив
                            if (ExtractArchiveToFolder(tempArchivePath, userFolder))
                            {
                                // Удаляем временный архив после распаковки
                                try { File.Delete(tempArchivePath); } catch { }

                                // Открываем папку с распакованными файлами
                                Process.Start("explorer.exe", userFolder);

                                statusText.Text = "Архив распакован и папка открыта";
                                AddToMiniLog($"Архив распакован и открыт: {record.Fio}");
                            }
                            else
                            {
                                MessageBox.Show("Не удалось распаковать архив из БД", "Ошибка",
                                              MessageBoxButton.OK, MessageBoxImage.Error);
                                AddToMiniLog($"Ошибка распаковки архива из БД: {record.Fio}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Не удалось загрузить архив из БД", "Ошибка",
                                          MessageBoxButton.OK, MessageBoxImage.Error);
                            AddToMiniLog($"Ошибка загрузки архива из БД: данные пустые");
                        }
                    }
                    else
                    {
                        AddToMiniLog($"Архив не найден в БД для ID: {record.Id}");

                        // ★ ПРОВЕРЯЕМ СТАРУЮ ЛОГИКУ ДЛЯ ФАЙЛОВОЙ СИСТЕМЫ ★
                        if (!string.IsNullOrEmpty(record.ArchivePath) && File.Exists(record.ArchivePath))
                        {
                            AddToMiniLog($"Найден архив в файловой системе: {record.ArchivePath}");

                            // ★ Создаем папку Certs в папке программы, если её нет
                            string certsBaseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs");
                            if (!Directory.Exists(certsBaseDir))
                                Directory.CreateDirectory(certsBaseDir);

                            // Создаем папку для конкретного пользователя
                            string userFolder = Path.Combine(certsBaseDir, MakeValidFolderName(record.Fio));
                            if (!Directory.Exists(userFolder))
                                Directory.CreateDirectory(userFolder);

                            // Распаковываем архив
                            if (ExtractArchiveToFolder(record.ArchivePath, userFolder))
                            {
                                // Открываем папку с распакованными файлами
                                Process.Start("explorer.exe", userFolder);

                                statusText.Text = "Архив распакован и папка открыта";
                                AddToMiniLog($"Открыт архив из файловой системы: {record.Fio}");
                            }
                            else
                            {
                                MessageBox.Show("Не удалось распаковать архив из файловой системы", "Ошибка",
                                              MessageBoxButton.OK, MessageBoxImage.Error);
                                AddToMiniLog($"Ошибка распаковки архива из файловой системы: {record.Fio}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Архив не найден ни в базе данных, ни в файловой системе", "Ошибка",
                                          MessageBoxButton.OK, MessageBoxImage.Error);
                            AddToMiniLog($"Архив не найден нигде для: {record.Fio}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при работе с архивом: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка архива: {ex.Message}");
                }
            }
        }

        private void OpenArchive(CertRecord record)
        {
            try
            {
                // Пробуем распаковать локальный путь (если задан)
                if (!string.IsNullOrEmpty(record.ArchivePath) && File.Exists(record.ArchivePath))
                {
                    var safeFio = MakeValidFolderName(record.Fio);
                    string certsBaseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs");
                    if (!Directory.Exists(certsBaseDir)) Directory.CreateDirectory(certsBaseDir);
                    string userFolder = Path.Combine(certsBaseDir, safeFio);
                    if (!Directory.Exists(userFolder)) Directory.CreateDirectory(userFolder);

                    if (ExtractArchiveToFolder(record.ArchivePath, userFolder))
                    {
                        statusText.Text = "Архив распакован и открыт";
                        Process.Start("explorer.exe", userFolder);
                    }
                    else
                    {
                        MessageBox.Show("Не удалось распаковать архив", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Архив отсутствует или не найден", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с архивом: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
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
                        ? Color.FromArgb(255, 80, 80, 80)  // Темно-серый для темной темы
                        : Color.FromArgb(255, 220, 220, 220)); // Светло-серый для светлой темы
                }
                else if (record.DaysLeft <= 10)
                {
                    // Красный фон для срочных сертификатов
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 120, 50, 50)   // Темно-красный для темной темы
                        : Color.FromArgb(255, 255, 200, 200)); // Светло-красный для светлой темы
                }
                else if (record.DaysLeft <= 30)
                {
                    // Желтый фон для предупреждения
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 120, 120, 50)   // Темно-желтый для темной темы
                        : Color.FromArgb(255, 255, 255, 200)); // Светло-желтый для светлой темы
                }
                else
                {
                    // Зеленый фон для нормальных сертификатов
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 50, 80, 50)     // Темно-зеленый для темной темы
                        : Color.FromArgb(255, 200, 255, 200)); // Светло-зеленый для светлой темы
                }

                // Принудительно устанавливаем цвет текста для всей строки
                e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                    ? Colors.White : Colors.Black);
            }
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

        private void AddArchiveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                if (!string.IsNullOrEmpty(record.ArchivePath) && File.Exists(record.ArchivePath))
                {
                    MessageBox.Show("Архив уже прикреплен к этой записи", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var openFileDialog = new OpenFileDialog
                {
                    Filter = "ZIP архивы (*.zip)|*.zip",
                    Title = "Выберите архив с подписью"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    try
                    {
                        string selectedFilePath = openFileDialog.FileName;

                        var safeFio = MakeValidFileName(record.Fio);
                        var certFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs", safeFio);

                        if (!Directory.Exists(certFolder))
                            Directory.CreateDirectory(certFolder);

                        string fileName = $"{record.CertNumber}_{Path.GetFileName(selectedFilePath)}";
                        string destinationPath = Path.Combine(certFolder, fileName);

                        File.Copy(selectedFilePath, destinationPath, true);

                        _db.UpdateArchivePath(record.Id, destinationPath);

                        LoadFromDb();

                        statusText.Text = "Архив успешно прикреплен";
                        AddToMiniLog($"Добавлен архив для: {record.Fio}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при прикреплении архива: {ex.Message}", "Ошибка",
                                      MessageBoxButton.OK, MessageBoxImage.Error);
                        AddToMiniLog($"Ошибка добавления архива: {ex.Message}");
                    }
                }
            }
        }

        private string MakeValidFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(name.Where(ch => !invalidChars.Contains(ch)).ToArray());
        }

        private void BtnDialectHelp_Click(object sender, RoutedEventArgs e)
        {
            string helpText = @"Диалект Firebird - версия SQL-синтаксиса:

• Диалект 3 (рекомендуется) - современный SQL
  - Полная поддержка Unicode
  - Раздельные типы DATE/TIME/TIMESTAMP
  - Современные возможности

• Диалект 1 - устаревший синтаксис
  - Только для совместимости со старыми базами
  - Ограниченная поддержка Unicode

Обычно используется диалект 3. 
Меняйте только при ошибках совместимости.";

            MessageBox.Show(helpText, "Справка: Диалект Firebird",
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnRefreshFolders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_settings.MailHost) || string.IsNullOrEmpty(_settings.MailLogin))
                {
                    MessageBox.Show("Заполните хост и логин перед загрузкой папок", "Предупреждение",
                                  MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Cursor = Cursors.Wait;
                btnRefreshFolders.IsEnabled = false;
                txtFoldersStatus.Text = "Загрузка списка папок...";

                Task.Run(() => LoadMailFolders());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке папок: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                ResetUIState();
            }
        }

        private void LoadMailFolders()
        {
            try
            {
                var tempSettings = new AppSettings
                {
                    MailHost = _settings.MailHost,
                    MailPort = _settings.MailPort,
                    MailUseSsl = _settings.MailUseSsl,
                    MailLogin = _settings.MailLogin,
                    MailPassword = pwdMailPassword.Password
                };

                var folders = GetMailFolders(tempSettings);

                Dispatcher.Invoke(() =>
                {
                    _availableFolders.Clear();
                    foreach (var folder in folders)
                    {
                        _availableFolders.Add(folder);
                    }

                    if (!string.IsNullOrEmpty(_settings.ImapFolder) &&
                        _availableFolders.Contains(_settings.ImapFolder))
                    {
                        cmbImapFolder.SelectedItem = _settings.ImapFolder;
                    }
                    else if (_availableFolders.Count > 0)
                    {
                        cmbImapFolder.SelectedIndex = 0;
                    }

                    txtFoldersStatus.Text = $"Загружено папок: {folders.Count}";
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"Ошибка загрузки папок: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    txtFoldersStatus.Text = "Ошибка загрузки папок";
                });
            }
            finally
            {
                Dispatcher.Invoke(ResetUIState);
            }
        }

        // Получить список папок с сервера (локально в UI)
        private List<string> GetMailFolders(AppSettings tempSettings)
        {
            var result = new List<string>();
            try
            {
                using (var client = new ImapClient())
                {
                    client.Timeout = 60000;
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    if (tempSettings.MailUseSsl)
                        client.Connect(tempSettings.MailHost, tempSettings.MailPort, SecureSocketOptions.SslOnConnect);
                    else
                        client.Connect(tempSettings.MailHost, tempSettings.MailPort, SecureSocketOptions.StartTlsWhenAvailable);

                    client.Authenticate(tempSettings.MailLogin, tempSettings.MailPassword);

                    // Получаем корневые папки пользователя
                    foreach (var ns in client.PersonalNamespaces)
                    {
                        try
                        {
                            var root = client.GetFolder(ns);
                            if (root != null)
                            {
                                result.Add(root.FullName);

                                try
                                {
                                    foreach (var f in root.GetSubfolders(false))
                                    {
                                        result.Add(f.FullName);
                                    }
                                }
                                catch { }
                            }
                        }
                        catch
                        {
                            // игнорируем ошибку конкретного namespace
                        }
                    }

                    // в качестве fallback: список top-level (если есть первый namespace)
                    try
                    {
                        var firstNs = client.PersonalNamespaces.FirstOrDefault();
                        if (firstNs != null)
                        {
                            foreach (var f in client.GetFolders(firstNs))
                            {
                                result.Add(f.FullName);
                            }
                        }
                    }
                    catch { }

                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                AddToMiniLog($"Ошибка получения списка папок: {ex.Message}");
            }

            return result.Distinct().ToList();
        }

        private void ResetUIState()
        {
            Cursor = Cursors.Arrow;
            btnRefreshFolders.IsEnabled = true;
        }

        private void BtnTestMailConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = Cursors.Wait;

                var tempSettings = new AppSettings
                {
                    MailHost = _settings.MailHost,
                    MailPort = _settings.MailPort,
                    MailUseSsl = _settings.MailUseSsl,
                    MailLogin = _settings.MailLogin,
                    MailPassword = pwdMailPassword.Password,
                    ImapFolder = cmbImapFolder.Text
                };

                var folders = GetMailFolders(tempSettings);

                string selectedFolder = cmbImapFolder.Text;
                bool folderExists = folders.Contains(selectedFolder);

                if (folderExists)
                {
                    MessageBox.Show($"Подключение к почтовому серверу успешно!\n" +
                                  $"Выбранная папка '{selectedFolder}' доступна.\n" +
                                  $"Всего папок на сервере: {folders.Count}",
                                  "Тест подключения",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"Подключение установлено, но папка '{selectedFolder}' не найдена.\n" +
                                  $"Доступные папки: {string.Join(", ", folders.Take(10))}" +
                                  (folders.Count > 10 ? $"\n... и еще {folders.Count - 10} папок" : ""),
                                  "Предупреждение",
                                  MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к почтовому серверу: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        private void BtnSaveMailSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.MailPassword = pwdMailPassword.Password;
                _settings.ImapFolder = cmbImapFolder.Text;

                SaveSettings();
                MessageBox.Show("Настройки почты сохранены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                AddToMiniLog("Сохранены настройки почты");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения настроек почты: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                AddToMiniLog($"Ошибка настроек почты: {ex.Message}");
            }
        }

        private void BtnSaveDbSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.FbPassword = pwdFbPassword.Password;

                SaveSettings();

                try
                {
                    _db = new DbHelper(_settings);
                    // пересоздаём watcher'ы с новой ссылкой на БД
                    _newWatcher = new ImapNewCertificatesWatcher(_settings, _db, AddToMiniLog);
                    _revWatcher = new ImapRevocationsWatcher(_settings, _db, AddToMiniLog);
                    LoadFromDb();
                }
                catch (Exception dbEx)
                {
                    MessageBox.Show($"Настройки сохранены, но ошибка переподключения к БД: {dbEx.Message}",
                                  "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    AddToMiniLog($"Ошибка переподключения БД: {dbEx.Message}");
                }

                MessageBox.Show("Настройки базы данных сохранены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                AddToMiniLog("Сохранены настройки БД");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения настроек БД: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                AddToMiniLog($"Ошибка настроек БД: {ex.Message}");
            }
        }

        private void BtnTestDbConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var tempSettings = new AppSettings
                {
                    FirebirdDbPath = _settings.FirebirdDbPath,
                    FbServer = _settings.FbServer,
                    FbUser = _settings.FbUser,
                    FbPassword = pwdFbPassword.Password,
                    FbDialect = _settings.FbDialect
                };

                var tempDb = new DbHelper(tempSettings);
                using (var conn = tempDb.GetConnection())
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = 'CERTS'";
                        var result = cmd.ExecuteScalar();
                    }
                }

                MessageBox.Show("Подключение к базе данных успешно!", "Тест подключения",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к базе данных: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBrowseDbPath_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Firebird Database (*.fdb)|*.fdb",
                DefaultExt = ".fdb",
                FileName = "certs.fdb"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                _settings.FirebirdDbPath = saveFileDialog.FileName;
                txtFirebirdDbPath.GetBindingExpression(TextBox.TextProperty).UpdateTarget();
            }
        }

        private void SaveSettings()
        {
            try
            {
                var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");

                var lines = new[]
                {
                    "# Mail settings",
                    $"MailHost={_settings.MailHost}",
                    $"MailPort={_settings.MailPort}",
                    $"MailUseSsl={_settings.MailUseSsl}",
                    $"MailLogin={_settings.MailLogin}",
                    $"MailPassword={_settings.MailPassword}",
                    $"ImapFolder={_settings.ImapFolder}",
                    "",
                    "# Filter",
                    $"FilterRecipient={_settings.FilterRecipient}",
                    $"FilterSubjectPrefix={_settings.FilterSubjectPrefix}",
                    "",
                    "# Firebird settings",
                    $"FirebirdDbPath={_settings.FirebirdDbPath}",
                    $"FbServer={_settings.FbServer}",
                    $"FbUser={_settings.FbUser}",
                    $"FbPassword={_settings.FbPassword}",
                    $"FbDialect={_settings.FbDialect}",
                    "",
                    "# App behavior",
                    $"CheckIntervalHours={_settings.CheckIntervalHours}"
                };

                File.WriteAllLines(settingsPath, lines, System.Text.Encoding.UTF8);

                _timer?.Dispose();
                if (_settings.CheckIntervalHours > 0)
                {
                    // Конвертируем часы в миллисекунды для таймера
                    var intervalMs = _settings.CheckIntervalHours * 60 * 60 * 1000;
                    _timer = new Timer(async _ => await DoCheckAsync(false),
                        null, TimeSpan.Zero, TimeSpan.FromMilliseconds(intervalMs));
                }

                txtInterval.Text = _settings.CheckIntervalHours.ToString(); // Обновляем отображение
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось сохранить настройки: {ex.Message}", ex);
            }
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

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _timer?.Dispose();
            _refreshTimer?.Stop();
            if (_refreshTimer != null)
            {
                _refreshTimer.Tick -= RefreshDaysLeft;
            }
        }
    }
}