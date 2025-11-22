using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using System;
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

// Альтернатива Excel без использования Office Interop
using System.Data;
using System.Data.OleDb;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window
    {
        private AppSettings _settings;
        private DbHelper _db;
        private ImapWatcher _watcher;
        private ObservableCollection<Models.CertRecord> _items = new ObservableCollection<Models.CertRecord>();
        private ObservableCollection<Models.CertRecord> _allItems = new ObservableCollection<Models.CertRecord>();
        private Timer _timer;
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

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                // Инициализация мини-лога
                _miniLogMessages = new ObservableCollection<string>();
                AddToMiniLog("Приложение запущено");

                var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
                _settings = SettingsLoader.Load(settingsPath);

                this.DataContext = _settings;

                pwdMailPassword.Password = _settings.MailPassword;
                pwdFbPassword.Password = _settings.FbPassword;

                cmbImapFolder.ItemsSource = _availableFolders;
                txtInterval.Text = _settings.CheckIntervalSeconds.ToString();

                // Загрузка темы
                LoadThemeSettings();

                try
                {
                    // Передаем ссылку на метод AddToMiniLog
                    _db = new DbHelper(_settings, AddToMiniLog);
                    _watcher = new ImapWatcher(_settings, _db, AddToMiniLog);

                    dgCerts.CanUserAddRows = false;
                    dgCerts.ItemsSource = _items;

                    dgCerts.CellEditEnding += DgCerts_CellEditEnding;
                    dgCerts.BeginningEdit += DgCerts_BeginningEdit;
                    dgCerts.PreparingCellForEdit += DgCerts_PreparingCellForEdit;

                    LoadFromDb();
                    LoadLogs();

                    _timer = new Timer(async _ => await DoCheckAsync(false),
                        null, TimeSpan.Zero, TimeSpan.FromSeconds(_settings.CheckIntervalSeconds));
                }
                catch (Exception dbEx)
                {
                    MessageBox.Show($"Ошибка инициализации базы данных: {dbEx.Message}\n\n" +
                                   "Проверьте настройки подключения к Firebird на вкладке 'Настройки БД'.",
                                   "Ошибка БД", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации приложения: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
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

                    // Обновляем текстовый блок только если он инициализирован
                    if (txtMiniLog != null)
                    {
                        UpdateMiniLogDisplay();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Мини-лог (txtMiniLog null): {message}");
                    }
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
                if (txtMiniLog != null)
                {
                    txtMiniLog.Text = string.Join(Environment.NewLine, _miniLogMessages);
                }
                else
                {
                    // Если txtMiniLog еще не инициализирован, просто выводим в Debug
                    System.Diagnostics.Debug.WriteLine("Мини-лог обновлен: " + string.Join("; ", _miniLogMessages));
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
            }

            // Принудительно обновляем стили и перезагружаем данные
            this.InvalidateVisual();

            // Обновляем строки DataGrid для применения новых цветов
            var tempItems = _items.ToList();
            _items.Clear();
            foreach (var item in tempItems)
            {
                _items.Add(item);
            }
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

        private async Task ExportToExcel(bool exportSelectedOnly)
        {
            try
            {
                exportProgress.Visibility = Visibility.Visible;
                statusText.Text = "Подготовка экспорта...";
                AddToMiniLog("Начат экспорт в Excel");

                var dataToExport = exportSelectedOnly && dgCerts.SelectedItem != null
                    ? new ObservableCollection<Models.CertRecord> { (Models.CertRecord)dgCerts.SelectedItem }
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



        private void GenerateCsvFile(ObservableCollection<Models.CertRecord> data, string filePath)
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

        private void GenerateExcelFile(ObservableCollection<Models.CertRecord> data, string filePath)
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

        private void TryGenerateExcelWithClosedXML(ObservableCollection<Models.CertRecord> data, string filePath)
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

        private void GenerateExcelWithClosedXML(ObservableCollection<Models.CertRecord> data, string filePath)
        {
            // Этот метод будет работать только если установлен ClosedXML
            dynamic workbook = Activator.CreateInstance(Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML"));
            dynamic worksheet = workbook.Worksheets.Add("Сертификаты ЭЦП");

            // Заголовок
            worksheet.Cell(1, 1).Value = "Экспорт сертификатов ЭЦП";
            worksheet.Range(1, 1, 1, 8).Merge();
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 14;
            worksheet.Cell(1, 1).Style.Alignment.Horizontal = (dynamic)Enum.Parse(Type.GetType("ClosedXML.Excel.XLAlignmentHorizontalValues, ClosedXML"), "Center");

            // Заголовки таблицы
            string[] headers = { "ФИО", "Номер сертификата", "Дата начала", "Дата окончания",
                               "Осталось дней", "Здание", "Примечание", "Статус" };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(3, i + 1).Value = headers[i];
                worksheet.Cell(3, i + 1).Style.Font.Bold = true;
                worksheet.Cell(3, i + 1).Style.Fill.BackgroundColor = (dynamic)Enum.Parse(Type.GetType("ClosedXML.Excel.XLColor, ClosedXML"), "LightGray");
                worksheet.Cell(3, i + 1).Style.Border.OutsideBorder = (dynamic)Enum.Parse(Type.GetType("ClosedXML.Excel.XLBorderStyleValues, ClosedXML"), "Thin");
            }

            // Данные
            int row = 4;
            foreach (var record in data)
            {
                worksheet.Cell(row, 1).Value = record.Fio;
                worksheet.Cell(row, 2).Value = record.CertNumber;
                worksheet.Cell(row, 3).Value = record.DateStart.ToString("dd.MM.yyyy");
                worksheet.Cell(row, 4).Value = record.DateEnd.ToString("dd.MM.yyyy");
                worksheet.Cell(row, 5).Value = record.DaysLeft;
                worksheet.Cell(row, 6).Value = record.Building;
                worksheet.Cell(row, 7).Value = record.Note;
                worksheet.Cell(row, 8).Value = record.IsDeleted ? "Удален" : "Активен";

                // Цветовое кодирование по сроку действия
                string colorName;
                if (record.DaysLeft <= 10)
                {
                    colorName = "LightCoral";
                }
                else if (record.DaysLeft <= 30)
                {
                    colorName = "LightYellow";
                }
                else
                {
                    colorName = "LightGreen";
                }

                worksheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = (dynamic)Enum.Parse(Type.GetType("ClosedXML.Excel.XLColor, ClosedXML"), colorName);

                // Границы для ячеек
                for (int col = 1; col <= headers.Length; col++)
                {
                    worksheet.Cell(row, col).Style.Border.OutsideBorder = (dynamic)Enum.Parse(Type.GetType("ClosedXML.Excel.XLBorderStyleValues, ClosedXML"), "Thin");
                }

                row++;
            }

            // Автоподбор ширины столбцов
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
        }

        private void LoadFromDb()
        {
            try
            {
                if (_db == null) return;

                var list = _db.LoadAll(_showDeleted, _currentBuildingFilter);
                Dispatcher.Invoke(() =>
                {
                    _allItems.Clear();
                    foreach (var e in list) _allItems.Add(e);

                    ApplySearchFilter();
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка загрузки из БД: " + ex.Message);
            }
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
            AddToMiniLog("Поиск очищен"); // ← ДОБАВИТЬ
        }

        private async Task DoCheckAsync(bool checkAllMessages)
        {
            if (_watcher == null)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка: компоненты не инициализированы");
                return;
            }

            // Показываем прогресс-бар
            Dispatcher.Invoke(() =>
            {
                statusText.Text = checkAllMessages ? "Обработка всех писем..." : "Проверка почты...";
                progressMailCheck.Visibility = Visibility.Visible;
                AddToMiniLog(checkAllMessages ? "Начата обработка всех писем" : "Начата проверка почты");
            });

            try
            {
                var entries = await Task.Run(() => _watcher.CheckMail(checkAllMessages));

                // Всегда обновляем данные из БД, даже если не было новых писем
                var newList = _db?.LoadAll(_showDeleted, _currentBuildingFilter);
                if (newList != null)
                {
                    Dispatcher.Invoke(() =>
                    {
                        _allItems.Clear();
                        foreach (var e in newList) _allItems.Add(e);
                        ApplySearchFilter();

                        if (entries != null && entries.Any())
                        {
                            statusText.Text = checkAllMessages
                                ? $"Обработано всех писем: {entries.Count} шт. ({DateTime.Now})"
                                : $"Найдено/обновлено: {entries.Count} шт. ({DateTime.Now})";
                            AddToMiniLog(checkAllMessages
                                ? $"Обработано писем: {entries.Count}"
                                : $"Найдено новых: {entries.Count}");
                        }
                        else
                        {
                            statusText.Text = checkAllMessages
                                ? $"Обработка завершена: подходящих писем не найдено ({DateTime.Now})"
                                : $"Проверка завершена: новых писем нет ({DateTime.Now})";
                            AddToMiniLog(checkAllMessages
                                ? "Подходящих писем не найдено"
                                : "Новых писем нет");
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
                // Скрываем прогресс-бар
                Dispatcher.Invoke(() =>
                {
                    progressMailCheck.Visibility = Visibility.Collapsed;
                });
            }
        }

        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            await DoCheckAsync(false);
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
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
            AddToMiniLog($"Фильтр: {_currentBuildingFilter}"); // ← ДОБАВИТЬ
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            LoadFromDb();
            AddToMiniLog(_showDeleted ? "Показаны удаленные" : "Скрыты удаленные"); // ← ДОБАВИТЬ
        }

        private void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is Models.CertRecord record)
            {
                try
                {
                    _db.UpdateNote(record.Id, textBox.Text);
                    statusText.Text = "Примечание сохранено";
                    AddToMiniLog($"Обновлено примечание: {record.Fio}"); // ← ДОБАВИТЬ
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

            if (sender is ComboBox comboBox && comboBox.DataContext is Models.CertRecord record)
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

            if (sender is ComboBox comboBox && comboBox.DataContext is Models.CertRecord record)
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
                if (comboBox != null && e.Row.DataContext is Models.CertRecord record)
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
                if (comboBox?.DataContext is Models.CertRecord record)
                {
                    // Не сохраняем здесь, т.к. уже сохранили в SelectionChanged
                }
            }
        }

        private void SaveBuilding(Models.CertRecord record)
        {
            try
            {
                if (record.Building != null)
                {
                    _db.UpdateBuilding(record.Id, record.Building);
                    statusText.Text = $"Здание сохранено: {record.Building}";
                    AddToMiniLog($"Обновлено здание: {record.Fio} -> {record.Building}"); // ← ДОБАВИТЬ
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
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                try
                {
                    _db.MarkAsDeleted(record.Id, true);
                    LoadFromDb();
                    statusText.Text = "Запись помечена на удаление";
                    AddToMiniLog($"Удален: {record.Fio}"); // ← ДОБАВИТЬ ЭТУ СТРОКУ
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при пометке на удаление: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка удаления: {ex.Message}"); // ← И ЭТУ ДЛЯ ОШИБОК
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
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                try
                {
                    _db.MarkAsDeleted(record.Id, false);
                    LoadFromDb();
                    statusText.Text = "Запись восстановлена";
                    AddToMiniLog($"Восстановлен: {record.Fio}"); // ← ДОБАВИТЬ
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
            if (sender is Button button && button.DataContext is Models.CertRecord record)
            {
                try
                {
                    if (_watcher.ExtractArchive(record.ArchivePath, record.Fio))
                    {
                        statusText.Text = "Архив распакован и открыт";
                        AddToMiniLog($"Открыт архив: {record.Fio}"); // ← ДОБАВИТЬ
                    }
                    else
                    {
                        MessageBox.Show("Не удалось распаковать архив", "Ошибка",
                                      MessageBoxButton.OK, MessageBoxImage.Error);
                        AddToMiniLog($"Ошибка открытия архива: {record.Fio}");
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

        private void OpenArchive(Models.CertRecord record)
        {
            try
            {
                if (_watcher.ExtractArchive(record.ArchivePath, record.Fio))
                {
                    statusText.Text = "Архив распакован и открыт";
                }
                else
                {
                    MessageBox.Show("Не удалось распаковать архив", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (e.Row.Item is Models.CertRecord record)
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
            if (dgCerts.SelectedItem is Models.CertRecord record)
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
            if (dgCerts.SelectedItem is Models.CertRecord record)
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
                        AddToMiniLog($"Добавлен архив для: {record.Fio}"); // ← ДОБАВИТЬ
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

                var folders = ImapWatcher.GetMailFolders(tempSettings);

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

                var folders = ImapWatcher.GetMailFolders(tempSettings);

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
                AddToMiniLog("Сохранены настройки почты"); // ← ДОБАВИТЬ
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
                    _watcher = new ImapWatcher(_settings, _db);
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
                AddToMiniLog("Сохранены настройки БД"); // ← ДОБАВИТЬ
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
                    $"CheckIntervalSeconds={_settings.CheckIntervalSeconds}"
                };

                File.WriteAllLines(settingsPath, lines, System.Text.Encoding.UTF8);

                _timer?.Dispose();
                if (_settings.CheckIntervalSeconds > 0)
                {
                    _timer = new Timer(async _ => await DoCheckAsync(false),
                        null, TimeSpan.Zero, TimeSpan.FromSeconds(_settings.CheckIntervalSeconds));
                }

                txtInterval.Text = _settings.CheckIntervalSeconds.ToString();
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
                    txtLogs.Text = "Папка логов не существует";
                    return;
                }

                // Ищем все папки с датами и сортируем по убыванию (самые новые первыми)
                var dateFolders = Directory.GetDirectories(logBaseDirectory)
                    .Where(dir => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(dir), @"^\d{4}-\d{2}-\d{2}$"))
                    .OrderByDescending(dir => dir)
                    .ToList();

                if (!dateFolders.Any())
                {
                    txtLogs.Text = "Логи не найдены";
                    return;
                }

                // Берем самую свежую папку с датой
                var latestDateFolder = dateFolders.First();

                // Ищем все log файлы в этой папке и сортируем по дате создания (новые первыми)
                var logFiles = Directory.GetFiles(latestDateFolder, "*.log")
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .ToList();

                if (!logFiles.Any())
                {
                    txtLogs.Text = "Логи не найдены в папке с датой";
                    return;
                }

                // Берем самый свежий лог-файл
                var latestLog = logFiles.First();
                var logContent = File.ReadAllText(latestLog);

                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = logContent;
                    logStatusText.Text = $"Загружен файл: {Path.GetFileName(latestDateFolder)}/{Path.GetFileName(latestLog)}";
                });

                AddToMiniLog($"Загружены логи: {Path.GetFileName(latestLog)}");
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = $"Ошибка загрузки логов: {ex.Message}";
                    logStatusText.Text = "Ошибка загрузки логов";
                });
                AddToMiniLog($"Ошибка загрузки логов: {ex.Message}");
            }
        }

        private void BtnRefreshLogs_Click(object sender, RoutedEventArgs e)
        {
            LoadLogs();
            AddToMiniLog("Логи обновлены"); // ← ДОБАВИТЬ
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
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                string logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [MainWindow] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

                // Также пишем в мини-лог
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
        }
    }
}