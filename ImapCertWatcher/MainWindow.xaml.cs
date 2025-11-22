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
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class MainWindow : Window
    {
        private AppSettings _settings;
        private DbHelper _db;
        private ImapWatcher _watcher;
        private ObservableCollection<Models.CertRecord> _items = new ObservableCollection<Models.CertRecord>();
        private Timer _timer;
        private ObservableCollection<string> _availableFolders = new ObservableCollection<string>();

        private string _currentBuildingFilter = "Все";
        private bool _showDeleted = false;

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                // УБИРАЕМ регистрацию конвертера в ресурсах

                var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
                _settings = SettingsLoader.Load(settingsPath);

                // Установка DataContext для привязки данных
                this.DataContext = _settings;

                // Инициализация паролей (они не сохраняются в привязке)
                pwdMailPassword.Password = _settings.MailPassword;
                pwdFbPassword.Password = _settings.FbPassword;

                // Настройка ComboBox для папок
                cmbImapFolder.ItemsSource = _availableFolders;

                txtInterval.Text = _settings.CheckIntervalSeconds.ToString();

                // Инициализация БД с обработкой ошибок
                try
                {
                    _db = new DbHelper(_settings);
                    _watcher = new ImapWatcher(_settings, _db);

                    // Убираем пустую строку в конце DataGrid
                    dgCerts.CanUserAddRows = false;
                    dgCerts.ItemsSource = _items;

                    // загрузим существующие записи
                    LoadFromDb();

                    // загрузим логи
                    LoadLogs();

                    // запустим таймер (проверка только последних писем)
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

        private void LoadFromDb()
        {
            try
            {
                if (_db == null) return;

                var list = _db.LoadAll(_showDeleted, _currentBuildingFilter);
                Dispatcher.Invoke(() =>
                {
                    _items.Clear();
                    foreach (var e in list) _items.Add(e);
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка загрузки из БД: " + ex.Message);
            }
        }

        private async Task DoCheckAsync(bool checkAllMessages)
        {
            if (_watcher == null)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка: компоненты не инициализированы");
                return;
            }

            Dispatcher.Invoke(() => statusText.Text = checkAllMessages ? "Обработка всех писем..." : "Проверка почты...");

            try
            {
                var entries = await Task.Run(() => _watcher.CheckMail(checkAllMessages));
                if (entries != null && entries.Any())
                {
                    var newList = _db?.LoadAll(_showDeleted, _currentBuildingFilter);
                    if (newList != null)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            _items.Clear();
                            foreach (var e in newList) _items.Add(e);
                            statusText.Text = checkAllMessages
                                ? $"Обработано всех писем: {entries.Count} шт. ({DateTime.Now})"
                                : $"Найдено/обновлено: {entries.Count} шт. ({DateTime.Now})";
                        });
                    }
                }
                else
                {
                    Dispatcher.Invoke(() => statusText.Text = checkAllMessages
                        ? $"Обработка завершена: подходящих писем не найдено ({DateTime.Now})"
                        : $"Проверка завершена: новых писем нет ({DateTime.Now})");
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => statusText.Text = "Ошибка при проверке: " + ex.Message);
            }
        }

        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            await DoCheckAsync(false); // Только последние 3 дня
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
            await DoCheckAsync(true); // Все письма
        }

        // Обработчики фильтров
        private void BuildingFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAllBuildings.IsChecked == true)
                _currentBuildingFilter = "Все";
            else if (rbKrasnoflotskaya.IsChecked == true)
                _currentBuildingFilter = "Краснофлотская";
            else if (rbPionerskaya.IsChecked == true)
                _currentBuildingFilter = "Пионерская";

            LoadFromDb();
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            LoadFromDb();
        }

        // Обработчики для редактирования данных
        private void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is Models.CertRecord record)
            {
                try
                {
                    _db.UpdateNote(record.Id, textBox.Text);
                    statusText.Text = "Примечание сохранено";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения примечания: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BuildingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.DataContext is Models.CertRecord record)
            {
                try
                {
                    _db.UpdateBuilding(record.Id, comboBox.SelectedItem?.ToString() ?? "");
                    statusText.Text = "Здание сохранено";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения здания: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // Обработчики для удаления/восстановления
        private void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                try
                {
                    _db.MarkAsDeleted(record.Id, true);
                    LoadFromDb();
                    statusText.Text = "Запись помечена на удаление";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при пометке на удаление: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
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
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при восстановлении: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для восстановления", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Обработчики для работы с архивами
        private void DgCerts_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dgCerts.SelectedItem is Models.CertRecord record && !string.IsNullOrEmpty(record.ArchivePath))
            {
                OpenArchive(record);
            }
        }

        private void BtnOpenArchive_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is Models.CertRecord record)
            {
                OpenArchive(record);
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

        // Подсветка строк по количеству оставшихся дней
        private void DgCerts_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is Models.CertRecord record)
            {
                if (record.IsDeleted)
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightGray;
                }
                else if (record.DaysLeft <= 10)
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightCoral; // Красный
                }
                else if (record.DaysLeft <= 30)
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightGoldenrodYellow; // Оранжевый
                }
                else
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightGreen; // Зеленый
                }
            }
        }

        // Контекстное меню для добавления архива
        private void DgCerts_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                var menuItem = (MenuItem)dgCerts.ContextMenu.Items[0];

                // Если архив уже есть, запрещаем добавление
                if (!string.IsNullOrEmpty(record.ArchivePath) && File.Exists(record.ArchivePath))
                {
                    menuItem.IsEnabled = false;
                    menuItem.ToolTip = "Архив уже прикреплен";
                }
                else
                {
                    menuItem.IsEnabled = true;
                    menuItem.ToolTip = "Добавить архив с подписью";
                }
            }
        }

        private void AddArchiveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                // Проверяем, что архив еще не прикреплен
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

                        // Создаем безопасное имя папки для ФИО
                        var safeFio = MakeValidFileName(record.Fio);
                        var certFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs", safeFio);

                        if (!Directory.Exists(certFolder))
                            Directory.CreateDirectory(certFolder);

                        // Копируем файл в папку сертификата
                        string fileName = $"{record.CertNumber}_{Path.GetFileName(selectedFilePath)}";
                        string destinationPath = Path.Combine(certFolder, fileName);

                        File.Copy(selectedFilePath, destinationPath, true);

                        // Обновляем запись в БД
                        _db.UpdateArchivePath(record.Id, destinationPath);

                        // Обновляем отображение
                        LoadFromDb();

                        statusText.Text = "Архив успешно прикреплен";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при прикреплении архива: {ex.Message}", "Ошибка",
                                      MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private string MakeValidFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(name.Where(ch => !invalidChars.Contains(ch)).ToArray());
        }

        // Загрузка списка папок с сервера
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
                txtFoldersStatus.Foreground = System.Windows.Media.Brushes.Blue;

                // Используем Task для асинхронной загрузки без блокировки UI
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

                // Обновляем UI в основном потоке
                Dispatcher.Invoke(() =>
                {
                    _availableFolders.Clear();
                    foreach (var folder in folders)
                    {
                        _availableFolders.Add(folder);
                    }

                    // Выбираем текущую папку, если она есть в списке
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
                    txtFoldersStatus.Foreground = System.Windows.Media.Brushes.Green;
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"Ошибка загрузки папок: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    txtFoldersStatus.Text = "Ошибка загрузки папок";
                    txtFoldersStatus.Foreground = System.Windows.Media.Brushes.Red;
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

        // Тест подключения к почте
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

        // Сохранение настроек почты
        private void BtnSaveMailSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.MailPassword = pwdMailPassword.Password;
                _settings.ImapFolder = cmbImapFolder.Text;

                SaveSettings();
                MessageBox.Show("Настройки почты сохранены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения настроек почты: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Сохранение настроек БД
        private void BtnSaveDbSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.FbPassword = pwdFbPassword.Password;

                SaveSettings();

                // Переинициализируем БД с новыми настройками
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
                }

                MessageBox.Show("Настройки базы данных сохранены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения настроек БД: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Тест подключения к БД
        private void BtnTestDbConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var tempSettings = new AppSettings
                {
                    FirebirdDbPath = _settings.FirebirdDbPath,
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

        // Выбор пути к файлу БД
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

        // Сохранение всех настроек в файл
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
                    $"FbUser={_settings.FbUser}",
                    $"FbPassword={_settings.FbPassword}",
                    $"FbDialect={_settings.FbDialect}",
                    "",
                    "# App behavior",
                    $"CheckIntervalSeconds={_settings.CheckIntervalSeconds}"
                };

                File.WriteAllLines(settingsPath, lines, System.Text.Encoding.UTF8);

                // Перезапускаем таймер с новым интервалом
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

        // Методы для работы с логами
        private void LoadLogs()
        {
            try
            {
                var logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logDirectory))
                {
                    txtLogs.Text = "Папка логов не существует";
                    return;
                }

                var logFiles = Directory.GetFiles(logDirectory, "*.log")
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .ToList();

                if (!logFiles.Any())
                {
                    txtLogs.Text = "Логи не найдены";
                    return;
                }

                // Читаем последний лог-файл
                var latestLog = logFiles.First();
                var logContent = File.ReadAllText(latestLog);

                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = logContent;
                    logStatusText.Text = $"Загружен файл: {Path.GetFileName(latestLog)}";
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    txtLogs.Text = $"Ошибка загрузки логов: {ex.Message}";
                    logStatusText.Text = "Ошибка загрузки логов";
                });
            }
        }

        private void BtnRefreshLogs_Click(object sender, RoutedEventArgs e)
        {
            LoadLogs();
        }

        private void BtnClearLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logDirectory))
                {
                    MessageBox.Show("Папка логов не существует", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var cutoffDate = DateTime.Now.AddMonths(-2);
                var logFiles = Directory.GetFiles(logDirectory, "*.log");
                int deletedCount = 0;

                foreach (var logFile in logFiles)
                {
                    if (File.GetCreationTime(logFile) < cutoffDate)
                    {
                        File.Delete(logFile);
                        deletedCount++;
                    }
                }

                MessageBox.Show($"Удалено логов: {deletedCount}", "Очистка логов",
                              MessageBoxButton.OK, MessageBoxImage.Information);

                LoadLogs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке логов: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnOpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                System.Diagnostics.Process.Start("explorer.exe", logDirectory);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии папки логов: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _timer?.Dispose();
        }
    }
}