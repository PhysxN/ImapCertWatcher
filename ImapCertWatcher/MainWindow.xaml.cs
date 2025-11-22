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

                var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
                _settings = SettingsLoader.Load(settingsPath);

                this.DataContext = _settings;

                pwdMailPassword.Password = _settings.MailPassword;
                pwdFbPassword.Password = _settings.FbPassword;

                cmbImapFolder.ItemsSource = _availableFolders;
                txtInterval.Text = _settings.CheckIntervalSeconds.ToString();

                try
                {
                    _db = new DbHelper(_settings);
                    _watcher = new ImapWatcher(_settings, _db);

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
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            LoadFromDb();
        }

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

        // Обработчики для ComboBox здания
        private void BuildingComboBox_DropDownOpened(object sender, EventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                // Убеждаемся, что список заполнен
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
                // Сохраняем только если значение действительно изменилось
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
                    // Настраиваем ComboBox при подготовке к редактированию
                    comboBox.ItemsSource = _availableBuildings;
                    comboBox.DisplayMemberPath = ".";
                    comboBox.SelectedItem = record.Building;

                    Log($"Подготовка ComboBox для {record.Fio}, текущее значение: {record.Building}");
                }
            }
        }

        private void DgCerts_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (e.Column.Header.ToString() == "Здание")
            {
                Log("Начало редактирования ячейки здания");
            }
        }

        private void DgCerts_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit && e.Column.Header.ToString() == "Здание")
            {
                var comboBox = e.EditingElement as ComboBox;
                if (comboBox?.DataContext is Models.CertRecord record)
                {
                    Log($"Завершение редактирования для {record.Fio}, новое значение: {record.Building}");
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
                    Log($"Сохранено здание для {record.Fio}: {record.Building}");

                    // НЕ обновляем всю таблицу, чтобы не закрывать ComboBox
                    // Вместо этого просто обновим статус
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения здания: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                Log($"Ошибка сохранения здания: {ex.Message}");
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
                    e.Row.Background = System.Windows.Media.Brushes.LightCoral;
                }
                else if (record.DaysLeft <= 30)
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightGoldenrodYellow;
                }
                else
                {
                    e.Row.Background = System.Windows.Media.Brushes.LightGreen;
                }
            }
        }

        private void DgCerts_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (dgCerts.SelectedItem is Models.CertRecord record)
            {
                var menuItem = (MenuItem)dgCerts.ContextMenu.Items[0];

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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения настроек почты: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void Log(string message)
        {
            try
            {
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                string logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy-MM-dd}.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [MainWindow] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

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