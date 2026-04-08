using ImapCertWatcher.Models;
using ImapCertWatcher.Server;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;

namespace ImapCertWatcher
{
    public partial class ServerSettingsWindow : Window, INotifyPropertyChanged
    {
        private readonly ServerSettings _settings;
        private readonly NotificationManager _notificationManager;
        private readonly ServerSettingsBackendService _backendService;

        private const string AUTOSTART_REG_PATH =
             @"Software\Microsoft\Windows\CurrentVersion\Run";

        private const string AUTOSTART_VALUE_NAME =
            "ImapCertWatcherServer";

        public ObservableCollection<string> ImapFolders { get; }
            = new ObservableCollection<string>();

        private bool _isFoldersLoading;
        public bool IsFoldersLoading
        {
            get => _isFoldersLoading;
            set
            {
                _isFoldersLoading = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsFoldersLoaded));
            }
        }

        public bool IsFoldersLoaded => !IsFoldersLoading;

        private string _autoStartStatusText = "Автозагрузка сервера: неизвестно";
        public string AutoStartStatusText
        {
            get => _autoStartStatusText;
            set
            {
                _autoStartStatusText = value;
                OnPropertyChanged();
            }
        }


        public ServerSettingsWindow(ServerSettings settings)
        {
            InitializeComponent();

            _settings = settings;
            DataContext = _settings;

            _notificationManager = new NotificationManager(_settings, s => { });
            _backendService = new ServerSettingsBackendService();

            MailPasswordBox.Password = _settings.MailPassword ?? "";
            DbPasswordBox.Password = _settings.FbPassword ?? "";
            BimoidPasswordBox.Password = _settings.BimoidPassword ?? "";

            RefreshAutoStartStatus();

            Loaded += async (_, __) => await AutoLoadFoldersAsync();
        }

        private string MakeFriendlyImapError(Exception ex)
        {
            var msg = ex.Message?.ToLower() ?? "";

            if (msg.Contains("invalid credentials") ||
                msg.Contains("authentication") ||
                msg.Contains("login"))
            {
                return "Неверный логин или пароль.\n\n" +
                       "Проверьте учетные данные почты.";
            }

            if (msg.Contains("imap is disabled"))
            {
                return "IMAP отключен в настройках почты.\n\n" +
                       "Включите IMAP в веб-интерфейсе почты.";
            }

            if (msg.Contains("ssl") || msg.Contains("tls"))
            {
                return "Ошибка SSL соединения.\n\n" +
                       "Проверьте порт и настройку SSL.";
            }

            if (msg.Contains("timeout"))
            {
                return "Сервер почты не отвечает.\n\n" +
                       "Проверьте интернет-соединение.";
            }

            if (msg.Contains("unable to connect") ||
                msg.Contains("connection refused"))
            {
                return "Не удалось подключиться к серверу.\n\n" +
                       "Проверьте адрес IMAP сервера и порт.";
            }

            return "Ошибка подключения к почте:\n\n" + ex.Message;
        }

        private async Task AutoLoadFoldersAsync()
        {
            IsFoldersLoading = true;

            try
            {
                var folders = await _backendService.LoadFoldersAsync(
                    _settings,
                    MailPasswordBox.Password);

                ApplyFolders(folders);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    MakeFriendlyImapError(ex),
                    "Ошибка IMAP",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            finally
            {
                IsFoldersLoading = false;
            }
        }

        private void ApplyFolders(string[] folders)
        {
            ImapFolders.Clear();

            foreach (var f in folders)
                ImapFolders.Add(f);

            cmbNewCertsFolder.SelectedItem = _settings.ImapNewCertificatesFolder;
            cmbRevocationsFolder.SelectedItem = _settings.ImapRevocationsFolder;
        }

        private void RefreshAutoStartStatus()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(AUTOSTART_REG_PATH, false))
                {
                    var value = key?.GetValue(AUTOSTART_VALUE_NAME) as string;

                    bool enabled = !string.IsNullOrWhiteSpace(value);

                    _settings.AutoStartServer = enabled;
                    AutoStartStatusText = enabled
                        ? "Автозагрузка сервера: включена"
                        : "Автозагрузка сервера: выключена";
                }
            }
            catch (Exception ex)
            {
                AutoStartStatusText = "Автозагрузка сервера: ошибка";
                MessageBox.Show(
                    "Ошибка чтения автозагрузки:\n\n" + ex.Message,
                    "Автозагрузка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void SaveServerSettingsSilently()
        {
            _settings.MailPassword = MailPasswordBox.Password;
            _settings.FbPassword = DbPasswordBox.Password;
            _settings.BimoidPassword = BimoidPasswordBox.Password;

            _settings.ImapNewCertificatesFolder = cmbNewCertsFolder.Text;
            _settings.ImapRevocationsFolder = cmbRevocationsFolder.Text;

            var settingsPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "server.settings.txt");

            SettingsSaver.SaveServer(settingsPath, _settings);
        }

        private void AddToAutoStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(AUTOSTART_REG_PATH))
                {
                    if (key == null)
                        throw new Exception("Не удалось открыть раздел автозагрузки.");

                    string exePath = Assembly.GetExecutingAssembly().Location;
                    string runValue = $"\"{exePath}\"";

                    key.SetValue(AUTOSTART_VALUE_NAME, runValue);
                }

                _settings.AutoStartServer = true;
                SaveServerSettingsSilently();
                RefreshAutoStartStatus();

                MessageBox.Show(
                    "Сервер добавлен в автозагрузку.",
                    "Автозагрузка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка добавления в автозагрузку:\n\n" + ex.Message,
                    "Автозагрузка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void RemoveFromAutoStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(AUTOSTART_REG_PATH))
                {
                    if (key == null)
                        throw new Exception("Не удалось открыть раздел автозагрузки.");

                    key.DeleteValue(AUTOSTART_VALUE_NAME, false);
                }

                _settings.AutoStartServer = false;
                SaveServerSettingsSilently();
                RefreshAutoStartStatus();

                MessageBox.Show(
                    "Сервер удалён из автозагрузки.",
                    "Автозагрузка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка удаления из автозагрузки:\n\n" + ex.Message,
                    "Автозагрузка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _settings.MailPassword = MailPasswordBox.Password;
            _settings.FbPassword = DbPasswordBox.Password;
            _settings.BimoidPassword = BimoidPasswordBox.Password;

            _settings.ImapNewCertificatesFolder = cmbNewCertsFolder.Text;
            _settings.ImapRevocationsFolder = cmbRevocationsFolder.Text;

            var settingsPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "server.settings.txt");

            SettingsSaver.SaveServer(settingsPath, _settings);

            MessageBox.Show(
                "Настройки сохранены.\nПерезапустите сервер.",
                "Сервер",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            DialogResult = true;
        }

        private async void TestMail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                await _backendService.TestMailAsync(
                    _settings,
                    MailPasswordBox.Password);

                MessageBox.Show(
                    "Подключение к почте успешно!",
                    "IMAP",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    MakeFriendlyImapError(ex),
                    "Ошибка IMAP",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }

        private void TestDb_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                _backendService.TestDb(
                    _settings,
                    DbPasswordBox.Password);

                MessageBox.Show(
                    "Подключение к Firebird успешно!",
                    "Firebird",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка подключения к Firebird:\n\n" + ex.Message,
                    "Firebird ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }

        private void BtnTestBimoid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.BimoidPassword = BimoidPasswordBox.Password;
                _notificationManager.SendTestMessage();

                MessageBox.Show(
                    "Тестовое сообщение отправлено.",
                    "Bimoid",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка отправки:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BtnTestNewUserNotify_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settings.BimoidPassword = BimoidPasswordBox.Password;

                var fake = new CertRecord
                {
                    Fio = "Тестовый пользователь",
                    CertNumber = "TEST-123456",
                    DateStart = DateTime.Today,
                    DateEnd = DateTime.Today.AddYears(1),
                    Building = "Тест"
                };

                _notificationManager.SendTestNewUserNotification(fake);

                MessageBox.Show(
                    "Тестовое уведомление об изменениях сертификата отправлено.",
                    "Тест",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка теста:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void BtnFullCertsRecheck_Click(object sender, RoutedEventArgs e)
        {
            if (!(Owner is ServerWindow owner))
            {
                MessageBox.Show("Окно сервера не найдено.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var r = MessageBox.Show(
                "Будет выполнен полный обход всей папки сертификатов.\n\nЗапустить?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                var result = await owner.RunFullCertificatesCheckFromSettingsAsync();

                MessageBox.Show(
                    "Команда завершена: " + result,
                    "Проверка сертификатов",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(
                    "Операция отменена.",
                    "Проверка сертификатов",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка:\n\n" + ex.Message,
                    "Проверка сертификатов",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }

        private async void BtnFullRevokesRecheck_Click(object sender, RoutedEventArgs e)
        {
            if (!(Owner is ServerWindow owner))
            {
                MessageBox.Show("Окно сервера не найдено.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var r = MessageBox.Show(
                "Будет выполнен полный обход всей папки аннулирований.\n\nЗапустить?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                var result = await owner.RunFullRevocationsCheckFromSettingsAsync();

                MessageBox.Show(
                    "Команда завершена: " + result,
                    "Проверка аннулирований",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(
                    "Операция отменена.",
                    "Проверка аннулирований",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка:\n\n" + ex.Message,
                    "Проверка аннулирований",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}