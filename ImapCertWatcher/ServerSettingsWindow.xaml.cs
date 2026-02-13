using ImapCertWatcher.Data;
using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ImapCertWatcher
{
    public partial class ServerSettingsWindow : Window, INotifyPropertyChanged
    {
        private readonly ServerSettings _settings;
        private readonly NotificationManager _notificationManager;
        
        // ============================
        // IMAP folders binding
        // ============================

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

        // кеш папок (общий на всё приложение)
        private static string[] _cachedImapFolders;

        // ============================
        // Constructor
        // ============================

        public ServerSettingsWindow(ServerSettings settings)
        {
            InitializeComponent();

            _settings = settings;

            DataContext = _settings;

            _notificationManager = new NotificationManager(_settings, s => { });

            MailPasswordBox.Password = _settings.MailPassword;
            DbPasswordBox.Password = _settings.FbPassword;

            Loaded += async (_, __) => await AutoLoadFoldersAsync();
        }

        // ============================
        // Auto IMAP folder loading
        // ============================

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

            // fallback
            return "Ошибка подключения к почте:\n\n" + ex.Message;
        }

        private async Task AutoLoadFoldersAsync()
        {
            if (_cachedImapFolders != null)
            {
                ApplyFolders(_cachedImapFolders);
                return;
            }

            IsFoldersLoading = true;

            try
            {
                var folders = await LoadFoldersFromServer();
                _cachedImapFolders = folders;
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

        private async Task<string[]> LoadFoldersFromServer()
        {
            var result = new ObservableCollection<string>();

            using (var client = new ImapClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                await client.ConnectAsync(
                    _settings.MailHost,
                    _settings.MailPort,
                    _settings.MailUseSsl
                        ? SecureSocketOptions.SslOnConnect
                        : SecureSocketOptions.StartTlsWhenAvailable);

                await client.AuthenticateAsync(
                    _settings.MailLogin,
                    MailPasswordBox.Password);

                var root = client.GetFolder(client.PersonalNamespaces[0]);

                var folders = await root.GetSubfoldersAsync(false);

                foreach (var folder in folders)
                {
                    await LoadFoldersRecursive(folder, result);
                }

                await client.DisconnectAsync(true);
            }

            return result.ToArray();
        }

        private async Task LoadFoldersRecursive(
    IMailFolder folder,
    ObservableCollection<string> acc)
        {
            acc.Add(folder.FullName);

            var subs = await folder.GetSubfoldersAsync(false);

            foreach (var sub in subs)
            {
                await LoadFoldersRecursive(sub, acc);
            }
        }

        private void ApplyFolders(string[] folders)
        {
            ImapFolders.Clear();

            foreach (var f in folders)
                ImapFolders.Add(f);

            // восстанавливаем сохранённые значения
            cmbNewCertsFolder.SelectedItem = _settings.ImapNewCertificatesFolder;
            cmbRevocationsFolder.SelectedItem = _settings.ImapRevocationsFolder;
        }

        // ============================
        // SAVE SETTINGS
        // ============================

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _settings.MailPassword = MailPasswordBox.Password;
            _settings.FbPassword = DbPasswordBox.Password;

            _settings.ImapNewCertificatesFolder = cmbNewCertsFolder.Text;
            _settings.ImapRevocationsFolder = cmbRevocationsFolder.Text;

            var settingsPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "server.settings.txt");

            SaveSettings(settingsPath, _settings);

            MessageBox.Show(
                "Настройки сохранены.\nПерезапустите сервер.",
                "Сервер",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            DialogResult = true;
        }

        // ============================
        // TEST MAIL
        // ============================

        private async void TestMail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                using (var client = new ImapClient())
                {
                    client.ServerCertificateValidationCallback =
                        (s, c, h, er) => true;

                    await client.ConnectAsync(
                        _settings.MailHost,
                        _settings.MailPort,
                        _settings.MailUseSsl
                            ? SecureSocketOptions.SslOnConnect
                            : SecureSocketOptions.StartTlsWhenAvailable);

                    await client.AuthenticateAsync(
                        _settings.MailLogin,
                        MailPasswordBox.Password);

                    await client.DisconnectAsync(true);
                }

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



        // ============================
        // TEST FIREBIRD
        // ============================

        private void TestDb_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                var testSettings = new ServerSettings
                {
                    FbServer = _settings.FbServer,
                    FirebirdDbPath = _settings.FirebirdDbPath,
                    FbUser = _settings.FbUser,
                    FbPassword = DbPasswordBox.Password,
                    FbDialect = _settings.FbDialect
                };

                var db = new DbHelper(testSettings);

                using (var conn = db.GetConnection())
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT 1 FROM RDB$DATABASE";
                        cmd.ExecuteScalar();
                    }
                }

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

        // ============================
        // BIMOID TEST
        // ============================

        private void BtnTestBimoid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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
                    "Тестовое уведомление о новом пользователе отправлено.",
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

        // ============================
        // SETTINGS SAVE FILE
        // ============================

        private static void SaveSettings(string path, ServerSettings s)
        {
            var lines = new[]
            {
                // ===== MAIL =====
                $"MailHost={s.MailHost}",
                $"MailPort={s.MailPort}",
                $"MailUseSsl={s.MailUseSsl}",
                $"MailLogin={s.MailLogin}",
                $"MailPassword={CryptoHelper.Protect(s.MailPassword)}",
                $"ImapNewCertificatesFolder={s.ImapNewCertificatesFolder}",
                $"ImapRevocationsFolder={s.ImapRevocationsFolder}",

                // ===== FIREBIRD =====
                $"FirebirdDbPath={s.FirebirdDbPath}",
                $"FbServer={s.FbServer}",
                $"FbUser={s.FbUser}",
                $"FbPassword={s.FbPassword}",
                $"FbDialect={s.FbDialect}",
                $"FbCharset={s.FbCharset}",

                // ===== SERVER =====
                $"CheckIntervalMinutes={s.CheckIntervalMinutes}",
                $"NotifyDaysThreshold={s.NotifyDaysThreshold}",
                $"NotifyOnlyInWorkHours={s.NotifyOnlyInWorkHours}",
                $"MinimizeToTrayOnClose={s.MinimizeToTrayOnClose}",

                // ===== BIMOID =====
                $"BimoidAccountsKrasnoflotskaya={s.BimoidAccountsKrasnoflotskaya?.Replace(Environment.NewLine, "\\n")}",
                $"BimoidAccountsPionerskaya={s.BimoidAccountsPionerskaya?.Replace(Environment.NewLine, "\\n")}",
            };

            File.WriteAllLines(path, lines, Encoding.UTF8);
        }

        // ============================
        // INotifyPropertyChanged
        // ============================

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));


    }
}
