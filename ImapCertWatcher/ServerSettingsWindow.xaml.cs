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
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ImapCertWatcher
{
    public partial class ServerSettingsWindow : Window
    {
        private readonly AppSettings _settings;
        private readonly NotificationManager _notificationManager;
        public ObservableCollection<string> ImapFolders { get; } =
    new ObservableCollection<string>();

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ServerSettingsWindow(AppSettings settings)
        {
            InitializeComponent();

            _settings = settings;
            DataContext = _settings;
            _notificationManager = new NotificationManager(_settings, s => { });

            // пароли вручную
            MailPasswordBox.Password = _settings.MailPassword;
            DbPasswordBox.Password = _settings.FbPassword;

            Loaded += ServerSettingsWindow_Loaded;
        }

        private void ServerSettingsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Восстанавливаем значения из settings.txt
            cmbNewCertsFolder.GetBindingExpression(ComboBox.TextProperty)?.UpdateTarget();
            cmbRevocationsFolder.GetBindingExpression(ComboBox.TextProperty)?.UpdateTarget();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _settings.MailPassword = MailPasswordBox.Password;
            _settings.FbPassword = DbPasswordBox.Password;

            SaveSettings("settings.txt", _settings);

            MessageBox.Show(
                "Настройки сохранены.\nПерезапустите сервер.",
                "Сервер",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            Close();
        }

        private async void RefreshFolders_Click(object sender, RoutedEventArgs e)
        {
            ImapFolders.Clear();

            try
            {
                using (var client = new ImapClient())
                {
                    client.ServerCertificateValidationCallback =
                        (clientSender, certificate, chain, errors) => true;

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

                    await LoadFoldersRecursive(root);

                    Dispatcher.Invoke(() =>
                    {
                        if (ImapFolders.Contains(_settings.ImapNewCertificatesFolder))
                            _settings.ImapNewCertificatesFolder = _settings.ImapNewCertificatesFolder;

                        if (ImapFolders.Contains(_settings.ImapRevocationsFolder))
                            _settings.ImapRevocationsFolder = _settings.ImapRevocationsFolder;
                    });

                    await client.DisconnectAsync(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Ошибка IMAP",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void TestMail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = System.Windows.Input.Cursors.Wait;

                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    client.ServerCertificateValidationCallback =
                        (s, c, h, er) => true;

                    await client.ConnectAsync(
                        _settings.MailHost,
                        _settings.MailPort,
                        _settings.MailUseSsl
                            ? MailKit.Security.SecureSocketOptions.SslOnConnect
                            : MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);

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
                    "Ошибка подключения к почте:\n\n" + ex.Message,
                    "IMAP ошибка",
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

                var testSettings = new AppSettings
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

        private async Task LoadFoldersRecursive(IMailFolder folder)
        {
            ImapFolders.Add(folder.FullName);

            var subs = await folder.GetSubfoldersAsync(false);
            foreach (var sub in subs)
            {
                await LoadFoldersRecursive(sub);
            }
        }

        private static void SaveSettings(string path, AppSettings s)
        {
            var lines = new[]
            {
        // ===== MAIL =====
        $"MailHost={s.MailHost}",
        $"MailPort={s.MailPort}",
        $"MailUseSsl={s.MailUseSsl}",
        $"MailLogin={s.MailLogin}",
        $"MailPassword={s.MailPassword}",
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

        // ===== BIMOID =====
        $"BimoidAccountsKrasnoflotskaya={s.BimoidAccountsKrasnoflotskaya?.Replace(Environment.NewLine, "\\n")}",
        $"BimoidAccountsPionerskaya={s.BimoidAccountsPionerskaya?.Replace(Environment.NewLine, "\\n")}",
    };

            File.WriteAllLines(path, lines, Encoding.UTF8);
        }
    }
}
