using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using ImapCertWatcher;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ImapCertWatcher.Server
{
    public partial class ServerWindow : Window
    {
        private readonly Server.ServerHost _server;
        private CancellationTokenSource _checkCts;
        private bool _forceExit;


        private System.Windows.Forms.NotifyIcon _trayIcon;
        public ServerWindow()
        {
            InitializeComponent();

            ServerSettings settings;

            try
            {
                var settingsPath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "server.settings.txt");

                settings = SettingsLoader.LoadServer(settingsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка загрузки server.settings.txt:\n\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                Close();
                return;
            }

            var db = new DbHelper(settings, AppendLog);

            if (!db.TryOpenConnection(out var err))
            {
                var res = MessageBox.Show(
                    "Не удалось подключиться к базе данных:\n\n" +
                    err.Message + "\n\nСоздать базу?",
                    "Ошибка подключения",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (res == MessageBoxResult.Yes)
                {
                    try
                    {
                        db.CreateDatabaseExplicit();
                        db.EnsureDatabaseAndTable();
                        MessageBox.Show("База данных создана.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            ex.Message,
                            "Ошибка создания БД",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        Close();
                        return;
                    }
                }
                else
                {
                    Close();
                    return;
                }
            }
            else
            {
                try
                {
                    db.EnsureDatabaseAndTable();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Ошибка подготовки базы данных:\n\n" + ex.Message,
                        "Ошибка БД",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    Close();
                    return;
                }
            }

            try
            {
                _server = new Server.ServerHost(
                    settings,
                    db,
                    AppendLog);

                _server.Start();
                InitTray();

                SetStatus("Сервер работает");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка запуска сервера:\n\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Close();
                return;
            }
        }

        private void InitTray()
        {
            _trayIcon = new System.Windows.Forms.NotifyIcon();

            _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(
                Assembly.GetExecutingAssembly().Location);

            _trayIcon.Text = "ImapCertWatcher Server";
            _trayIcon.Visible = true;

            _trayIcon.DoubleClick += (s, e) => RestoreFromTray();

            var menu = new System.Windows.Forms.ContextMenuStrip();

            menu.Items.Add("Открыть", null, (_, __) =>
            {
                RestoreFromTray();
            });

            menu.Items.Add("Перезапустить сервер", null, (_, __) =>
            {
                Restart_Click(null, null);
            });

            menu.Items.Add("Выход", null, (_, __) =>
            {
                ExitApplication();
            });

            _trayIcon.ContextMenuStrip = menu;
        }

        private void RestoreFromTray()
        {
            Show();
            if (WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;

            Activate();
        }

        private void ExitApplication()
        {
            _forceExit = true;

            try
            {
                _trayIcon.Visible = false;
            }
            catch { }

            Application.Current.Shutdown();
        }

        public void SetStatus(string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => SetStatus(text)));
                return;
            }

            StatusText.Text = text ?? "";
        }

        private void SetCheckingState(bool checking)
        {

            CheckNewCertsButton.IsEnabled = !checking;
            CheckNewRevokesButton.IsEnabled = !checking;
            RestartButton.IsEnabled = !checking;
            SettingsButton.IsEnabled = !checking;

            CheckNewCertsButton.Visibility = checking ? Visibility.Collapsed : Visibility.Visible;
            CheckNewRevokesButton.Visibility = checking ? Visibility.Collapsed : Visibility.Visible;
            CancelButton.Visibility = checking ? Visibility.Visible : Visibility.Collapsed;

            Progress.Visibility = checking ? Visibility.Visible : Visibility.Collapsed;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            AppendLog("Отмена операции...");

            _checkCts?.Cancel();
        }

        public async Task<string> RunFullCertificatesCheckFromSettingsAsync()
        {
            if (_checkCts != null)
                return "BUSY RUNNING";

            try
            {
                _checkCts = new CancellationTokenSource();
                SetCheckingState(true);
                AppendLog("ПОЛНАЯ перепроверка сертификатов");
                SetStatus("Идёт полная проверка папки сертификатов...");

                var result = await _server.RequestFullCertificatesCheckAsync(_checkCts.Token);

                AppendLog("FULL CERTS CHECK RESULT: " + result);

                if (result != null && result.StartsWith("BUSY"))
                    SetStatus("Сервер занят");
                else if (result != null && result.StartsWith("ERROR"))
                    SetStatus("Ошибка проверки");
                else
                    SetStatus("Сервер работает");

                return result;
            }
            catch (OperationCanceledException)
            {
                AppendLog("Полная перепроверка сертификатов отменена");
                SetStatus("Отменено");
                throw;
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка полной перепроверки сертификатов: " + ex.Message);
                SetStatus("Ошибка");
                return "ERROR FULL CERTS CHECK";
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
            }
        }

        public async Task<string> RunFullRevocationsCheckFromSettingsAsync()
        {
            if (_checkCts != null)
                return "BUSY RUNNING";

            try
            {
                _checkCts = new CancellationTokenSource();
                SetCheckingState(true);
                AppendLog("ПОЛНАЯ перепроверка аннулирований");
                SetStatus("Идёт полная проверка папки аннулирований...");

                var result = await _server.RequestFullRevocationsCheckAsync(_checkCts.Token);

                AppendLog("FULL REVOKES CHECK RESULT: " + result);

                if (result != null && result.StartsWith("BUSY"))
                    SetStatus("Сервер занят");
                else if (result != null && result.StartsWith("ERROR"))
                    SetStatus("Ошибка проверки");
                else
                    SetStatus("Сервер работает");

                return result;
            }
            catch (OperationCanceledException)
            {
                AppendLog("Полная перепроверка аннулирований отменена");
                SetStatus("Отменено");
                throw;
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка полной перепроверки аннулирований: " + ex.Message);
                SetStatus("Ошибка");
                return "ERROR FULL REVOKES CHECK";
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
            }
        }



        private const int MaxUiLogLines = 1000;

        public void AppendLog(string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => AppendLog(text)));
                return;
            }

            string safeText = text ?? "";
            string fileLine = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [ServerWindow] {safeText}";

            try
            {
                System.IO.File.AppendAllText(
                    ImapCertWatcher.Utils.LogSession.SessionLogFile,
                    fileLine + Environment.NewLine,
                    System.Text.Encoding.UTF8);
            }
            catch
            {
                // Лог в файл не должен валить UI
            }

            System.Windows.Documents.Paragraph paragraph;

            if (LogBox.Document.Blocks.FirstBlock is System.Windows.Documents.Paragraph p)
            {
                paragraph = p;
            }
            else
            {
                LogBox.Document.Blocks.Clear();
                paragraph = new System.Windows.Documents.Paragraph();
                LogBox.Document.Blocks.Add(paragraph);
            }

            var run = new System.Windows.Documents.Run($"{DateTime.Now:HH:mm:ss} {safeText}{Environment.NewLine}");

            if (safeText.IndexOf("ошибка", StringComparison.OrdinalIgnoreCase) >= 0 ||
                safeText.IndexOf("критич", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                run.Foreground = System.Windows.Media.Brushes.Red;
            }
            else if (safeText.IndexOf("добавлено", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     safeText.IndexOf("обновл", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                run.Foreground = System.Windows.Media.Brushes.Green;
            }
            else if (safeText.IndexOf("пропущено", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                run.Foreground = System.Windows.Media.Brushes.Gray;
            }
            else if (safeText.StartsWith("ИТОГО", StringComparison.OrdinalIgnoreCase))
            {
                run.Foreground = System.Windows.Media.Brushes.SteelBlue;
            }

            paragraph.Inlines.Add(run);

            while (paragraph.Inlines.Count > MaxUiLogLines)
            {
                paragraph.Inlines.Remove(paragraph.Inlines.FirstInline);
            }

            LogBox.ScrollToEnd();
        }

        private async void CheckNewCerts_Click(object sender, RoutedEventArgs e)
        {
            if (_checkCts != null)
                return;

            try
            {
                _checkCts = new CancellationTokenSource();

                SetCheckingState(true);

                AppendLog("Ручная проверка новых сертификатов");
                SetStatus("Идёт проверка новых сертификатов...");

                var result = await _server.RequestFastCertificatesCheckAsync(_checkCts.Token);

                AppendLog("FAST CERTS CHECK RESULT: " + result);

                if (result != null && result.StartsWith("BUSY"))
                    SetStatus("Сервер занят");
                else if (result != null && result.StartsWith("ERROR"))
                    SetStatus("Ошибка проверки");
                else
                    SetStatus("Сервер работает");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Проверка новых сертификатов отменена");
                SetStatus("Отменено");
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка проверки новых сертификатов: " + ex.Message);
                SetStatus("Ошибка");
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
            }
        }

        private async void CheckNewRevokes_Click(object sender, RoutedEventArgs e)
        {
            if (_checkCts != null)
                return;

            try
            {
                _checkCts = new CancellationTokenSource();

                SetCheckingState(true);

                AppendLog("Ручная проверка новых аннулирований");
                SetStatus("Идёт проверка новых аннулирований...");

                var result = await _server.RequestFastRevocationsCheckAsync(_checkCts.Token);

                AppendLog("FAST REVOKES CHECK RESULT: " + result);

                if (result != null && result.StartsWith("BUSY"))
                    SetStatus("Сервер занят");
                else if (result != null && result.StartsWith("ERROR"))
                    SetStatus("Ошибка проверки");
                else
                    SetStatus("Сервер работает");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Проверка новых аннулирований отменена");
                SetStatus("Отменено");
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка проверки новых аннулирований: " + ex.Message);
                SetStatus("Ошибка");
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
            }
        }

        private void CopySettings(ServerSettings target, ServerSettings source)
        {
            target.MailHost = source.MailHost;
            target.MailPort = source.MailPort;
            target.MailUseSsl = source.MailUseSsl;
            target.MailLogin = source.MailLogin;
            target.MailPassword = source.MailPassword;

            target.ImapNewCertificatesFolder = source.ImapNewCertificatesFolder;
            target.ImapRevocationsFolder = source.ImapRevocationsFolder;

            target.FirebirdDbPath = source.FirebirdDbPath;
            target.FbServer = source.FbServer;
            target.FbUser = source.FbUser;
            target.FbPassword = source.FbPassword;
            target.FbDialect = source.FbDialect;
            target.FbCharset = source.FbCharset;
            target.IsDevelopment = source.IsDevelopment;

            target.ServerPort = source.ServerPort;
            target.CheckIntervalMinutes = source.CheckIntervalMinutes;
            target.NotifyDaysThreshold = source.NotifyDaysThreshold;
            target.NotifyOnlyInWorkHours = source.NotifyOnlyInWorkHours;
            target.AutoStartServer = source.AutoStartServer;
            target.MinimizeToTrayOnClose = source.MinimizeToTrayOnClose;

            target.BimoidAccountsKrasnoflotskaya = source.BimoidAccountsKrasnoflotskaya;
            target.BimoidAccountsPionerskaya = source.BimoidAccountsPionerskaya;
            target.BimoidSenderExePath = source.BimoidSenderExePath;
            target.BimoidJobDirectory = source.BimoidJobDirectory;
            target.BimoidServer = source.BimoidServer;
            target.BimoidPort = source.BimoidPort;
            target.BimoidLogin = source.BimoidLogin;
            target.BimoidPassword = source.BimoidPassword;
            target.BimoidDelayBetweenMessagesMs = source.BimoidDelayBetweenMessagesMs;
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            var oldSettings = _server.Settings.Clone();
            var editedSettings = _server.Settings.Clone();

            var wnd = new ServerSettingsWindow(editedSettings)
            {
                Owner = this
            };

            if (wnd.ShowDialog() == true)
            {
                bool needRestart = NeedRestart(oldSettings, editedSettings);

                CopySettings(_server.Settings, editedSettings);

                if (needRestart)
                {
                    MessageBox.Show(
                        "Некоторые изменения вступят в силу после перезапуска сервера.",
                        "Перезапуск требуется",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    AppendLog("Настройки применены без перезапуска.");
                }
            }
        }

        private bool NeedRestart(ServerSettings oldS, ServerSettings newS)
        {
            // БД
            if (oldS.FirebirdDbPath != newS.FirebirdDbPath) return true;
            if (oldS.FbServer != newS.FbServer) return true;
            if (oldS.FbUser != newS.FbUser) return true;
            if (oldS.FbPassword != newS.FbPassword) return true;
            if (oldS.FbDialect != newS.FbDialect) return true;
            if (oldS.FbCharset != newS.FbCharset) return true;

            // Почта / IMAP
            if (oldS.MailHost != newS.MailHost) return true;
            if (oldS.MailPort != newS.MailPort) return true;
            if (oldS.MailUseSsl != newS.MailUseSsl) return true;
            if (oldS.MailLogin != newS.MailLogin) return true;
            if (oldS.MailPassword != newS.MailPassword) return true;
            if (oldS.ImapNewCertificatesFolder != newS.ImapNewCertificatesFolder) return true;
            if (oldS.ImapRevocationsFolder != newS.ImapRevocationsFolder) return true;

            // TCP / таймер
            if (oldS.ServerPort != newS.ServerPort) return true;
            if (oldS.CheckIntervalMinutes != newS.CheckIntervalMinutes) return true;

            return false;
        }

        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AppendLog("Перезапуск сервера...");

                string exePath = Assembly.GetExecutingAssembly().Location;

                Process.Start(new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true
                });

                ExitApplication();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка перезапуска:\n\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!_forceExit && _server?.Settings?.MinimizeToTrayOnClose == true)
            {
                e.Cancel = true;
                Hide();
                AppendLog("Сервер свернут в трей");
                return;
            }

            try
            {
                _server?.Dispose();
            }
            catch { }

            try
            {
                if (_trayIcon != null)
                    _trayIcon.Visible = false;
            }
            catch { }

            try
            {
                _trayIcon?.Dispose();
            }
            catch { }

            base.OnClosing(e);
        }
    }
}
