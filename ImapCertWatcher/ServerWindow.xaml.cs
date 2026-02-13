using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace ImapCertWatcher
{
    public partial class ServerWindow : Window
    {
        private readonly Server.ServerHost _server;
        private bool _isChecking;
        private CancellationTokenSource _checkCts;
        private const string AUTOSTART_REG_PATH =
            @"Software\Microsoft\Windows\CurrentVersion\Run";

        private const string AUTOSTART_VALUE_NAME =
            "ImapCertWatcherServer";
        private System.Windows.Forms.NotifyIcon _trayIcon;
        public ServerWindow()
        {
            InitializeComponent();
            InitTray();

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
                db.EnsureDatabaseAndTable();
            }

            _server = new Server.ServerHost(
                settings,
                db,
                AppendLog);

            ApplyServerAutoStart();   // 🔥 ТОЛЬКО ПОСЛЕ создания _server

            _server.Start();

            SetStatus("Сервер работает");
            AppendLog("ServerHost запущен");
        }

        private void InitTray()
        {
            _trayIcon = new System.Windows.Forms.NotifyIcon();

            _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(
                        Assembly.GetExecutingAssembly().Location);
            _trayIcon.Text = "ImapCertWatcher Server";
            _trayIcon.Visible = true;

            _trayIcon.DoubleClick += (s, e) =>
            {
                Show();
                WindowState = WindowState.Normal;
                Activate();
            };

            var menu = new System.Windows.Forms.ContextMenuStrip();

            menu.Items.Add("Открыть", null, (_, __) =>
            {
                Show();
                WindowState = WindowState.Normal;
                Activate();
            });

            menu.Items.Add("Перезапустить сервер", null, (_, __) =>
            {
                Restart_Click(null, null);
            });

            menu.Items.Add("Выход", null, (_, __) =>
            {
                _trayIcon.Visible = false;
                Application.Current.Shutdown();
            });

            _trayIcon.ContextMenuStrip = menu;
        }
        public void SetStatus(string text)
        {
            StatusText.Text = text;
        }

        private void SetCheckingState(bool checking)
        {
            _isChecking = checking;

            CheckNowButton.IsEnabled = !checking;
            CheckAllButton.IsEnabled = !checking;
            RestartButton.IsEnabled = !checking;
            SettingsButton.IsEnabled = !checking;

            CheckNowButton.Visibility = checking ? Visibility.Collapsed : Visibility.Visible;
            CancelButton.Visibility = checking ? Visibility.Visible : Visibility.Collapsed;

            Progress.Visibility = checking ? Visibility.Visible : Visibility.Collapsed;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            AppendLog("Отмена операции...");

            _checkCts?.Cancel();
        }


        private void ApplyServerAutoStart()
        {
            try
            {
                using (var key =
                    Registry.CurrentUser.OpenSubKey(AUTOSTART_REG_PATH, true))
                {
                    if (key == null) return;

                    string exePath = Assembly.GetExecutingAssembly().Location;
                    exePath = $"\"{exePath}\"";

                    if (_server.Settings.AutoStartServer)
                    {
                        key.SetValue(AUTOSTART_VALUE_NAME, exePath);
                        AppendLog("Автозапуск сервера включен");
                    }
                    else
                    {
                        key.DeleteValue(AUTOSTART_VALUE_NAME, false);
                        AppendLog("Автозапуск сервера выключен");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка автозапуска: " + ex.Message);
            }
        }
        public void AppendLog(string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendLog(text));
                return;
            }

            // создаём Paragraph один раз
            if (LogBox.Document.Blocks.Count == 0)
                LogBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph());

            var paragraph = (System.Windows.Documents.Paragraph)LogBox.Document.Blocks.FirstBlock;

            var run = new System.Windows.Documents.Run(
                $"{DateTime.Now:HH:mm:ss} {text}\n");

            if (text.Contains("ОШИБКА") || text.Contains("КРИТИЧЕСК"))
                run.Foreground = System.Windows.Media.Brushes.Red;
            else if (text.Contains("ДОБАВЛЕНО") || text.Contains("обновл"))
                run.Foreground = System.Windows.Media.Brushes.Green;
            else if (text.Contains("ПРОПУЩЕНО"))
                run.Foreground = System.Windows.Media.Brushes.Gray;
            else if (text.StartsWith("ИТОГО"))
                run.Foreground = System.Windows.Media.Brushes.SteelBlue;

            paragraph.Inlines.Add(run);

            LogBox.ScrollToEnd();
        }

        private async void CheckNow_Click(object sender, RoutedEventArgs e)
        {
            if (_isChecking)
                return;

            try
            {
                _checkCts = new CancellationTokenSource();

                SetCheckingState(true);

                AppendLog("Ручная FAST-проверка");
                SetStatus("Идёт проверка почты...");

                await _server.RequestFastCheckAsync(_checkCts.Token);

                SetStatus("Сервер работает");
            }
            catch (OperationCanceledException)
            {
                AppendLog("Проверка отменена");
                SetStatus("Отменено");
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
            }
        }

        private async void CheckAll_Click(object sender, RoutedEventArgs e)
        {
            if (_isChecking)
                return;

            try
            {
                _checkCts = new CancellationTokenSource();
                SetCheckingState(true);
                CheckAllButton.IsEnabled = false;
                StartProgress();
                AppendLog("ПОЛНАЯ проверка всех писем");
                SetStatus("Идёт полная проверка почты...");

                await _server.RequestFullCheckAsync();

                SetStatus("Сервер работает");
            }
            finally
            {
                SetCheckingState(false);
                _checkCts = null;
                CheckAllButton.IsEnabled = true;
            }
        }

        private void StartProgress()
        {
            Progress.Visibility = Visibility.Visible;
        }

        private void StopProgress()
        {
            Progress.Visibility = Visibility.Collapsed;
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            // Копируем текущие настройки
            var oldSettings = _server.Settings.Clone();

            var wnd = new ServerSettingsWindow(_server.Settings)
            {
                Owner = this
            };

            if (wnd.ShowDialog() == true)
            {
                bool needRestart = NeedRestart(oldSettings, _server.Settings);

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
            // 🔴 БД — всегда требует перезапуска
            if (oldS.FirebirdDbPath != newS.FirebirdDbPath) return true;
            if (oldS.FbServer != newS.FbServer) return true;
            if (oldS.FbUser != newS.FbUser) return true;
            if (oldS.FbPassword != newS.FbPassword) return true;

            // 🟡 Почта — лучше перезапуск
            if (oldS.MailHost != newS.MailHost) return true;
            if (oldS.MailPort != newS.MailPort) return true;
            if (oldS.MailUseSsl != newS.MailUseSsl) return true;
            if (oldS.MailLogin != newS.MailLogin) return true;
            if (oldS.MailPassword != newS.MailPassword) return true;

            return false;
        }

        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AppendLog("Перезапуск сервера...");

                string exePath = Assembly.GetExecutingAssembly().Location;

                // Получаем аргументы текущего запуска
                string args = string.Join(" ", Environment.GetCommandLineArgs().Skip(1));

                Process.Start(new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = args,
                    UseShellExecute = true
                });

                Application.Current.Shutdown();
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
            if (_server.Settings.MinimizeToTrayOnClose)
            {
                e.Cancel = true;
                Hide();
                AppendLog("Сервер свернут в трей");
                return;
            }

            _trayIcon.Visible = false;

            base.OnClosing(e);
        }
    }
}
