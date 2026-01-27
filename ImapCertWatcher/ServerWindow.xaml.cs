using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
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
        public ServerWindow()

        {
            InitializeComponent();

            AppSettings settings;
            try
            {
                settings = SettingsLoader.Load("settings.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка загрузки settings.txt:\n\n" + ex.Message,
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

            _server.Start();

            SetStatus("Сервер работает");
            AppendLog("ServerHost запущен");
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
            var wnd = new ServerSettingsWindow(_server.Settings)
            {
                Owner = this
            };
            wnd.ShowDialog();
            AppendLog("Настройки изменены. Перезапустите сервер.");
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
    }
}
