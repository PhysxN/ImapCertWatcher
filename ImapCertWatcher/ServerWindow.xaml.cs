using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace ImapCertWatcher
{
    public partial class ServerWindow : Window
    {
        private readonly Server.ServerHost _server;
        private bool _isChecking;
        public ServerWindow(AppSettings settings)
        {
            InitializeComponent();
                        
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
                _isChecking = true;
                CheckNowButton.IsEnabled = false;

                AppendLog("Ручная FAST-проверка");
                SetStatus("Идёт проверка почты...");

                await _server.RequestFastCheckAsync();

                SetStatus("Сервер работает");
            }
            finally
            {
                _isChecking = false;
                CheckNowButton.IsEnabled = true;
            }
        }

        private async void CheckAll_Click(object sender, RoutedEventArgs e)
        {
            if (_isChecking)
                return;

            try
            {
                _isChecking = true;
                CheckNowButton.IsEnabled = false;
                CheckAllButton.IsEnabled = false;

                AppendLog("ПОЛНАЯ проверка всех писем");
                SetStatus("Идёт полная проверка почты...");

                await _server.RequestFullCheckAsync();

                SetStatus("Сервер работает");
            }
            finally
            {
                _isChecking = false;
                CheckNowButton.IsEnabled = true;
                CheckAllButton.IsEnabled = true;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
