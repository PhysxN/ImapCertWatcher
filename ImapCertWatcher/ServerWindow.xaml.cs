using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.Windows;

namespace ImapCertWatcher
{
    public partial class ServerWindow : Window
    {
        private readonly Server.ServerHost _server;

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

            LogBox.AppendText($"{DateTime.Now:HH:mm:ss} {text}{Environment.NewLine}");
            LogBox.ScrollToEnd();
        }

        private async void CheckNow_Click(object sender, RoutedEventArgs e)
        {
            AppendLog("Ручная FAST-проверка");
            await _server.RequestFastCheckAsync();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
