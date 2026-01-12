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

        public ServerWindow()
        {
            InitializeComponent();

            var settings = new AppSettings();
            var db = new DbHelper(settings, AppendLog);

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
