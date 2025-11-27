using System;
using System.Windows;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class App : Application
    {
        private SplashScreen _splashScreen;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Показываем splash screen
            _splashScreen = new SplashScreen();
            _splashScreen.Show();

            // Создаем главное окно
            var mainWindow = new MainWindow();

            // Подписываемся на события прогресса
            mainWindow.ProgressUpdated += (message, progress) =>
            {
                _splashScreen?.UpdateStatus(message);
                _splashScreen?.UpdateProgress(progress);
            };

            // ★ ПОДПИСЫВАЕМСЯ НА СОБЫТИЕ ПОЛНОЙ ЗАГРУЗКИ ДАННЫХ ★
            mainWindow.DataLoaded += () =>
            {
                // Закрываем splash screen только когда данные полностью загружены
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _splashScreen?.CloseSplash();
                    _splashScreen = null; // Освобождаем ссылку
                }), DispatcherPriority.Background);
            };

            // Показываем главное окно
            mainWindow.Show();

            // ★ ДОБАВЛЯЕМ ЗАЩИТУ ОТ ВИСЯЩЕГО SPLASH SCREEN ★
            var safetyTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(15) // Уменьшили до 15 секунд
            };
            safetyTimer.Tick += (s, args) =>
            {
                safetyTimer.Stop();
                if (_splashScreen != null && _splashScreen.IsVisible)
                {
                    System.Diagnostics.Debug.WriteLine("Safety timer: принудительное закрытие SplashScreen");
                    _splashScreen.CloseSplash();
                    _splashScreen = null;
                }
            };
            safetyTimer.Start();
        }
    }
}