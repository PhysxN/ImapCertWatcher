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
            // 1) СНАЧАЛА применяем тему (чтобы Splash сразу был в нужных цветах)
            ApplyThemeFromSettings();

            // 2) Показываем splash screen
            _splashScreen = new SplashScreen();
            _splashScreen.Show();

            // 3) Создаем главное окно
            var mainWindow = new MainWindow();

            // Подписываемся на события прогресса
            mainWindow.ProgressUpdated += (message, progress) =>
            {
                _splashScreen?.UpdateStatus(message);
                _splashScreen?.UpdateProgress(progress);
            };

            // Когда данные загружены — закрываем сплэш
            mainWindow.DataLoaded += () =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _splashScreen?.CloseSplash();
                    _splashScreen = null;
                }), DispatcherPriority.Background);
            };

            // Показываем главное окно
            mainWindow.Show();

            // Защита от зависшего сплэша
            var safetyTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(15)
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

        /// <summary>
        /// Применяет тему к глобальным ресурсам приложения.
        /// Сейчас просто включаем тёмную; при желании сюда можно
        /// подставить чтение из AppSettings.
        /// </summary>
        private void ApplyThemeFromSettings()
        {
            // TODO: здесь можно прочитать реальную настройку из файла/БД.
            // Пока просто считаем, что всегда нужна тёмная тема:
            bool useDarkTheme = true;

            var res = Current.Resources;

            if (useDarkTheme)
            {
                res["WindowBackgroundBrush"] = res["DarkWindowBackground"];
                res["ControlBackgroundBrush"] = res["DarkControlBackground"];
                res["TextColorBrush"] = res["DarkTextColor"];
                res["BorderColorBrush"] = res["DarkBorderColor"];
                res["AccentColorBrush"] = res["DarkAccentColor"];
                res["HoverColorBrush"] = res["DarkHoverColor"];
            }
            else
            {
                res["WindowBackgroundBrush"] = res["LightWindowBackground"];
                res["ControlBackgroundBrush"] = res["LightControlBackground"];
                res["TextColorBrush"] = res["LightTextColor"];
                res["BorderColorBrush"] = res["LightBorderColor"];
                res["AccentColorBrush"] = res["LightAccentColor"];
                res["HoverColorBrush"] = res["LightHoverColor"];
            }
        }
    }
}
