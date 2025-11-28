using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class App : Application
    {
        private SplashScreen _splashScreen;

        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            // 1) Сначала применяем тему
            ApplyThemeFromSettings();

            // 2) Показываем splash
            _splashScreen = new SplashScreen();
            _splashScreen.Show();

            // Замер времени показа splash
            var sw = Stopwatch.StartNew();

            // 3) Создаем главное окно
            var mainWindow = new MainWindow();

            // Прогресс на splash
            mainWindow.ProgressUpdated += (message, progress) =>
            {
                _splashScreen?.UpdateStatus(message);
                _splashScreen?.UpdateProgress(progress);
            };

            // Когда данные загружены
            mainWindow.DataLoaded += async () =>
            {
                sw.Stop();
                const int minShowMs = 2000; // минимум 2 секунды

                int remain = minShowMs - (int)sw.ElapsedMilliseconds;
                if (remain > 0)
                {
                    // ждём, чтобы суммарное время показа было не меньше 2 сек
                    await Task.Delay(remain);
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _splashScreen?.CloseSplash();  // там теперь будет плавное исчезновение
                    _splashScreen = null;
                }), DispatcherPriority.Background);
            };

            // 4) Показываем главное окно
            mainWindow.Show();

            // 5) Защита от зависшего splash (если вдруг DataLoaded не вызовется)
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
        /// </summary>
        private void ApplyThemeFromSettings()
        {
            bool useDarkTheme = true; // как у тебя и было

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
