using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Media;

namespace ImapCertWatcher
{
    public partial class App : Application
    {
        private SplashScreen _splashScreen;

        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            // 1) Сначала применяем тему (чтобы Splash сразу был в нужных цветах)
            ApplyThemeFromSettings();

            // 2) Показываем splash
            _splashScreen = new SplashScreen();
            try
            {
                _splashScreen.Show();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Не удалось показать SplashScreen: " + ex);
            }

            // Замер времени показа splash
            var sw = Stopwatch.StartNew();

            // 3) Создаем главное окно
            var mainWindow = new MainWindow();

            // Перенаправляем прогресс в splash
            mainWindow.ProgressUpdated += (message, progress) =>
            {
                try
                {
                    _splashScreen?.UpdateStatus(message);
                    _splashScreen?.UpdateProgress(progress);
                }
                catch { /* игнорируем ошибки в обновлении сплэша */ }
            };

            // Когда данные загружены — подождём минимум и затем закроем splash
            mainWindow.DataLoaded += async () =>
            {
                try
                {
                    sw.Stop();
                    const int minShowMs = 2000; // минимум 2 секунды показа

                    int remain = minShowMs - (int)sw.ElapsedMilliseconds;
                    if (remain > 0)
                    {
                        await Task.Delay(remain);
                    }

                    // Закрываем splash и ждём завершения анимации.
                    // CloseSplashAsync должен выполняться в UI-потоке, поэтому используем Dispatcher.InvokeAsync
                    if (_splashScreen != null)
                    {
                        var op = Dispatcher.InvokeAsync(() => _splashScreen.CloseSplashAsync(), DispatcherPriority.Background);
                        // op.Result is Task (CloseSplashAsync), unwrap:
                        await (await op.Task).ConfigureAwait(false);
                        _splashScreen = null;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Ошибка в DataLoaded handler (splash): " + ex);
                    try
                    {
                        if (_splashScreen != null)
                        {
                            var op = Dispatcher.InvokeAsync(() => _splashScreen.CloseSplashAsync(), DispatcherPriority.Background);
                            await (await op.Task).ConfigureAwait(false);
                            _splashScreen = null;
                        }
                    }
                    catch { /* silent */ }
                }
            };

            // 4) Показываем главное окно
            try
            {
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Не удалось показать MainWindow: " + ex);
            }

            // 5) Защита от зависшего splash (если вдруг DataLoaded не вызовется)
            var safetyTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(15)
            };
            safetyTimer.Tick += async (s, args) =>
            {
                safetyTimer.Stop();
                try
                {
                    if (_splashScreen != null && _splashScreen.IsVisible)
                    {
                        Debug.WriteLine("Safety timer: принудительное закрытие SplashScreen");
                        try
                        {
                            var op = Dispatcher.InvokeAsync(() => _splashScreen.CloseSplashAsync(), DispatcherPriority.Background);
                            await (await op.Task).ConfigureAwait(false);
                        }
                        catch (Exception exClose)
                        {
                            Debug.WriteLine("Ошибка при safety CloseSplashAsync(): " + exClose);
                            try { _splashScreen?.Close(); } catch { }
                        }
                        finally
                        {
                            _splashScreen = null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Ошибка в safetyTimer.Tick: " + ex);
                }
            };
            safetyTimer.Start();
        }

        /// <summary>
        /// Применяет тему к глобальным ресурсам приложения.
        /// </summary>
        private void ApplyThemeFromSettings()
        {
            // Читаем theme.txt (если есть) — fallback: light
            bool useDarkTheme = false;
            try
            {
                var themeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "theme.txt");
                if (File.Exists(themeFile))
                {
                    var theme = File.ReadAllText(themeFile).Trim();
                    useDarkTheme = string.Equals(theme, "dark", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { /* ignore */ }

            var res = Application.Current.Resources;

            if (useDarkTheme)
            {
                res["WindowBackgroundBrush"] = res["DarkWindowBackground"];
                res["ControlBackgroundBrush"] = res["DarkControlBackground"];
                res["TextColorBrush"] = res["DarkTextColor"];
                res["BorderColorBrush"] = res["DarkBorderColor"];
                res["AccentColorBrush"] = res["DarkAccentColor"];
                res["HoverColorBrush"] = res["DarkHoverColor"];

                // Привязать системные кисти для всплывающих окон/выделения
                try
                {
                    res[SystemColors.WindowBrushKey] = res["DarkWindowBackground"];
                    res[SystemColors.ControlTextBrushKey] = res["DarkTextColor"];
                    res[SystemColors.HighlightBrushKey] = res["DarkAccentColor"];
                    res[SystemColors.HighlightTextBrushKey] = Brushes.White;
                }
                catch { }
            }
            else
            {
                res["WindowBackgroundBrush"] = res["LightWindowBackground"];
                res["ControlBackgroundBrush"] = res["LightControlBackground"];
                res["TextColorBrush"] = res["LightTextColor"];
                res["BorderColorBrush"] = res["LightBorderColor"];
                res["AccentColorBrush"] = res["LightAccentColor"];
                res["HoverColorBrush"] = res["LightHoverColor"];

                try
                {
                    res[SystemColors.WindowBrushKey] = res["LightWindowBackground"];
                    res[SystemColors.ControlTextBrushKey] = res["LightTextColor"];
                    res[SystemColors.HighlightBrushKey] = res["LightAccentColor"];
                    res[SystemColors.HighlightTextBrushKey] = Brushes.White;
                }
                catch { }
            }
        }
    }
}
