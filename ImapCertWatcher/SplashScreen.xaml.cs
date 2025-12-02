using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace ImapCertWatcher
{
    public partial class SplashScreen : Window
    {
        private readonly string _gifPath;
        private MediaElement _gifAnimation;
        private bool _isClosing = false;

        public SplashScreen(string gifPath = null)
        {
            InitializeComponent();

            // применяем тему до инициализации анимации
            ApplyThemeFromFile();

            _gifPath = gifPath ?? FindDefaultGif();
            InitializeAnimation();
        }

        private void ApplyThemeFromFile()
        {
            try
            {
                var themeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "theme.txt");
                bool isDark = false;

                if (File.Exists(themeFile))
                {
                    var theme = File.ReadAllText(themeFile).Trim();
                    isDark = theme.Equals("dark", StringComparison.OrdinalIgnoreCase);
                }

                if (isDark)
                {
                    // тёмная тема – цвета такие же, как в MainWindow
                    rootBorder.Background = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x30));   // DarkWindowBackground
                    rootBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x00, 0x7A, 0xCC));  // DarkAccentColor

                    statusText.Foreground = Brushes.White;
                    progressBar.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0x7A, 0xCC));
                    spinner.Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x7A, 0xCC));
                }
                else
                {
                    // светлая тема – исходные цвета
                    rootBorder.Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0));
                    rootBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD7));

                    statusText.Foreground = Brushes.Black;
                    progressBar.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD7));
                    spinner.Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD7));
                }
            }
            catch
            {
                // На всякий случай ничего не делаем при ошибке – останутся дефолтные цвета
            }
        }

        private string FindDefaultGif()
        {
            // Ищем GIF в папке с программой
            var appDir = AppDomain.CurrentDomain.BaseDirectory;
            var possiblePaths = new[]
            {
                Path.Combine(appDir, "loading.gif"),
                Path.Combine(appDir, "Resources", "loading.gif"),
                Path.Combine(appDir, "Assets", "loading.gif"),
                Path.Combine(appDir, "loadinganimation.gif")
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            return null;
        }

        private void InitializeAnimation()
        {
            if (!string.IsNullOrEmpty(_gifPath) && File.Exists(_gifPath))
            {
                try
                {
                    // Создаем MediaElement для GIF анимации
                    _gifAnimation = new MediaElement
                    {
                        LoadedBehavior = MediaState.Play,
                        UnloadedBehavior = MediaState.Manual,
                        Stretch = Stretch.Uniform,
                        Source = new Uri(_gifPath)
                    };

                    _gifAnimation.MediaEnded += GifAnimation_MediaEnded;

                    // Заменяем XAML анимацию на GIF
                    var animationContainer = (Viewbox)this.FindName("animationContainer");
                    if (animationContainer != null)
                    {
                        animationContainer.Child = _gifAnimation;
                    }
                }
                catch (Exception ex)
                {
                    // В случае ошибки используем XAML анимацию
                    System.Diagnostics.Debug.WriteLine($"Ошибка загрузки GIF: {ex.Message}");
                }
            }
            // Если GIF не найден, используется XAML анимация по умолчанию
        }

        public void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                var statusTextBlock = (TextBlock)this.FindName("statusText");
                if (statusTextBlock != null)
                {
                    statusTextBlock.Text = message;
                }
            });
        }

        public void UpdateProgress(double progress)
        {
            Dispatcher.Invoke(() =>
            {
                var progressBarControl = (ProgressBar)this.FindName("progressBar");
                if (progressBarControl != null)
                {
                    if (progress >= 0 && progress <= 100)
                    {
                        progressBarControl.IsIndeterminate = false;
                        progressBarControl.Value = progress;
                    }
                    else
                    {
                        progressBarControl.IsIndeterminate = true;
                    }
                }
            });
        }

        private void GifAnimation_MediaEnded(object sender, RoutedEventArgs e)
        {
            // Зацикливаем анимацию
            if (_gifAnimation != null)
            {
                _gifAnimation.Position = TimeSpan.Zero;
                _gifAnimation.Play();
            }
        }

        /// <summary>
        /// Плавно закрывает SplashScreen и возвращает Task,
        /// завершающийся по окончании анимации закрытия.
        /// Можно вызывать из любого потока — реализация перемещается в UI-поток через Dispatcher.
        /// </summary>
        public Task CloseSplashAsync()
        {
            // Если уже закрываем — вернуть завершённую задачу
            if (_isClosing)
                return Task.CompletedTask;

            var tcs = new TaskCompletionSource<object>();

            // Выполняем всю работу в UI-потоке
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (_isClosing)
                    {
                        tcs.TrySetResult(null);
                        return;
                    }

                    _isClosing = true;

                    // Останавливаем анимацию перед закрытием
                    try
                    {
                        if (_gifAnimation != null)
                        {
                            _gifAnimation.Stop();
                            _gifAnimation.MediaEnded -= GifAnimation_MediaEnded;
                        }
                    }
                    catch { /* ignore */ }

                    // Если окно уже скрыто/закрыто — сразу завершаем
                    if (!this.IsVisible)
                    {
                        try { this.Close(); } catch { }
                        tcs.TrySetResult(null);
                        return;
                    }

                    // Плавное исчезновение окна
                    var anim = new DoubleAnimation
                    {
                        From = this.Opacity,
                        To = 0.0,
                        Duration = TimeSpan.FromMilliseconds(300),
                        FillBehavior = FillBehavior.Stop
                    };

                    anim.Completed += (s, e) =>
                    {
                        try
                        {
                            // Устанавливаем Opacity в 0 и закрываем окно
                            this.Opacity = 0;
                            try { this.Close(); } catch { }
                        }
                        catch { /* ignore */ }
                        finally
                        {
                            tcs.TrySetResult(null);
                        }
                    };

                    // Запускаем анимацию
                    try
                    {
                        this.BeginAnimation(Window.OpacityProperty, anim);
                    }
                    catch (Exception exStart)
                    {
                        System.Diagnostics.Debug.WriteLine("Ошибка запуска анимации закрытия Splash: " + exStart);
                        try { this.Close(); } catch { }
                        tcs.TrySetResult(null);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Ошибка в CloseSplashAsync: " + ex);
                    try { this.Close(); } catch { }
                    tcs.TrySetResult(null);
                }
            }), System.Windows.Threading.DispatcherPriority.Normal);

            return tcs.Task;
        }
    }
}
