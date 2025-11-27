using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class SplashScreen : Window
    {
        private readonly string _gifPath;
        private MediaElement _gifAnimation;

        public SplashScreen(string gifPath = null)
        {
            InitializeComponent();
            _gifPath = gifPath ?? FindDefaultGif();

            InitializeAnimation();
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
                        Stretch = System.Windows.Media.Stretch.Uniform,
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
                var progressBar = (ProgressBar)this.FindName("progressBar");
                if (progressBar != null)
                {
                    if (progress >= 0 && progress <= 100)
                    {
                        progressBar.IsIndeterminate = false;
                        progressBar.Value = progress;
                    }
                    else
                    {
                        progressBar.IsIndeterminate = true;
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

        public void CloseSplash()
        {
            Dispatcher.Invoke(() =>
            {
                System.Diagnostics.Debug.WriteLine("CloseSplash вызван");

                // Останавливаем анимацию перед закрытием
                if (_gifAnimation != null)
                {
                    _gifAnimation.Stop();
                    _gifAnimation.MediaEnded -= GifAnimation_MediaEnded;
                }

                // ★ ПРОСТОЕ И БЫСТРОЕ ЗАКРЫТИЕ БЕЗ АНИМАЦИИ ★
                this.Close();
            });
        }
    }
}