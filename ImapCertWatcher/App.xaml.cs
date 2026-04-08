using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Animation;

namespace ImapCertWatcher.Client
{
    public partial class App : Application
    {
        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;

            SplashScreen splash = null;
            MainWindow mainWindow = null;

            try
            {
                splash = new SplashScreen();
                splash.Show();

                mainWindow = new MainWindow();
                MainWindow = mainWindow;

                mainWindow.ProgressUpdated += (message, progress) =>
                {
                    splash.UpdateStatus(message);
                    splash.UpdateProgress(progress);
                };

                await mainWindow.InitializeAsync();

                await splash.CloseSplashAsync();

                mainWindow.Opacity = 0;
                mainWindow.Show();
                mainWindow.Activate();

                var fade = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromMilliseconds(250)
                };

                mainWindow.BeginAnimation(Window.OpacityProperty, fade);

                ShutdownMode = ShutdownMode.OnMainWindowClose;
            }
            catch (Exception ex)
            {
                try
                {
                    if (splash != null)
                        await splash.CloseSplashAsync();
                }
                catch
                {
                }

                MessageBox.Show(
                    "Ошибка запуска приложения:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                Shutdown(-1);
            }
        }
    }
}