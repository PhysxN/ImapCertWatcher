using ImapCertWatcher.Client;
using ImapCertWatcher.Services;
using ImapCertWatcher.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private void OnProgressUpdated(string message, double progress)
        {
            ProgressUpdated?.Invoke(message, progress);
        }

        private async Task ReportProgressStep(string text, double percent, int minDelayMs = 250)
        {
            OnProgressUpdated(text, percent);
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            await Task.Delay(minDelayMs);
        }

        public async Task InitializeAsync()
        {
            _ = CleanupCertsFolderAsync();

            System.Diagnostics.Debug.WriteLine("MainWindow.InitializeAsync начат");

            try
            {
                await ReportProgressStep("Подготовка среды...", 5);

                await ReportProgressStep("Инициализация API...", 10);
                _api = new ServerApiClient(_clientSettings);

                _tokenService = new TokenService(_api);
                TokensVm = new TokensViewModel(_tokenService);

                _tokenService.TokensChanged += RebuildAvailableTokens;

                await ReportProgressStep("Подключение к серверу...", 20);

                try
                {
                    await _tokenService.Reload();
                }
                catch (Exception ex)
                {
                    SetServerState(ServerConnectionState.Offline);
                    AddToMiniLog("Сервер недоступен, токены не загружены: " + ex.Message);
                }

                await LoadFromServer(false);

                _certsLoadedOnce = _serverState == ServerConnectionState.Online;

                System.Diagnostics.Debug.WriteLine("Данные получены");

                await ReportProgressStep("Анализ полученных данных...", 35);
                _knownCertIds = new HashSet<int>(_allItems.Select(r => r.Id));

                await ReportProgressStep("Формирование списков токенов...", 50);
                RebuildAvailableTokens();

                await ReportProgressStep("Применение фильтров...", 60);
                ApplySearchFilter();

                await ReportProgressStep("Настройка таблицы...", 70);
                AutoFitDataGridColumns();

                await ReportProgressStep("Запуск таймеров...", 80);
                InitializeRefreshTimer();

                await ReportProgressStep("Загрузка логов...", 90);
                await LoadLogs();

                await ReportProgressStep("Запуск мониторинга сервера...", 95);
                StartServerMonitor();

                await ReportProgressStep("Готово", 100, 400);

                statusText.Text = _serverState == ServerConnectionState.Online
                    ? "Готово"
                    : "Ожидание сервера...";

                AddToMiniLog($"Инициализация завершена. Загружено записей: {_allItems.Count}");

                System.Diagnostics.Debug.WriteLine("MainWindow.InitializeAsync завершен");

                DataLoaded?.Invoke();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в MainWindow.InitializeAsync: {ex.Message}");

                OnProgressUpdated($"Ошибка: {ex.Message}", -1);

                statusText.Text = "Ошибка инициализации";
                AddToMiniLog($"Ошибка инициализации: {ex.Message}");

                MessageBox.Show(
                    $"Ошибка инициализации: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                StartServerMonitor();
                DataLoaded?.Invoke();

                throw;
            }
        }

        private void InitializeRefreshTimer()
        {
            try
            {
                if (_refreshTimer != null)
                {
                    _refreshTimer.Stop();
                    _refreshTimer.Tick -= RefreshDaysLeft;
                }

                _refreshTimer = new DispatcherTimer();
                _refreshTimer.Interval = TimeSpan.FromMinutes(10);
                _refreshTimer.Tick += RefreshDaysLeft;
                _refreshTimer.Start();

                AddToMiniLog("Таймер обновления дней запущен (интервал: 10 минут)");
            }
            catch (Exception ex)
            {
                Log($"Ошибка инициализации таймера обновления: {ex.Message}");
            }
        }

        private void RefreshDaysLeft(object sender, EventArgs e)
        {
            try
            {
                foreach (var item in _items)
                    item?.RefreshDaysLeft();
            }
            catch (Exception ex)
            {
                Log($"Ошибка при обновлении дней: {ex.Message}");
            }
        }

        private void ManualRefreshDaysLeft()
        {
            try
            {
                RefreshDaysLeft(null, EventArgs.Empty);
                AddToMiniLog("Ручное обновление дней выполнено");
            }
            catch (Exception ex)
            {
                Log($"Ошибка при ручном обновлении дней: {ex.Message}");
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_serverMonitorTimer != null)
            {
                _serverMonitorTimer.Stop();
                _serverMonitorTimer.Tick -= ServerMonitorTimer_Tick;
            }

            if (_connectingAnimationTimer != null)
            {
                _connectingAnimationTimer.Stop();
                _connectingAnimationTimer.Tick -= ConnectingAnimationTick;
            }

            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer.Tick -= RefreshDaysLeft;
            }

            if (_tokenService != null)
                _tokenService.TokensChanged -= RebuildAvailableTokens;

            base.OnClosed(e);
        }
    }
}