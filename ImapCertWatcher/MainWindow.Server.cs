using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private void SetServerState(ServerConnectionState state)
        {
            if (_serverState == state && state != ServerConnectionState.Connecting)
                return;

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.InvokeAsync(() => SetServerState(state));
                return;
            }

            _serverState = state;

            if (state != ServerConnectionState.Busy)
                pbServerProgress.Visibility = Visibility.Collapsed;

            if (state == ServerConnectionState.Connecting)
                StartConnectingAnimation();
            else
                StopConnectingAnimation();

            switch (state)
            {
                case ServerConnectionState.Connecting:
                    txtServerStatus.Text = "🟡 Подключение к серверу...";
                    txtServerStatus.Foreground = Brushes.Gold;
                    txtNextCheck.Visibility = Visibility.Visible;
                    ApplyBusyUiState(true);
                    break;

                case ServerConnectionState.Online:
                    txtServerStatus.Text = "🟢 Сервер готов";
                    txtServerStatus.Foreground = Brushes.LimeGreen;
                    txtNextCheck.Visibility = Visibility.Visible;
                    ApplyOfflineUiState(false);
                    ApplyBusyUiState(false);
                    break;

                case ServerConnectionState.Busy:
                    txtServerStatus.Text = "🟠 Сервер работает";
                    txtServerStatus.Foreground = Brushes.Orange;
                    txtNextCheck.Visibility = Visibility.Collapsed;
                    ApplyOfflineUiState(false);
                    ApplyBusyUiState(true);
                    break;

                case ServerConnectionState.Offline:
                    txtServerStatus.Text = "🔴 Сервер OFFLINE";
                    txtServerStatus.Foreground = Brushes.Red;
                    txtNextCheck.Text = "Следующая проверка: --:--";
                    txtNextCheck.Visibility = Visibility.Visible;
                    ApplyOfflineUiState(true);
                    ApplyBusyUiState(true);
                    break;
            }

            AdjustPollingInterval();
        }

        private void ApplyBusyUiState(bool isBusy)
        {
            if (btnCheckNewCerts != null)
                btnCheckNewCerts.IsEnabled = !isBusy;

            if (btnCheckNewRevokes != null)
                btnCheckNewRevokes.IsEnabled = !isBusy;

            if (grpMaintenance != null)
                grpMaintenance.IsEnabled = !isBusy && !_isOfflineUiLocked;
        }

        private async void ServerMonitorTimer_Tick(object sender, EventArgs e)
        {
            await UpdateServerMonitor();
        }

        private void StartServerMonitor()
        {
            if (_serverMonitorTimer != null)
                return;

            _serverMonitorTimer = new DispatcherTimer();
            _serverMonitorTimer.Interval = PollConnecting;
            _serverMonitorTimer.Tick += ServerMonitorTimer_Tick;
            _serverMonitorTimer.Start();

            _ = UpdateServerMonitor();
        }

        private async Task ReloadDataAfterServerWorkAsync()
        {
            try
            {
                statusText.Text = "Обновление данных после проверки...";
                AddToMiniLog("Сервер завершил работу, перечитываем данные");

                await LoadFromServer();

                statusText.Text = "Данные обновлены";
                AddToMiniLog("Данные обновлены после завершения проверки");
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка автообновления: " + ex.Message);
            }
        }

        private async Task ReloadDataAfterReconnectAsync()
        {
            try
            {
                statusText.Text = "Связь восстановлена, обновляем данные...";
                AddToMiniLog("Связь с сервером восстановлена, перечитываем данные");

                await LoadFromServer();

                statusText.Text = "Данные обновлены после восстановления связи";
                AddToMiniLog("Данные обновлены после восстановления связи");
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка обновления после восстановления связи: " + ex.Message);
            }
        }

        private void AdjustPollingInterval()
        {
            if (_serverMonitorTimer == null)
                return;

            switch (_serverState)
            {
                case ServerConnectionState.Busy:
                    _serverMonitorTimer.Interval = PollBusy;
                    break;

                case ServerConnectionState.Online:
                    _serverMonitorTimer.Interval = PollOnline;
                    break;

                case ServerConnectionState.Connecting:
                    _serverMonitorTimer.Interval = PollConnecting;
                    break;

                case ServerConnectionState.Offline:
                    _serverMonitorTimer.Interval = TimeSpan.FromSeconds(_reconnectDelaySeconds);
                    break;
            }
        }

        private async Task UpdateServerMonitor()
        {
            if (_api == null)
                return;

            if (Interlocked.Exchange(ref _monitorBusyFlag, 1) == 1)
                return;

            try
            {
                var response = (await _api.SendCommand("STATUS"))?.Trim();

                if (string.IsNullOrWhiteSpace(response) || !response.StartsWith("STATE|"))
                {
                    await StartReconnectBackoff();
                    return;
                }

                ParseServerState(response);
            }
            catch
            {
                await StartReconnectBackoff();
            }
            finally
            {
                Interlocked.Exchange(ref _monitorBusyFlag, 0);
            }
        }

        private void ParseServerState(string stateLine)
        {
            try
            {
                var parts = stateLine.Split('|');

                bool isBusy = false;
                int progress = 0;
                int timerMinutes = 0;
                string stage = "";

                foreach (var part in parts)
                {
                    if (part.StartsWith("BUSY="))
                        isBusy = part.Substring(5) == "1";
                    else if (part.StartsWith("PROGRESS="))
                        int.TryParse(part.Substring(9), out progress);
                    else if (part.StartsWith("TIMER="))
                        int.TryParse(part.Substring(6), out timerMinutes);
                    else if (part.StartsWith("STAGE="))
                        stage = part.Substring(6);
                }

                bool wasBusy = _serverWasBusy;
                bool wasOffline = _isOfflineUiLocked;

                _serverWasBusy = isBusy;
                _reconnectDelaySeconds = 2;

                if (wasOffline)
                    _reloadAfterReconnectWhenIdle = true;

                SetServerState(isBusy
                    ? ServerConnectionState.Busy
                    : ServerConnectionState.Online);

                if (isBusy)
                {
                    pbServerProgress.Visibility = Visibility.Visible;
                    pbServerProgress.Value = progress;

                    if (!string.IsNullOrWhiteSpace(stage) &&
                        !string.Equals(stage, "IDLE", StringComparison.OrdinalIgnoreCase))
                    {
                        txtServerStatus.Text = "🟠 Сервер работает: " + stage;
                    }
                }
                else
                {
                    pbServerProgress.Visibility = Visibility.Collapsed;
                    txtNextCheck.Text = timerMinutes <= 0
                        ? "Следующая проверка: сейчас"
                        : $"Следующая проверка через {timerMinutes} мин";
                }

                if (wasBusy && !isBusy && _reloadAfterServerWork)
                {
                    _reloadAfterServerWork = false;
                    _reloadAfterReconnectWhenIdle = false;
                    _ = ReloadDataAfterServerWorkAsync();
                }
                else if (!isBusy && _reloadAfterReconnectWhenIdle && !_isLoadingFromServer)
                {
                    _reloadAfterReconnectWhenIdle = false;
                    _ = ReloadDataAfterReconnectAsync();
                }
            }
            catch
            {
                SetServerState(ServerConnectionState.Offline);
            }
        }

        private void StartConnectingAnimation()
        {
            if (_connectingAnimationTimer == null)
            {
                _connectingAnimationTimer = new DispatcherTimer();
                _connectingAnimationTimer.Interval = TimeSpan.FromMilliseconds(400);
                _connectingAnimationTimer.Tick += ConnectingAnimationTick;
            }

            if (!_connectingAnimationTimer.IsEnabled)
                _connectingAnimationTimer.Start();
        }

        private void ConnectingAnimationTick(object sender, EventArgs e)
        {
            if (_serverState != ServerConnectionState.Connecting)
                return;

            _connectingDots = (_connectingDots + 1) % 4;
            txtServerStatus.Text = "🟡 Подключение к серверу" + new string('.', _connectingDots);
        }

        private void StopConnectingAnimation()
        {
            if (_connectingAnimationTimer != null)
            {
                _connectingAnimationTimer.Stop();
                _connectingDots = 0;
            }
        }

        private async Task SendServerCommand(string cmd)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            bool isCheckCommand =
                cmd == "FAST_CERTS_CHECK" ||
                cmd == "FAST_REVOKES_CHECK" ||
                cmd == "FULL_CERTS_CHECK" ||
                cmd == "FULL_REVOKES_CHECK" ||
                cmd == "FAST_CHECK" ||
                cmd == "FULL_CHECK";

            try
            {
                statusText.Text = "Связь с сервером...";

                var response = (await _api.SendCommand(cmd))?.Trim();

                if (string.IsNullOrWhiteSpace(response))
                {
                    statusText.Text = "Нет ответа от сервера";
                    AddToMiniLog("SERVER: <пустой ответ>");
                    MessageBox.Show(
                        "Сервер вернул пустой ответ",
                        "Предупреждение",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                AddToMiniLog("SERVER: " + response);

                if (response.StartsWith("STARTED", StringComparison.OrdinalIgnoreCase))
                {
                    if (isCheckCommand)
                    {
                        _reloadAfterServerWork = true;
                        _serverWasBusy = true;
                        SetServerState(ServerConnectionState.Busy);
                    }

                    if (cmd == "FAST_CERTS_CHECK")
                        statusText.Text = "Запущена проверка новых сертификатов";
                    else if (cmd == "FAST_REVOKES_CHECK")
                        statusText.Text = "Запущена проверка новых аннулирований";
                    else if (cmd == "FULL_CERTS_CHECK")
                        statusText.Text = "Запущена полная перепроверка сертификатов";
                    else if (cmd == "FULL_REVOKES_CHECK")
                        statusText.Text = "Запущена полная перепроверка аннулирований";
                    else if (cmd == "FAST_CHECK")
                        statusText.Text = "Запущена быстрая проверка";
                    else if (cmd == "FULL_CHECK")
                        statusText.Text = "Запущена полная проверка";
                    else
                        statusText.Text = "Команда запущена";

                    return;
                }

                if (response.StartsWith("OK"))
                {
                    if (cmd == "RESET_REVOKES")
                    {
                        statusText.Text = "Аннулирования сброшены";
                        await LoadFromServer();
                        AddToMiniLog("Данные обновлены после RESET_REVOKES");
                    }
                    else if (isCheckCommand)
                    {
                        statusText.Text = "Проверка завершена";
                        SetServerState(ServerConnectionState.Online);
                    }
                    else
                    {
                        statusText.Text = "Команда выполнена";
                    }

                    return;
                }


                if (response.StartsWith("BUSY WAIT", StringComparison.OrdinalIgnoreCase))
                {
                    _reloadAfterServerWork = false;
                    _serverWasBusy = false;

                    SetServerState(ServerConnectionState.Online);
                    statusText.Text = "Повторную проверку пока запускать рано";
                    AddToMiniLog("Сервер отклонил запуск: " + response);
                    return;
                }

                if (response.StartsWith("BUSY", StringComparison.OrdinalIgnoreCase))
                {
                    if (isCheckCommand)
                    {
                        _reloadAfterServerWork = true;
                        _serverWasBusy = true;
                    }

                    SetServerState(ServerConnectionState.Busy);
                    statusText.Text = "Сервер занят";
                    return;
                }

                if (response.StartsWith("CANCELLED"))
                {
                    statusText.Text = "Операция отменена";
                    AddToMiniLog("Сервер: " + response);
                    return;
                }

                if (response.StartsWith("ERROR"))
                {
                    statusText.Text = "Ошибка сервера";
                    MessageBox.Show(
                        response,
                        "Сервер",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                statusText.Text = "Ответ сервера: " + response;
            }
            catch (Exception ex)
            {
                if (isCheckCommand)
                {
                    try
                    {
                        await Task.Delay(700);

                        var status = (await _api.SendCommand("STATUS"))?.Trim();

                        if (!string.IsNullOrWhiteSpace(status) &&
                            status.StartsWith("STATE|") &&
                            status.Contains("BUSY=1"))
                        {
                            _reloadAfterServerWork = true;
                            _serverWasBusy = true;

                            SetServerState(ServerConnectionState.Busy);

                            if (cmd == "FAST_CERTS_CHECK")
                                statusText.Text = "Проверка новых сертификатов запущена";
                            else if (cmd == "FAST_REVOKES_CHECK")
                                statusText.Text = "Проверка новых аннулирований запущена";
                            else if (cmd == "FULL_CERTS_CHECK")
                                statusText.Text = "Полная перепроверка сертификатов запущена";
                            else if (cmd == "FULL_REVOKES_CHECK")
                                statusText.Text = "Полная перепроверка аннулирований запущена";
                            else
                                statusText.Text = "Проверка запущена";

                            AddToMiniLog("Сервер выполняет команду: " + cmd);
                            return;
                        }
                    }
                    catch
                    {
                    }
                }

                MessageBox.Show($"Ошибка подключения:\n{ex.Message}");
                statusText.Text = "Сервер недоступен";
                SetServerState(ServerConnectionState.Offline);
            }
        }

        private async Task StartReconnectBackoff()
        {
            if (Interlocked.Exchange(ref _reconnectBusy, 1) == 1)
                return;

            try
            {
                if (_serverState != ServerConnectionState.Offline)
                    SetServerState(ServerConnectionState.Offline);

                await Task.Delay(_reconnectDelaySeconds * 1000);

                if (_serverState == ServerConnectionState.Offline)
                {
                    _reconnectDelaySeconds =
                        Math.Min(_reconnectDelaySeconds * 2, MAX_RECONNECT_DELAY);

                    SetServerState(ServerConnectionState.Connecting);
                }
            }
            finally
            {
                Interlocked.Exchange(ref _reconnectBusy, 0);
            }
        }

        private async void BtnManualCheck_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FAST_CHECK");
        }

        private async void BtnProcessAll_Click(object sender, RoutedEventArgs e)
        {
            var r = MessageBox.Show(
                "Полная проверка может занять длительное время.\n\nЗапустить?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            await SendServerCommand("FULL_CHECK");
        }

        private async void BtnCheckNewCerts_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FAST_CERTS_CHECK");
        }

        private async void BtnCheckNewRevokes_Click(object sender, RoutedEventArgs e)
        {
            await SendServerCommand("FAST_REVOKES_CHECK");
        }

        private async void BtnFullRecheckCerts_Click(object sender, RoutedEventArgs e)
        {
            var r = MessageBox.Show(
                "Будет выполнен полный обход всей папки сертификатов.\n\nЗапустить?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            await SendServerCommand("FULL_CERTS_CHECK");
        }

        private async void BtnFullRecheckRevokes_Click(object sender, RoutedEventArgs e)
        {
            var r = MessageBox.Show(
                "Будет выполнен полный обход всей папки аннулирований.\n\nЗапустить?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r != MessageBoxResult.Yes)
                return;

            await SendServerCommand("FULL_REVOKES_CHECK");
        }

        private async void BtnResetRevokes_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            var r1 = MessageBox.Show(
                "Эта операция удалит ВСЮ информацию об аннулировании сертификатов.\n\nПродолжить?",
                "ВНИМАНИЕ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (r1 != MessageBoxResult.Yes)
                return;

            var r2 = MessageBox.Show(
                "Вы ТОЧНО уверены?\n\nОперация необратима!",
                "ПОДТВЕРЖДЕНИЕ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Error);

            if (r2 != MessageBoxResult.Yes)
                return;

            try
            {
                var resp = await _api.ResetRevokes();

                if (string.IsNullOrWhiteSpace(resp))
                {
                    MessageBox.Show("Нет ответа от сервера");
                    statusText.Text = "Нет ответа от сервера";
                    return;
                }

                AddToMiniLog("RESET_REVOKES: " + resp);

                if (resp.StartsWith("BUSY", StringComparison.OrdinalIgnoreCase))
                {
                    statusText.Text = "Сервер занят";
                    MessageBox.Show(
                        "Сервер сейчас выполняет другую операцию.\nСброс аннулирований пока недоступен.",
                        "Сервер занят",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                if (!resp.StartsWith("OK", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        "Ошибка сброса аннулирований:\n" + resp,
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    statusText.Text = "Ошибка сброса аннулирований";
                    return;
                }

                await LoadFromServer();
                statusText.Text = "Аннулирования сброшены";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка сброса аннулирований:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                statusText.Text = "Ошибка сброса аннулирований";
                AddToMiniLog("RESET_REVOKES ERROR: " + ex.Message);
            }
        }
    }
}