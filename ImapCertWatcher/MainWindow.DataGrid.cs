using ImapCertWatcher.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Input;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private bool _isSwitchingTabInternally;
        private void DgCerts_Loaded(object sender, RoutedEventArgs e)
        {
            AutoFitDataGridColumns();
        }

        private void AutoFitDataGridColumns()
        {
            try
            {
                if (dgCerts == null || dgCerts.Columns == null)
                    return;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        DataGridColumn noteColumn = null;

                        foreach (var column in dgCerts.Columns)
                        {
                            string header = column.Header?.ToString() ?? "";

                            if (string.Equals(header, "Примечание", StringComparison.Ordinal))
                            {
                                noteColumn = column;
                                continue;
                            }

                            if (column is DataGridTextColumn textColumn)
                            {
                                AutoFitTextColumn(textColumn);
                            }
                            else if (column is DataGridTemplateColumn templateColumn)
                            {
                                AutoFitTemplateColumn(templateColumn);
                            }
                            else if (column is DataGridCheckBoxColumn checkBoxColumn)
                            {
                                AutoFitCheckBoxColumn(checkBoxColumn);
                            }
                        }

                        if (noteColumn != null)
                            noteColumn.Width = new DataGridLength(1, DataGridLengthUnitType.Star);

                        dgCerts.UpdateLayout();
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка автоподбора столбцов: {ex.Message}");
                    }
                }), DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                Log($"Ошибка в AutoFitDataGridColumns: {ex.Message}");
            }
        }

        private void AutoFitTextColumn(DataGridTextColumn column)
        {
            try
            {
                column.Width = DataGridLength.Auto;
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора TextColumn: {ex.Message}");
            }
        }

        private void AutoFitTemplateColumn(DataGridTemplateColumn column)
        {
            try
            {
                column.Width = DataGridLength.Auto;
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора TemplateColumn: {ex.Message}");
            }
        }

        private void AutoFitCheckBoxColumn(DataGridCheckBoxColumn column)
        {
            try
            {
                column.Width = new DataGridLength(80);
            }
            catch (Exception ex)
            {
                Log($"Ошибка автоподбора CheckBoxColumn: {ex.Message}");
            }
        }

        private async Task<bool> LoadFromServer(bool showErrorDialog = false)
        {
            if (_isLoadingFromServer)
            {
                _reloadRequestedWhileLoading = true;
                return false;
            }

            _reloadRequestedWhileLoading = false;
            _isLoadingFromServer = true;
            SetServerState(ServerConnectionState.Connecting);

            try
            {
                if (_api == null)
                {
                    AddToMiniLog("API ещё не инициализирован");
                    SetServerState(ServerConnectionState.Offline);
                    statusText.Text = "Ошибка: API не инициализирован";

                    if (showErrorDialog)
                    {
                        MessageBox.Show(
                            "Не удалось получить список сертификатов с сервера.\n\nAPI не инициализирован.",
                            "GET_CERTS",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }

                    return false;
                }

                statusText.Text = "Загрузка данных с сервера...";
                AddToMiniLog("Запрос GET_CERTS отправлен");

                var list = await _api.GetCertificates() ?? new List<CertRecord>();

                if (_tokenService != null)
                {
                    try
                    {
                        await _tokenService.Reload();
                    }
                    catch (Exception tokenEx)
                    {
                        AddToMiniLog("Ошибка загрузки токенов: " + tokenEx.Message);

                        _tokenService.Tokens.Clear();
                    }
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    _isApplyingFromServer = true;

                    try
                    {
                        var allTokens = (_tokenService?.Tokens?.ToList() ?? new List<TokenRecord>())
                            .Where(t => t != null)
                            .GroupBy(t => t.Id)
                            .Select(g => g.First())
                            .OrderBy(t => t.Sn)
                            .ToList();

                        _allItems.Clear();

                        foreach (var item in list)
                        {
                            var allowedTokens = allTokens
                                .Where(t => !t.OwnerCertId.HasValue || t.OwnerCertId.Value == item.Id)
                                .Select(t => new TokenRecord
                                {
                                    Id = t.Id,
                                    Sn = t.Sn,
                                    OwnerCertId = t.OwnerCertId,
                                    OwnerFio = t.OwnerFio
                                })
                                .ToList();

                            if (item.TokenId.HasValue &&
                                allowedTokens.All(t => t.Id != item.TokenId.Value) &&
                                !string.IsNullOrWhiteSpace(item.TokenSn))
                            {
                                allowedTokens.Insert(0, new TokenRecord
                                {
                                    Id = item.TokenId.Value,
                                    Sn = item.TokenSn,
                                    OwnerCertId = item.Id,
                                    OwnerFio = item.Fio
                                });
                            }

                            item.AvailableTokens = new ObservableCollection<TokenRecord>(allowedTokens);

                            if (item.TokenId.HasValue)
                            {
                                var selectedToken = item.AvailableTokens.FirstOrDefault(t => t.Id == item.TokenId.Value);

                                if (selectedToken != null)
                                {
                                    item.SelectedToken = selectedToken;
                                }
                                else
                                {
                                    item.SelectedToken = null;
                                }
                            }
                            else
                            {
                                item.SelectedToken = null;
                            }

                            _allItems.Add(item);
                        }

                        ApplySearchFilter();
                        dgCerts.Items.Refresh();
                        UpdateDeleteRestoreButtonState();

                        statusText.Text = $"Загружено записей: {_allItems.Count}";
                        SetServerState(ServerConnectionState.Online);
                    }
                    finally
                    {
                        _isApplyingFromServer = false;
                    }
                });

                AddToMiniLog($"Получено с сервера: {list.Count} записей");
                return true;
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка загрузки: " + ex.Message);
                SetServerState(ServerConnectionState.Offline);
                statusText.Text = "Ошибка загрузки с сервера: " + ex.Message;

                if (showErrorDialog)
                {
                    MessageBox.Show(
                        "Не удалось получить список сертификатов с сервера.\n\n" + ex.Message,
                        "GET_CERTS",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }

                return false;
            }
            finally
            {
                _isLoadingFromServer = false;

                if (_reloadRequestedWhileLoading)
                {
                    _reloadRequestedWhileLoading = false;

                    var _ = Dispatcher.BeginInvoke(
                        new Action(RunDeferredReloadFromServer),
                        DispatcherPriority.Background);
                }
            }
        }

        private void RunDeferredReloadFromServer()
        {
            var _ = LoadFromServer();
        }

        private void DgCerts_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;

            while (source != null && !(source is DataGridRow))
                source = VisualTreeHelper.GetParent(source);

            var row = source as DataGridRow;
            if (row == null)
                return;

            row.IsSelected = true;
            dgCerts.SelectedItem = row.Item;
            row.Focus();
        }
        private void ApplySearchFilter()
        {
            try
            {
                if (_items == null || _allItems == null || searchStatusText == null)
                    return;

                _items.Clear();

                IEnumerable<CertRecord> query = _allItems.Where(x => x != null);

                if (!_showDeleted)
                {
                    query = query.Where(x =>
                        !x.IsDeleted &&
                        x.DateEnd != DateTime.MinValue &&
                        x.DaysLeft >= 0);
                }

                if (_currentBuildingFilter != "Все")
                {
                    query = query.Where(x =>
                        string.Equals(
                            x.Building,
                            _currentBuildingFilter,
                            StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    var terms = _searchText
                        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    query = query.Where(item =>
                    {
                        var fio = item.Fio ?? "";
                        return terms.All(t => fio.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);
                    });
                }

                foreach (var item in query)
                    _items.Add(item);

                searchStatusText.Text = $"Показано: {_items.Count} из {_allItems.Count}";
                UpdateDeleteRestoreButtonState();

                Dispatcher.BeginInvoke(
                    new Action(AutoFitDataGridColumns),
                    DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                Log("Ошибка фильтра: " + ex.Message);
            }
        }

        private void TxtSearchFio_TextChanged(object sender, TextChangedEventArgs e)
        {
            _searchText = txtSearchFio.Text.Trim();
            ApplySearchFilter();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            txtSearchFio.Text = "";
            _searchText = "";
            ApplySearchFilter();
            AddToMiniLog("Поиск очищен");
        }

        private void BuildingFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAllBuildings.IsChecked == true)
                _currentBuildingFilter = "Все";
            else if (rbKrasnoflotskaya.IsChecked == true)
                _currentBuildingFilter = "Краснофлотская";
            else if (rbPionerskaya.IsChecked == true)
                _currentBuildingFilter = "Пионерская";

            ApplySearchFilter();
            AddToMiniLog($"Фильтр: {_currentBuildingFilter}");
        }

        private void ChkShowDeleted_Changed(object sender, RoutedEventArgs e)
        {
            _showDeleted = chkShowDeleted.IsChecked == true;
            ApplySearchFilter();
            AddToMiniLog(_showDeleted
                ? "Показаны все сертификаты"
                : "Показаны только актуальные сертификаты");
        }

        private void UpdateDeleteRestoreButtonState()
        {
            if (btnDeleteRestoreSelected == null)
                return;

            var record = dgCerts?.SelectedItem as CertRecord;

            if (record == null || record.IsRevoked)
            {
                btnDeleteRestoreSelected.Visibility = Visibility.Collapsed;
                return;
            }

            btnDeleteRestoreSelected.Visibility = Visibility.Visible;
            btnDeleteRestoreSelected.Content = record.IsDeleted
                ? "Восстановить"
                : "Пометить на удаление";
        }

        private void DgCerts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDeleteRestoreButtonState();
        }


        private async void BuildingQuickButton_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_isApplyingFromServer)
                return;

            if (_isSavingBuilding)
                return;

            if (!(sender is Button button))
                return;

            if (!(button.DataContext is CertRecord record))
                return;

            string newValue = button.Tag as string ?? "";
            string oldValue = record.Building ?? "";

            if (string.Equals(newValue, oldValue, StringComparison.Ordinal))
                return;

            _isSavingBuilding = true;

            try
            {
                var resp = await _api.UpdateBuilding(record.Id, newValue);

                if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                {
                    MessageBox.Show("Ошибка сохранения здания");
                    return;
                }

                record.Building = newValue;

                AddToMiniLog(
                    $"Здание изменено: {record.Fio} → {(string.IsNullOrWhiteSpace(newValue) ? "<пусто>" : newValue)}");

                statusText.Text = "Здание сохранено";
                dgCerts.Items.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения здания:\n" + ex.Message);
            }
            finally
            {
                _isSavingBuilding = false;
            }
        }


        private void BtnDeleteRestoreSelected_Click(object sender, RoutedEventArgs e)
        {
            if (!(dgCerts.SelectedItem is CertRecord record))
                return;

            if (record.IsRevoked)
                return;

            if (record.IsDeleted)
                BtnRestoreSelected_Click(sender, e);
            else
                BtnDeleteSelected_Click(sender, e);
        }

        private async void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgCerts.SelectedItem is CertRecord record)
            {
                if (_api == null)
                {
                    MessageBox.Show("Сервер не подключен");
                    return;
                }

                try
                {
                    var result = MessageBox.Show(
                        $"Пометить сертификат:\n\n{record.Fio}\n\nкак удалённый?",
                        "Подтверждение",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result != MessageBoxResult.Yes)
                        return;

                    Log($"Пометка на удаление: {record.Fio} (ID: {record.Id})");
                    var resp = await _api.MarkDeleted(record.Id, true);

                    if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка удаления");
                        return;
                    }

                    record.IsDeleted = true;
                    ApplySearchFilter();
                    UpdateDeleteRestoreButtonState();

                    statusText.Text = "Запись помечена на удаление";
                    AddToMiniLog($"Удален: {record.Fio}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Ошибка при пометке на удаление: {ex.Message}",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    AddToMiniLog($"Ошибка удаления: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show(
                    "Выберите запись для пометки на удаление",
                    "Информация",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        private async void BtnRestoreSelected_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (!(dgCerts.SelectedItem is CertRecord record))
            {
                MessageBox.Show(
                    "Выберите запись для восстановления",
                    "Информация",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                var resp = await _api.MarkDeleted(record.Id, false);

                if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                {
                    MessageBox.Show("Ошибка восстановления:\n" + (resp ?? "Нет ответа сервера"));
                    return;
                }

                await LoadFromServer();

                statusText.Text = "Запись восстановлена";
                AddToMiniLog("Восстановлено: " + record.Fio);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка восстановления:\n" + ex.Message);
            }
        }

        private void DgCerts_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is CertRecord record)
            {
                if (record.IsRevoked)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 90, 60, 60)
                        : Color.FromArgb(255, 255, 180, 180));

                    e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                        ? Colors.White
                        : Colors.Black);

                    return;
                }

                if (record.IsDeleted)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 80, 80, 80)
                        : Color.FromArgb(255, 220, 220, 220));
                }
                else if (record.IsExpired)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 85, 55, 25)
                        : Color.FromArgb(255, 255, 230, 190));
                }
                else if (record.DateEnd == DateTime.MinValue)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 70, 70, 90)
                        : Color.FromArgb(255, 230, 230, 245));
                }
                else if (record.DaysLeft <= 10)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 120, 50, 50)
                        : Color.FromArgb(255, 255, 200, 200));
                }
                else if (record.DaysLeft <= 30)
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 120, 120, 50)
                        : Color.FromArgb(255, 255, 255, 200));
                }
                else
                {
                    e.Row.Background = new SolidColorBrush(_isDarkTheme
                        ? Color.FromArgb(255, 50, 80, 50)
                        : Color.FromArgb(255, 200, 255, 200));
                }

                e.Row.Foreground = new SolidColorBrush(_isDarkTheme
                    ? Colors.White
                    : Colors.Black);
            }
        }

        private void DgCerts_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (!(dgCerts.SelectedItem is CertRecord record))
            {
                miAddArchive.IsEnabled = false;
                miAddArchive.ToolTip = null;
                miExportSelected.IsEnabled = false;
                return;
            }

            if (record.HasArchive)
            {
                miAddArchive.IsEnabled = false;
                miAddArchive.ToolTip = "Архив уже прикреплен";
            }
            else
            {
                miAddArchive.IsEnabled = true;
                miAddArchive.ToolTip = "Добавить архив с подписью";
            }

            miExportSelected.IsEnabled = true;
        }

        private bool IsBlockedOfflineTab(TabItem tab)
        {
            if (!_isOfflineUiLocked)
                return false;

            if (tab == null)
                return false;

            var header = tab.Header?.ToString() ?? "";
            return !string.Equals(header, "Настройки", StringComparison.Ordinal);
        }

        private static TabItem FindParentTabItem(DependencyObject source)
        {
            while (source != null)
            {
                if (source is TabItem tabItem)
                    return tabItem;

                source = VisualTreeHelper.GetParent(source);
            }

            return null;
        }

        private void MainTabControl_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!_isOfflineUiLocked)
                return;

            var source = e.OriginalSource as DependencyObject;
            var clickedTab = FindParentTabItem(source);

            if (clickedTab == null)
                return;

            if (!IsBlockedOfflineTab(clickedTab))
                return;

            e.Handled = true;

            _isSwitchingTabInternally = true;
            try
            {
                mainTabControl.SelectedItem = tabSettings;
            }
            finally
            {
                _isSwitchingTabInternally = false;
            }
        }

        private void MainTabControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!_isOfflineUiLocked)
                return;

            switch (e.Key)
            {
                case Key.Left:
                case Key.Right:
                case Key.Tab:
                case Key.Home:
                case Key.End:
                    e.Handled = true;

                    _isSwitchingTabInternally = true;
                    try
                    {
                        mainTabControl.SelectedItem = tabSettings;
                    }
                    finally
                    {
                        _isSwitchingTabInternally = false;
                    }
                    break;
            }
        }

        private async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(sender, mainTabControl))
                return;

            if (_isSwitchingTabInternally)
                return;

            if (!(mainTabControl.SelectedItem is TabItem tab))
                return;

            if (IsBlockedOfflineTab(tab))
            {
                _isSwitchingTabInternally = true;
                try
                {
                    mainTabControl.SelectedItem = tabSettings;
                }
                finally
                {
                    _isSwitchingTabInternally = false;
                }

                return;
            }

            var header = tab.Header?.ToString();

            try
            {
                if (header == "Сертификаты")
                {
                    if (!_certsLoadedOnce)
                    {
                        await LoadFromServer();
                        _certsLoadedOnce = _serverState == ServerConnectionState.Online;
                    }
                }
                else if (header == "Токены")
                {
                    if (_serverState == ServerConnectionState.Online && _tokenService != null)
                    {
                        await _tokenService.Reload();

                        if (tokensStatusText != null)
                            tokensStatusText.Text = $"Загружено токенов: {_tokenService.Tokens.Count}";
                    }
                }
                else if (header == "Логи")
                {
                    await LoadLogs();
                }
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка при переключении вкладки: " + ex.Message);

                if (header == "Токены" && tokensStatusText != null)
                    tokensStatusText.Text = "Ошибка загрузки токенов";
            }
        }

        private void NoteTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb)
                tb.Tag = tb.Text ?? "";
        }

        private async void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_api == null)
                return;

            if (_isApplyingFromServer)
                return;

            if (!(sender is TextBox tb) || !(tb.DataContext is CertRecord record))
                return;

            string oldNote = tb.Tag as string ?? "";
            string newNote = tb.Text ?? "";

            if (string.Equals(newNote, oldNote, StringComparison.Ordinal))
                return;

            try
            {
                string base64Note = Convert.ToBase64String(Encoding.UTF8.GetBytes(newNote));

                var resp = await _api.UpdateNote(record.Id, base64Note);

                if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                {
                    record.Note = oldNote;
                    tb.Text = oldNote;

                    MessageBox.Show("Ошибка сохранения примечания");
                    return;
                }

                tb.Tag = newNote;
                AddToMiniLog("Примечание сохранено");
            }
            catch (Exception ex)
            {
                record.Note = oldNote;
                tb.Text = oldNote;

                MessageBox.Show("Ошибка сохранения примечания:\n" + ex.Message);
            }
        }
    }
}