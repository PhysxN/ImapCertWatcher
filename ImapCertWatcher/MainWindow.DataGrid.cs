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

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
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
                        foreach (var column in dgCerts.Columns)
                        {
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

        private async Task LoadFromServer()
        {
            if (_isLoadingFromServer)
                return;

            _isLoadingFromServer = true;
            SetServerState(ServerConnectionState.Connecting);

            try
            {
                if (_api == null)
                {
                    AddToMiniLog("API ещё не инициализирован");
                    SetServerState(ServerConnectionState.Offline);
                    statusText.Text = "Ошибка: API не инициализирован";
                    return;
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
                    }
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    _isApplyingFromServer = true;

                    try
                    {
                        _allItems.Clear();

                        foreach (var item in list)
                            _allItems.Add(item);

                        RebuildAvailableTokens();
                        ApplySearchFilter();

                        statusText.Text = $"Загружено записей: {_allItems.Count}";
                        SetServerState(ServerConnectionState.Online);
                    }
                    finally
                    {
                        _isApplyingFromServer = false;
                    }
                });

                AddToMiniLog($"Получено с сервера: {list.Count} записей");
            }
            catch (Exception ex)
            {
                AddToMiniLog("Ошибка загрузки: " + ex.Message);
                SetServerState(ServerConnectionState.Offline);
                statusText.Text = "Ошибка загрузки с сервера: " + ex.Message;

                MessageBox.Show(
                    "Не удалось получить список сертификатов с сервера.\n\n" + ex.Message,
                    "GET_CERTS",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _isLoadingFromServer = false;
            }
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
            AddToMiniLog(_showDeleted ? "Показаны удаленные" : "Скрыты удаленные");
        }

        private async void BuildingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
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

            if (!(sender is ComboBox comboBox))
                return;

            if (!(comboBox.DataContext is CertRecord record))
                return;

            if (e.AddedItems.Count == 0)
                return;

            string newValue = e.AddedItems[0] as string ?? "";
            string oldValue = e.RemovedItems.Count > 0
                ? e.RemovedItems[0] as string ?? ""
                : record.Building ?? "";

            if (string.Equals(newValue, oldValue, StringComparison.Ordinal))
                return;

            _isSavingBuilding = true;

            try
            {
                var resp = await _api.UpdateBuilding(record.Id, newValue);

                if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                {
                    record.Building = oldValue;
                    CollectionViewSource.GetDefaultView(comboBox.ItemsSource)?.Refresh();
                    MessageBox.Show("Ошибка сохранения здания");
                    return;
                }

                AddToMiniLog(
                    $"Здание изменено: {record.Fio} → {(string.IsNullOrWhiteSpace(newValue) ? "<пусто>" : newValue)}");

                statusText.Text = "Здание сохранено";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения здания:\n" + ex.Message);

                record.Building = oldValue;
                CollectionViewSource.GetDefaultView(comboBox.ItemsSource)?.Refresh();
            }
            finally
            {
                _isSavingBuilding = false;
            }
        }

        private void DgCerts_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
        }

        private void DgCerts_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
        }

        private void DgCerts_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
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

        private async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(sender is TabControl tc))
                return;

            if (!(tc.SelectedItem is TabItem tab))
                return;

            var header = tab.Header?.ToString();

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
                if (_tokenService != null && _tokenService.Tokens.Count == 0)
                    await _tokenService.Reload();
            }
        }

        private async void NoteTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_api == null)
                return;

            if (sender is TextBox tb && tb.DataContext is CertRecord record)
            {
                try
                {
                    string note = tb.Text ?? "";
                    string base64Note = Convert.ToBase64String(Encoding.UTF8.GetBytes(note));

                    var resp = await _api.UpdateNote(record.Id, base64Note);

                    if (string.IsNullOrEmpty(resp) || !resp.StartsWith("OK"))
                    {
                        MessageBox.Show("Ошибка сохранения примечания");
                        return;
                    }

                    AddToMiniLog("Примечание сохранено");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка сохранения примечания:\n" + ex.Message);
                }
            }
        }
    }
}