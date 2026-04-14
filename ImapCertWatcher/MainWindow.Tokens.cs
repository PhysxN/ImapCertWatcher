using ImapCertWatcher.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private bool _suppressTokenComboTextChanged;
        private bool _isFilteringTokenCombo;
        private async void BtnAddToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            var sn = Prompt.ShowDialog(
                "Введите серийный номер токена",
                "Добавление токена");

            if (string.IsNullOrWhiteSpace(sn))
                return;

            sn = sn.Trim().ToUpperInvariant();

            try
            {
                var response = await _api.AddToken(sn);

                if (string.IsNullOrWhiteSpace(response))
                {
                    MessageBox.Show("Нет ответа от сервера");
                    return;
                }

                response = response.Trim();

                if (response.StartsWith("ERROR|ADD_TOKEN|TOKEN_ALREADY_EXISTS"))
                {
                    MessageBox.Show(
                        "Токен с таким серийным номером уже существует.",
                        "Дубликат",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (response.StartsWith("ERROR|ADD_TOKEN|"))
                {
                    MessageBox.Show(
                        "Ошибка добавления токена:\n" + response.Substring("ERROR|ADD_TOKEN|".Length),
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                if (!response.StartsWith("OK"))
                {
                    MessageBox.Show(
                        "Ошибка добавления токена:\n" + response,
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                await LoadFromServer();
                AddToMiniLog($"Токен {sn} добавлен");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка добавления токена:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                AddToMiniLog("Ошибка добавления токена: " + ex.Message);
            }
        }

        private async void BtnUnassignToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            if (!TryGetToken(sender, out var token))
                return;

            if (token.OwnerCertId == null)
            {
                MessageBox.Show("Токен уже свободен.");
                return;
            }

            if (MessageBox.Show(
                $"Освободить токен {token.Sn}?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            try
            {
                var ok = await _tokenService.Unassign(token.Id);

                if (!ok)
                {
                    MessageBox.Show(
                        "Ошибка освобождения токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                await LoadFromServer();
                AddToMiniLog($"Токен {token.Sn} освобожден");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка освобождения токена:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                AddToMiniLog("Ошибка освобождения токена: " + ex.Message);
            }
        }

        private async void BtnDeleteToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            if (!TryGetToken(sender, out var token))
                return;

            if (token.OwnerCertId != null)
            {
                MessageBox.Show(
                    "Нельзя удалить занятый токен.\nСначала освободите его.",
                    "Токен занят",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show(
                $"Удалить токен {token.Sn}?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            try
            {
                var ok = await _tokenService.Delete(token.Id);

                if (!ok)
                {
                    MessageBox.Show(
                        "Ошибка удаления токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                if (_tokenService.Tokens.Any(t => t.Id == token.Id))
                {
                    MessageBox.Show("Сервер ответил OK, но токен остался в списке.");
                    return;
                }

                await LoadFromServer();
                AddToMiniLog($"Токен {token.Sn} удалён");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка удаления токена:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                AddToMiniLog("Ошибка удаления токена: " + ex.Message);
            }
        }

        private bool TryGetToken(object sender, out TokenRecord token)
        {
            token = (sender as FrameworkElement)?.DataContext as TokenRecord;
            return token != null;
        }

        private void RebuildAvailableTokens()
        {
            if (_isLoadingFromServer || _isApplyingFromServer)
                return;

            void apply()
            {
                var allTokens = (_tokenService?.Tokens?.ToList() ?? new List<TokenRecord>())
                    .Where(t => t != null)
                    .GroupBy(t => t.Id)
                    .Select(g => g.First())
                    .OrderBy(t => t.Sn)
                    .ToList();

                foreach (var cert in _allItems)
                {
                    var actualSelectedToken = cert.TokenId.HasValue
                        ? allTokens.FirstOrDefault(t => t.Id == cert.TokenId.Value)
                        : null;

                    var allowedTokens = allTokens
                        .Where(t => !t.OwnerCertId.HasValue || t.OwnerCertId.Value == cert.Id)
                        .Select(t => new TokenRecord
                        {
                            Id = t.Id,
                            Sn = t.Sn,
                            OwnerCertId = t.OwnerCertId,
                            OwnerFio = t.OwnerFio
                        })
                        .ToList();

                    if (cert.AvailableTokens == null)
                        cert.AvailableTokens = new ObservableCollection<TokenRecord>();
                    else
                        cert.AvailableTokens.Clear();

                    foreach (var token in allowedTokens)
                        cert.AvailableTokens.Add(token);

                    if (actualSelectedToken != null && actualSelectedToken.OwnerCertId == cert.Id)
                    {
                        cert.SelectedToken = cert.AvailableTokens.FirstOrDefault(t => t.Id == actualSelectedToken.Id);
                    }
                    else
                    {
                        cert.SelectedToken = null;
                    }
                }

                dgCerts.Items.Refresh();
            }

            if (Dispatcher.CheckAccess())
                apply();
            else
                Dispatcher.Invoke(apply);
        }

        private static TextBox GetTokenComboEditableTextBox(ComboBox comboBox)
        {
            if (comboBox == null)
                return null;

            comboBox.ApplyTemplate();
            return comboBox.Template.FindName("PART_EditableTextBox", comboBox) as TextBox;
        }

        private void TokenComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb == null)
                return;

            var tb = GetTokenComboEditableTextBox(cb);
            if (tb == null)
                return;

            tb.Tag = cb;
            tb.TextChanged -= TokenComboBoxEditableTextBox_TextChanged;
            tb.TextChanged += TokenComboBoxEditableTextBox_TextChanged;
        }

        private void TokenComboBox_GotKeyboardFocus(object sender, RoutedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb == null || !cb.IsEnabled)
                return;

            _suppressTokenComboTextChanged = true;
            _isFilteringTokenCombo = true;

            try
            {
                ResetTokenComboItems(cb);
            }
            finally
            {
                _isFilteringTokenCombo = false;
                _suppressTokenComboTextChanged = false;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                var tb = GetTokenComboEditableTextBox(cb);
                if (tb == null)
                    return;

                tb.SelectAll();
                cb.IsDropDownOpen = true;
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        private void TokenComboBox_DropDownClosed(object sender, EventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb == null)
                return;

            _suppressTokenComboTextChanged = true;
            _isFilteringTokenCombo = true;

            try
            {
                ResetTokenComboItems(cb);

                var tb = GetTokenComboEditableTextBox(cb);
                if (tb != null)
                {
                    if (cb.DataContext is CertRecord cert)
                        tb.Text = cert.TokenSn ?? "";
                    else
                        tb.Text = "";
                }
            }
            finally
            {
                _isFilteringTokenCombo = false;
                _suppressTokenComboTextChanged = false;
            }
        }

        private void TokenComboBoxEditableTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTokenComboTextChanged)
                return;

            if (_isApplyingFromServer || _isAssigningToken)
                return;

            var tb = sender as TextBox;
            if (tb == null)
                return;

            var cb = tb.Tag as ComboBox;
            if (cb == null)
                return;

            _isFilteringTokenCombo = true;
            try
            {
                ApplyTokenComboFilterSafe(cb, tb.Text);
            }
            finally
            {
                _isFilteringTokenCombo = false;
            }

            if (!cb.IsDropDownOpen)
                cb.IsDropDownOpen = true;

            tb.CaretIndex = tb.Text.Length;
        }

        private static void ResetTokenComboItems(ComboBox cb)
        {
            var cert = cb?.DataContext as CertRecord;
            if (cert == null)
                return;

            cb.ItemsSource = cert.AvailableTokens;
        }

        private static void ApplyTokenComboFilterSafe(ComboBox cb, string filterText)
        {
            var cert = cb?.DataContext as CertRecord;
            if (cert == null)
                return;

            var allTokens = cert.AvailableTokens?.ToList() ?? new List<TokenRecord>();
            var text = (filterText ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(text))
            {
                cb.ItemsSource = cert.AvailableTokens;
                return;
            }

            var filtered = allTokens
                .Where(t => !string.IsNullOrWhiteSpace(t.Sn) &&
                            t.Sn.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            if (cert.TokenId.HasValue && filtered.All(t => t.Id != cert.TokenId.Value))
            {
                var currentToken = allTokens.FirstOrDefault(t => t.Id == cert.TokenId.Value);
                if (currentToken != null)
                    filtered.Insert(0, currentToken);
            }

            cb.ItemsSource = filtered;
        }


        private void CommitCertificatesGridEditSafe()
        {
            try
            {
                if (dgCerts != null)
                {
                    dgCerts.CommitEdit(DataGridEditingUnit.Cell, true);
                    dgCerts.CommitEdit(DataGridEditingUnit.Row, true);
                }
            }
            catch
            {
                // специально ничего не делаем
            }
        }

        private void ReloadFromServerAfterTokenEdit()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _ = LoadFromServer();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private async void TokenComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_api == null)
                return;

            if (_tokenService == null)
                return;

            if (_isApplyingFromServer)
                return;

            if (_isAssigningToken)
                return;

            if (_isFilteringTokenCombo)
                return;

            if (!(sender is ComboBox cb))
                return;

            if (!(cb.DataContext is CertRecord cert))
                return;

            if (cert.IsDeleted || cert.IsRevoked)
                return;

            if (e.AddedItems == null || e.AddedItems.Count == 0)
                return;

            var token = e.AddedItems[0] as TokenRecord;
            if (token == null)
                return;

            var previousToken = e.RemovedItems != null && e.RemovedItems.Count > 0
                ? e.RemovedItems[0] as TokenRecord
                : null;

            if (previousToken != null && previousToken.Id == token.Id)
                return;

            if (!cb.IsKeyboardFocusWithin && !cb.IsDropDownOpen)
                return;

            _isAssigningToken = true;

            try
            {
                CommitCertificatesGridEditSafe();

                var ok = await _tokenService.Assign(cert.Id, token.Id);

                if (!ok)
                {
                    MessageBox.Show(
                        "Ошибка назначения токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    ReloadFromServerAfterTokenEdit();
                    return;
                }

                cert.TokenId = token.Id;
                cert.TokenSn = token.Sn;
                cert.SelectedToken = cert.AvailableTokens?.FirstOrDefault(t => t.Id == token.Id);

                AddToMiniLog($"Назначен токен {token.Sn} → {cert.Fio}");

                ReloadFromServerAfterTokenEdit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка назначения токена:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                ReloadFromServerAfterTokenEdit();
            }
            finally
            {
                _isAssigningToken = false;
            }
        }
    }
}